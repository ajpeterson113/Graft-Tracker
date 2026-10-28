/**
 * Graft Inbox
 * -----------
 * Forward an email to ajpeterson113+grafts@gmail.com. Within 5 minutes this
 * script copies it (the text plus its attachments) into the private
 * Graft-Tracker repo at inbox/email/<date>-<subject>/. Claude files what's
 * there at the morning update, or whenever AJ says "check the inbox".
 *
 * In Gmail, a copied email gets the "Graft Inbox" label and is archived.
 * Only mail from AJ's own addresses is copied: his Google account, plus the
 * ALLOWED_FROM script property (default: anything @smith-nephew.com).
 * Anything else gets the "Graft Inbox Blocked" label, stays where it is, and
 * AJ gets one email about that sender.
 *
 * Setup and troubleshooting: tools/graft_inbox/README.md in the repo.
 * Functions to run by hand: setup (once), stop (turns it off).
 */

const CONFIG = {
  address: 'ajpeterson113+grafts@gmail.com',
  repo: 'ajpeterson113/Graft-Tracker',
  branch: 'main',
  folder: 'inbox/email',
  label: 'Graft Inbox',                // copied to GitHub
  blockedLabel: 'Graft Inbox Blocked', // not from an allowed sender
  defaultAllowed: '@smith-nephew.com', // used when ALLOWED_FROM is empty
  timeZone: 'America/Chicago',
  maxAttachmentBytes: 25 * 1024 * 1024,
  minInlineImageBytes: 20 * 1024,      // smaller inline images are logos/signatures
  maxThreadsPerRun: 25,
  maxRunMs: 4.5 * 60 * 1000,           // Apps Script stops a run at 6 minutes
  keepDoneIds: 250,
};

/** Run once by hand: checks the token, makes the labels, starts the 5-minute timer. */
function setup() {
  token_(); // throws a clear error if GITHUB_TOKEN is missing
  // Creating a loose blob proves the token can write to the repo. It touches no branch.
  gh_('post', '/repos/' + CONFIG.repo + '/git/blobs', {
    content: Utilities.base64Encode('Graft Inbox check', Utilities.Charset.UTF_8),
    encoding: 'base64',
  });
  label_(CONFIG.label);
  label_(CONFIG.blockedLabel);
  stop();
  ScriptApp.newTrigger('syncGraftInbox').timeBased().everyMinutes(5).create();
  syncGraftInbox();
  Logger.log('Graft Inbox is on. Forward emails to ' + CONFIG.address + '.');
}

/** Run by hand to turn the inbox off. */
function stop() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'syncGraftInbox'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

/** Runs every 5 minutes. */
function syncGraftInbox() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // the previous run is still going
  try {
    run_();
  } finally {
    lock.releaseLock();
  }
}

function run_() {
  const started = Date.now();
  const props = PropertiesService.getScriptProperties();
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  const addr = CONFIG.address.toLowerCase();
  const allowed = (props.getProperty('ALLOWED_FROM') || CONFIG.defaultAllowed)
    .split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String);
  const done = JSON.parse(props.getProperty('DONE_IDS') || '[]');
  const saveDone = function (id) {
    done.push(id);
    props.setProperty('DONE_IDS', JSON.stringify(done.slice(-CONFIG.keepDoneIds)));
  };
  const copiedLabel = label_(CONFIG.label);
  const blockedLabel = label_(CONFIG.blockedLabel);
  const query = '{to:' + addr + ' deliveredto:' + addr + '}' +
    ' -label:' + searchName_(CONFIG.label) +
    ' -label:' + searchName_(CONFIG.blockedLabel) +
    ' newer_than:30d';
  const threads = GmailApp.search(query, 0, CONFIG.maxThreadsPerRun);
  const errors = [];

  for (let i = 0; i < threads.length; i++) {
    if (Date.now() - started > CONFIG.maxRunMs) break; // the next run picks up the rest
    const thread = threads[i];
    const msgs = thread.getMessages();
    let ours = 0, copied = 0, blocked = 0, failed = 0, onlyOurs = true;
    for (let j = 0; j < msgs.length; j++) {
      const m = msgs[j];
      if (!isForInbox_(m, addr)) { onlyOurs = false; continue; }
      if (m.isInTrash()) continue;
      ours++;
      if (done.indexOf(m.getId()) >= 0) continue;
      const from = addressOf_(m.getFrom());
      if (from !== me && !isAllowed_(from, allowed)) {
        blocked++;
        saveDone(m.getId());
        notifyOnce_(props, 'BLOCKED ' + from, 24 * 365,
          'Graft Inbox ignored an email from ' + from,
          'Graft Inbox only copies email from your own addresses, so it left this one in Gmail ' +
          'with the "' + CONFIG.blockedLabel + '" label:\n\n' +
          '  From: ' + m.getFrom() + '\n  Subject: ' + m.getSubject() + '\n\n' +
          'If that address is yours, add it to the ALLOWED_FROM script property ' +
          '(Apps Script > Graft Inbox > Project Settings > Script properties; separate ' +
          'addresses with commas, or use @domain.com for a whole domain), then forward the email again.');
        continue;
      }
      try {
        copyMessage_(m);
        copied++;
        saveDone(m.getId());
      } catch (e) {
        failed++;
        errors.push((m.getSubject() || '(no subject)') + ': ' + e.message);
      }
    }
    if (!ours || failed) continue; // nothing for the inbox here, or retry the failures next run
    if (blocked) thread.addLabel(blockedLabel);
    if (copied || !blocked) {
      thread.addLabel(copiedLabel);
      if (onlyOurs && !blocked) thread.moveToArchive();
    }
  }

  if (errors.length) {
    notifyOnce_(props, 'ERROR', 12,
      'Graft Inbox couldn\'t copy an email to GitHub',
      'These emails are still in Gmail. Graft Inbox tries them again every 5 minutes:\n\n  ' +
      errors.join('\n  ') + '\n\n' +
      'A 401 or 403 from GitHub usually means the token expired or lost access. Make a new ' +
      'token and paste it into the GITHUB_TOKEN script property ' +
      '(Apps Script > Graft Inbox > Project Settings > Script properties).');
    Logger.log('Errors:\n' + errors.join('\n'));
  }
}

/** Copies one email into inbox/email/<folder>/ as a single commit. */
function copyMessage_(m) {
  const subject = m.getSubject() || '(no subject)';
  const stamp = Utilities.formatDate(m.getDate(), CONFIG.timeZone, 'yyyy-MM-dd-HHmm');
  const folder = CONFIG.folder + '/' + stamp + '-' + slug_(subject) + '-' + m.getId().slice(-6);
  const taken = { 'email.md': true, 'email.html': true };
  const files = [];
  const listed = [];
  const skipped = [];

  const attachments = m.getAttachments({ includeInlineImages: false, includeAttachments: true });
  const inline = m.getAttachments({ includeInlineImages: true, includeAttachments: false })
    .filter(function (a) { return a.getSize() >= CONFIG.minInlineImageBytes; });
  attachments.concat(inline).forEach(function (a) {
    const size = a.getSize();
    const original = a.getName() || 'attachment';
    if (size > CONFIG.maxAttachmentBytes) {
      skipped.push(original + ' (' + size_(size) + ', too big to copy)');
      return;
    }
    const name = uniqueName_(cleanName_(original), taken);
    files.push({ path: folder + '/' + name, b64: Utilities.base64Encode(a.getBytes()) });
    listed.push(name + ' (' + size_(size) + ')');
  });

  const html = m.getBody() || '';
  if (html && html.length <= 2 * 1024 * 1024) {
    files.push({ path: folder + '/email.html', b64: Utilities.base64Encode(html, Utilities.Charset.UTF_8) });
  }
  files.unshift({
    path: folder + '/email.md',
    b64: Utilities.base64Encode(emailMd_(m, listed, skipped), Utilities.Charset.UTF_8),
  });
  commit_(files, ('Graft Inbox: ' + subject).slice(0, 120));
}

function emailMd_(m, listed, skipped) {
  const cell = function (s) { return String(s || '').replace(/\|/g, '\\|').replace(/\s*\r?\n\s*/g, ' '); };
  const rows = [
    ['Forwarded', Utilities.formatDate(m.getDate(), CONFIG.timeZone, 'EEE M/d/yyyy h:mm a')],
    ['From', m.getFrom()],
    ['To', m.getTo()],
  ];
  if (m.getCc()) rows.push(['Cc', m.getCc()]);
  rows.push(['Attachments', listed.length ? listed.join(', ') : 'none']);
  if (skipped.length) rows.push(['Not copied', skipped.join(', ')]);
  let body = (m.getPlainBody() || '').replace(/\r\n/g, '\n').trim();
  if (body.length > 200000) body = body.slice(0, 200000) + '\n[cut here: the rest is in email.html]';
  return '# ' + cell(m.getSubject() || '(no subject)') + '\n\n' +
    '| | |\n|---|---|\n' +
    rows.map(function (r) { return '| ' + r[0] + ' | ' + cell(r[1]) + ' |'; }).join('\n') + '\n\n' +
    'Gmail message id: ' + m.getId() + '\n\n' +
    '~~~~text\n' + body.replace(/~~~~/g, '~ ~ ~ ~') + '\n~~~~\n';
}

/** One commit on CONFIG.branch with all the files; retries if the branch moved meanwhile. */
function commit_(files, message) {
  const repo = '/repos/' + CONFIG.repo;
  const tree = files.map(function (f) {
    const blob = gh_('post', repo + '/git/blobs', { content: f.b64, encoding: 'base64' });
    return { path: f.path, mode: '100644', type: 'blob', sha: blob.sha };
  });
  for (let attempt = 1; attempt <= 4; attempt++) {
    const head = gh_('get', repo + '/git/ref/heads/' + CONFIG.branch).object.sha;
    const base = gh_('get', repo + '/git/commits/' + head).tree.sha;
    const newTree = gh_('post', repo + '/git/trees', { base_tree: base, tree: tree });
    const commit = gh_('post', repo + '/git/commits', { message: message, tree: newTree.sha, parents: [head] });
    try {
      gh_('patch', repo + '/git/refs/heads/' + CONFIG.branch, { sha: commit.sha, force: false });
      return commit.sha;
    } catch (e) {
      if (e.status !== 422 && e.status !== 409) throw e; // 422: the branch moved; build on the new head
      Utilities.sleep(1000 * attempt);
    }
  }
  throw new Error('the GitHub branch kept changing; trying again next run');
}

function gh_(method, path, body) {
  const options = {
    method: method,
    headers: {
      Authorization: 'Bearer ' + token_(),
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    muteHttpExceptions: true,
  };
  if (body) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(body);
  }
  const res = UrlFetchApp.fetch('https://api.github.com' + path, options);
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) {
    const err = new Error('GitHub ' + code + ' on ' + method.toUpperCase() + ' ' + path + ': ' + text.slice(0, 200));
    err.status = code;
    throw err;
  }
  return text ? JSON.parse(text) : {};
}

function token_() {
  const t = (PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN') || '').trim();
  if (!t) throw new Error('Add your GitHub token as the script property GITHUB_TOKEN (Project Settings > Script properties).');
  return t;
}

function label_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/** Gmail search writes a label's spaces and slashes as hyphens. */
function searchName_(name) {
  return name.toLowerCase().replace(/[\s\/]+/g, '-');
}

function isForInbox_(m, addr) {
  return [m.getTo(), m.getCc(), m.getBcc(), m.getHeader('Delivered-To'), m.getHeader('X-Original-To')]
    .join(' ').toLowerCase().indexOf(addr) >= 0;
}

function addressOf_(from) {
  const s = String(from || '');
  const angle = s.match(/<([^>]+)>/);
  return (angle ? angle[1] : s).trim().toLowerCase();
}

function isAllowed_(from, allowed) {
  return allowed.some(function (a) {
    return a.charAt(0) === '@' ? from.slice(-a.length) === a : from === a;
  });
}

function notifyOnce_(props, key, hours, subject, body) {
  const last = Number(props.getProperty('NOTICE ' + key) || 0);
  if (Date.now() - last < hours * 3600 * 1000) return;
  GmailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, body);
  props.setProperty('NOTICE ' + key, String(Date.now()));
}

function slug_(subject) {
  return subject
    .replace(/^\s*((fwd?|re)\s*:\s*)+/i, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 50)
    .replace(/-+$/, '') || 'email';
}

function cleanName_(name) {
  const clean = name.replace(/[^\w.\- ]+/g, '_').replace(/\s+/g, '_').replace(/^[._]+/, '');
  if (clean.length <= 80) return clean || 'attachment';
  const dot = clean.lastIndexOf('.');
  const ext = dot > 0 && clean.length - dot <= 8 ? clean.slice(dot) : '';
  return clean.slice(0, 80 - ext.length) + ext;
}

function uniqueName_(name, taken) {
  let candidate = name;
  const dot = name.lastIndexOf('.');
  const stem = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : '';
  for (let n = 2; taken[candidate.toLowerCase()]; n++) candidate = stem + '-' + n + ext;
  taken[candidate.toLowerCase()] = true;
  return candidate;
}

function size_(bytes) {
  return bytes >= 1024 * 1024 ? (bytes / 1024 / 1024).toFixed(1) + ' MB' : Math.max(1, Math.round(bytes / 1024)) + ' KB';
}
