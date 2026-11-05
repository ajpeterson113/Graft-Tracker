# Chi West app: keeping it current

Chi West (https://claude.ai/artifact/HzjrjSVDB4oF6oAfWUuwDN) is AJ's combined app (Home · Patients · Tracker · Grow). The Patients tab has the coverage planner (copied as-is from the Chi West Patients app on 11/4; Plan coverage on the schedule, the Coverage card and Home's To do; Apply queues one `cover` action per visit with that visit's date), coverage email and Read with Claude (coverage reply, clinic schedule, calendar); Home and Tracker have the projection email. He started using it on 11/4 alongside the Graft Tracker, Chi West Patients and Chi West Targets apps, which stay the source of truth until he says he has fully switched.

It reads six documents in its db collection `data`, all written by `tools/chiwest_sync.py`:
`core` (period, sales, forecast, accounts, year, open action items), `cases` (last three month tabs), `forecast` (projection rows), `stock` (shelves), `patients` (patient list, visits, application log, day notes from Chi_West_Patient_List.xlsx, plus names, who covers each visit, next-visit location, verify flags and AJ's events from the Chi West Patients db) and `targets` (target list + Targets-app outreach). AJ wants names (shown on tap, as in the Patients app) and coverage in Chi West (11/4).
A seventh, `data/research`, is written by Claude only: `{briefs: {<target id>: {date, items: [..]}}, styles: {<target id>: {style: data|rel|eff|skep, why}}}`. The sync never overwrites it.

## Every refresh (after the Graft Tracker app, Patients app and repo are updated)

1. Download into new, empty folders (ArtifactData list, limit 1000, out_dir):
   - Graft Tracker app (https://claude.ai/artifact/QygXA2j9AUKiYq42UimDBn): `meta`, `months`, `todos` → GTDB (or reuse the run's DB + OUT: OUT/meta, OUT/months and DB/todos).
   - Chi West Targets (https://claude.ai/artifact/JCPaNuwhst38msAxKippjo): `outreach` → TGT.
   - Chi West Patients (https://claude.ai/artifact/AwS5M2wc9Wm7Pe6oTuJuwQ): `patients` and `events` → PDB. If ArtifactData refuses out_dir for these (it can, because AJ writes them through the app), list them without out_dir and write PDB/overlay.json yourself: {"patients": [{id, name, cover, status, apps, next, nextTime, nextLoc, verify, fuVerify, ivrDate, paExp} for every patient], "events": [{id, title, kind, place, date, time, end, note}]}. Names stay exactly as in the db (drop an insurance or DOB tacked onto a name).
   - Chi West: `data` → CW (note each version) and `actions`, `uploads`.
2. `python3 -I tools/chiwest_sync.py --gt GTDB --patients Chi_West_Patient_List.xlsx --pdb PDB --targets Chi_West_Surgeon_Targets.xlsx --outreach TGT/outreach --cw CW OUT_CW` (or `--overlay PDB/overlay.json` instead of `--pdb PDB`). Never run it without one of the two: the app would lose names and coverage.
   It prints each doc's size (keep every doc under 240 KB) and writes OUT_CW/batches.json. Replace any "VERSION" with that doc's version from step 1 and send each batch with ArtifactData `batch` on the Chi West url. `core` goes last.

## AJ's taps in Chi West (apply them in the same run, before step 2)

`actions` docs with status "new". Apply each, then update it to status "applied" (pinned with if_version). If one can't be applied confidently, leave it "new" and list it.
`status` on an action is only this queue flag. A patient's category travels as `pstatus`, and only when AJ picked a different one in the form.

**AJ's rule (11/4): never assume a change in a patient's category (grafting status) or graft size.** The only category suggestions are made in the app forms, where AJ sees them before saving (Chi West and the Chi West Patients app, 11/4): Healed (Not grafted) → Done, A8 at Alexian WCC → Done, and a new start's first graft → Grafting. They arrive as `pstatus` like any other pick. Ask AJ before adding any other. Change `status` only from a `pstatus`, and the patient's `size` only from a `status` action's `size` or an `edit` action's `changes.size`. A logged visit records the size used on its log entry and tracker row and changes nothing else on the patient. Never infer Done, On hold, Grafting or a size step-down from a reason ("Healed", "Infection"), a note, a schedule, a calendar or a projection rule.
- `sfdc` {caseId, to: "Yes" | "In SFDC" | "No"}: same as a Graft Tracker Submitted / In SFDC tap on that case id.
- `grafted` {pid, ini, acct, product, size, qty, app, date, pstatus?, note}: log the visit on that patient in Chi West Patients (log entry with the size used, apps, last; `status` only if `pstatus`) and add the tracker row as waiting on a PO, exactly as a Grafted tap in the Patients app.
- `notgrafted` {pid, date, reason, pstatus?, note}: log it in Chi West Patients. The reason is just the log note; the category stays unless `pstatus` is set.
- `status` {pid, pstatus?, size?, cover, next, note}: update the patient in Chi West Patients: `status` = pstatus and `size` = size only when present (AJ changed them in the Status form).
- `cover` {pid, date, cover: "" | "Caroline" | "Pete" | "Not covered"}: the patient's `cover` field in Chi West Patients (the schedule's coverage dropdown, the coverage planner's Apply, or a coverage reply read by Claude). "" means AJ. It applies to the patient's next visit; skip it (leave "new" and list it) if `date` is no longer that patient's `next`.
- `next` {pid, date, time, loc}: move the patient's next visit (from a clinic schedule or AJ's calendar read by Claude): `next`, `nextTime` (if given), `nextLoc` (if given); clear `fuVerify`; clear `cover` when the date changed. Never change `status` for it.
- `verify` {pid, text, confirmed?}: set the patient's `verify` field to text; empty text clears `verify` and `fuVerify`. With `confirmed` {date, time, app}, AJ tapped Confirm on the schedule row: the visit is right as shown (that date, time and application number), so clear both flags and make sure `next`/`nextTime`/`apps` match it.
- `edit` {pid, pstatus?, changes, was}: AJ's Edit form on a patient (the Patients app's Edit fields). `changes` holds only the fields he changed, under the Chi West Patients field names (initials, name, account, provider, product, size, freq, apps, coverage, pa, ivrDate, paExp, caseDate, cover, next, nextTime, nextLoc, notes); `was` holds what Chi West showed before. Set each changed field on that patient in Chi West Patients, `status` = pstatus when present; if `next` changed, clear `fuVerify` (as the Patients app's Edit does). If a field's current value in the Patients app no longer matches `was` (changed there since), don't overwrite that field: apply the rest, mark the action "applied" and list the skipped field for AJ.
- `time` {pid, date, time}: the visit on that date moved to that time — Chi West Patients `nextTime` (clear `fuVerify`), same as the Patients app's time picker.
- `appno` {pid, app}: the next visit is application #app — Chi West Patients `apps` = app − 1, same as the Patients app's A# picker.
- `note` {pid | caseId, text}: dated note on the patient, or the tracker row's Notes / right column.
- `newpatient` {acct, prov, ini, product, size, freq, pstatus, first, time, loc, note}: add to Chi West Patients (Pending unless pstatus says otherwise; `first`/`time`/`loc` become next/nextTime/nextLoc).

AJ 11/4: never change a patient's category (status) unless his own tap says so (`status` actions, or a Grafted/Not grafted tap that carries a status). Clinic schedules, calendars and coverage replies only move visits, set coverage, add notes and set verify flags.
- `todo` {todo}: set that Graft Tracker `todos` item to done (doneAt now).
- `touch` / `lunch` {target, who, type, approach, newStatus, next, note, edate, etime}: add a history entry and status / followUp on the Targets app `outreach` doc for that target (lunch: also the Lunch booked entry). Keep the action's `approach` — it is how Chi West learns.
- `objection` {target, text}: the target's objection field in the Targets app.
- `style` {target, style}: `data/research.styles[target]` in Chi West.
- `research` {target}: research that target on the web (practice, offices, affiliations, OR/clinic days, publications, who they share a building with, competitive product, claims signal from the target list) and write 4–8 short findings to `data/research.briefs[target]` with today's date. No patient data.
- `event` {title, ekind, place, edate, etime, who}: add to the Chi West Patients `events` collection.

`uploads` docs with status "new" (kinds case_sheet, stickers, po, stock, note, sfdc): read the asset (Artifact read, url = Chi West, path = asset id) and file it like any other upload, then set status "filed".
