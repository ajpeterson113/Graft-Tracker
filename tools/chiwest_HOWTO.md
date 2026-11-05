# Chi West app: keeping it current

Chi West (https://claude.ai/artifact/HzjrjSVDB4oF6oAfWUuwDN) is AJ's combined app (Today · Patients · Tracker · Grow). He started using it on 11/4 alongside the Graft Tracker, Chi West Patients and Chi West Targets apps, which stay the source of truth until he says he has fully switched.

It reads six documents in its db collection `data`, all written by `tools/chiwest_sync.py`:
`core` (period, sales, forecast, accounts, year, open action items), `cases` (last three month tabs), `forecast` (projection rows), `stock` (shelves), `patients` (patient list, visits, application log, day notes — initials only, from Chi_West_Patient_List.xlsx) and `targets` (target list + Targets-app outreach).
A seventh, `data/research`, is written by Claude only: `{briefs: {<target id>: {date, items: [..]}}, styles: {<target id>: {style: data|rel|eff|skep, why}}}`. The sync never overwrites it.

## Every refresh (after the Graft Tracker app, Patients app and repo are updated)

1. Download into new, empty folders (ArtifactData list, limit 1000, out_dir):
   - Graft Tracker app (https://claude.ai/artifact/QygXA2j9AUKiYq42UimDBn): `meta`, `months`, `todos` → GTDB (or reuse the run's DB + OUT: OUT/meta, OUT/months and DB/todos).
   - Chi West Targets (https://claude.ai/artifact/JCPaNuwhst38msAxKippjo): `outreach` → TGT.
   - Chi West: `data` → CW (note each version) and `actions`, `uploads`.
2. `python3 -I tools/chiwest_sync.py --gt GTDB --patients Chi_West_Patient_List.xlsx --targets Chi_West_Surgeon_Targets.xlsx --outreach TGT/outreach --cw CW OUT_CW`
   It prints each doc's size (keep every doc under 240 KB) and writes OUT_CW/batches.json. Replace any "VERSION" with that doc's version from step 1 and send each batch with ArtifactData `batch` on the Chi West url. `core` goes last.

## AJ's taps in Chi West (apply them in the same run, before step 2)

`actions` docs with status "new". Apply each, then update it to status "applied" (pinned with if_version). If one can't be applied confidently, leave it "new" and list it.
- `sfdc` {caseId, to: "Yes" | "In SFDC" | "No"}: same as a Graft Tracker Submitted / In SFDC tap on that case id.
- `grafted` {pid, ini, acct, product, size, qty, app, date, status, note}: log the visit on that patient in Chi West Patients (log entry, apps, status, last) and add the tracker row as waiting on a PO, exactly as a Grafted tap in the Patients app.
- `notgrafted` {pid, date, reason, status, note}: log it in Chi West Patients.
- `status` {pid, status, cover, next, note}: update the patient in Chi West Patients.
- `note` {pid | caseId, text}: dated note on the patient, or the tracker row's Notes / right column.
- `newpatient` {acct, prov, ini, product, size, freq, status, first, note}: add to Chi West Patients (Pending unless status says otherwise).
- `todo` {todo}: set that Graft Tracker `todos` item to done (doneAt now).
- `touch` / `lunch` {target, who, type, approach, newStatus, next, note, edate, etime}: add a history entry and status / followUp on the Targets app `outreach` doc for that target (lunch: also the Lunch booked entry). Keep the action's `approach` — it is how Chi West learns.
- `objection` {target, text}: the target's objection field in the Targets app.
- `style` {target, style}: `data/research.styles[target]` in Chi West.
- `research` {target}: research that target on the web (practice, offices, affiliations, OR/clinic days, publications, who they share a building with, competitive product, claims signal from the target list) and write 4–8 short findings to `data/research.briefs[target]` with today's date. No patient data.
- `event` {title, ekind, place, edate, etime, who}: add to the Chi West Patients `events` collection.

`uploads` docs with status "new" (kinds case_sheet, stickers, po, stock, note, sfdc): read the asset (Artifact read, url = Chi West, path = asset id) and file it like any other upload, then set status "filed".
