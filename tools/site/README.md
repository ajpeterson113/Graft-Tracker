# Page generator

`build_site.py` writes this repo's web view from the same data the two apps show, so the pages work like the apps: one tab bar on every page (Home · To do · Cases · Projection · Stock · Patients · Targets), a Home with the period pace and what needs doing, and the detail one click away.

## When to run it

**Every update, no exceptions (AJ 11/5):** the home page (README.md and view/home*.png) must show the latest numbers after every tracker update. If the Patients db can't be exported, use Chi West's data/patients.json (below) rather than skipping the run.

After every tracker update, once `graft_app.py build` has run and the Chi West Patients db has been exported:

```
python3 -I tools/site/build_site.py --app OUT --patients PDB --repo .
```

| Option | What it is |
| --- | --- |
| `--app OUT` | The output folder of `graft_app.py build` (meta/snap.json, meta/forecast.json, meta/stock.json, months/*.json) |
| `--patients PDB` | The Chi West Patients db, exported with ArtifactData `list` and `out_dir=PDB` for the collections `patients`, `inbox` and `uploads`. When that export is refused (ArtifactData won't save docs AJ wrote in the app to files), pass Chi West's `data/patients.json` instead: the `data` collection of Chi West (https://claude.ai/artifact/HzjrjSVDB4oF6oAfWUuwDN) exported with `out_dir`, or the `chiwest_sync.py` output. Same pages, minus the inbox/upload counts. |
| `--xlsx FILE` | Optional. The master tracker, read for the Transfers tab (default: Graft_Tracker_BIOT_Chicago_West.xlsx in the repo) |
| `--now 2026-10-12T22:45` | Optional. The Chicago date and time that counts as now (default: the clock) |

Then commit everything it wrote along with the tracker and the patient list Excel.

## What it writes

- README.md (Home) and view/home.png, home-dark.png, home-narrow.png, home-narrow-dark.png (the pace picture for desktop and phone, light and dark)
- tracker/todo.md, tracker/cases.md, tracker/projection.md, tracker/stock.md
- Patients.md, patients/week.md, patients/todo.md and one page per status (patients/grafting.md, pending.md, …)
- the tab bar at the top of Prospects.md, between the `tabs:start` and `tabs:end` markers (the rest of that page is left alone)

## Rules

- Don't hand-edit the generated pages: the next update overwrites them. Change the script.
- Patient pages carry no sales numbers (same rule as the Patients app).
- **No patient names while the repo is public (11/9).** The script writes initials only unless the file `tools/site/NAMES_OK` exists. Add that file only after AJ has made the repo private (GitHub › Settings › Danger Zone › Change visibility); then full names come back behind a click, as before. Never add it while the repo is public, and never work around it.
- Needs Python 3.9+, Pillow and openpyxl. The fonts in fonts/ (Barlow Semi Condensed, Source Sans 3) are under the SIL Open Font License.
