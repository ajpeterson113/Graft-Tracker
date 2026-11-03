# BIOT Chicago West Graft Tracker — System Overview (for Microsoft Copilot)

**Owner:** AJ Peterson, Account Manager, Smith+Nephew Bio Tissue (BIOT Chicago West)
**Products:** GRAFIX (incl. Grafix PL Prime — always labeled "GRAFIX"), STRAVIX, OASIS. Axolotl is discontinued (2025 only).
**Purpose:** Track every graft applied in the territory from application → case sheet → purchase order → SFDC submission, forecast each sales period, manage the patient roster behind those sales, and prospect new surgeons.

This document describes the whole system so Copilot can read, update or reason over the files without further context. Everything lives in two places: a private GitHub repo (`ajpeterson113/Graft-Tracker`) that the work laptop can open, and a set of Claude-hosted phone apps. The GitHub repo is the source of truth that Copilot can see.

---

## 1. The files (GitHub repo `ajpeterson113/Graft-Tracker`)

| File | What it is | Built by |
|---|---|---|
| `Graft_Tracker_BIOT_Chicago_West.xlsx` | **Master graft tracker** — all months, all accounts, projections, stock. Fixed path, overwritten after every update (date lives inside the file name on the chat copy only). | Claude (chat + scheduled task) |
| `Chi_West_Patient_List.xlsx` | Patient roster export (one tab per account + All Patients, Follow-ups, Application Log). | `tools/patient_list.py` from the Chi West Patients app |
| `Chi_West_Surgeon_Targets.xlsx` | 181 surgeon/clinician prospects by hospital, with claims data, contacts, and what changed. | Chi West Targets app |
| `README.md` | Laptop home page: period pace, To-do counts, this week's visits, P-period forecast, full year, download links. **Generated — never hand-edit.** | `tools/site/build_site.py` |
| `tracker/todo.md`, `cases.md`, `projection.md`, `stock.md` | Read-only web views of the tracker tabs. Generated. | `tools/site/build_site.py` |
| `Patients.md`, `patients/*.md` | Read-only patient list by account/status + this week's schedule. Generated. | `tools/site/build_site.py` |
| `Prospects.md` | Read-only target list. Generated. | Targets app build |
| `case-sheets/<initials>-<mmdd>.{jpg,pdf,md}` | One case sheet (implant record) per application. The `.md` page carries date, patient initials, account, surgeon, site, product, REF, lot/SN, expiry, price, PO and SFDC status, plus the image. Tracker "Case Sheet" column links here. | `tools/case_sheet.py` |
| `ivr/` | IVR (insurance verification) screenshots + summary pages per patient. | Claude |
| `inbox/email/`, `inbox/new-patients/`, `inbox/associate-notes/` | Drop folders: emails forwarded to ajpeterson113+grafts@gmail.com (Apps Script in `tools/graft_inbox`), new-patient and associate-note submissions from the laptop page. Filed, then deleted. | Gmail Apps Script / GitHub page buttons |
| `tools/` | `case_sheet.py` (build a case sheet from sticker photos), `patient_list.py`, `clinic_todos.py`, `site/build_site.py` (web pages). | — |

Download links (bookmarked in Edge on the work laptop):
- Tracker: `https://github.com/ajpeterson113/Graft-Tracker/raw/main/Graft_Tracker_BIOT_Chicago_West.xlsx`
- Patient list: `https://github.com/ajpeterson113/Graft-Tracker/raw/main/Chi_West_Patient_List.xlsx`
- Targets: `https://github.com/ajpeterson113/Graft-Tracker/raw/main/Chi_West_Surgeon_Targets.xlsx`

**Sync is one-way: Claude → GitHub.** Edits made on the laptop are NOT pushed back. AJ tells the system in chat when something has been submitted.

---

## 2. The master tracker workbook (`Graft_Tracker_BIOT_Chicago_West.xlsx`)

### Tabs
| Tab | Contents |
|---|---|
| `Dashboard` | Quarterly sales/commission summary, average graft sale, commission plan tiers, instructions. |
| `Prices` | Contract price grid per product/size (VA grid + account rates). |
| `Returns` | Return log (expired / scrap / credit). |
| `Transfers` | Consignment units moved between accounts (unit leaves "from" count, billed to "to" account). A GRAFIX/STRAVIX Transport Log PDF is filed for each. |
| `XC` | Grafix XC application log. |
| `January` … `December` | **Application log, one tab per month.** Each tab opens with a period snapshot (Sales, Quota, % to quota, Commission rate, Commission, Cases) and then the log. |
| `Analytics` | Cross-month analysis. |
| `Projections` / `P11 Projections` | Current-period forecast by account and patient, Upside vs Downside, WCC stock offset, potential new starts. |
| `Stock`, `St Alexius Stock`, `Good Sam OR Stock`, `GlenOaks OR Stock`, `Elmhurst OR Stock` | Shelf counts per account, reorder points, weekly use. |
| `Pieces` | Every individual graft piece: Tab, Date, Initials, Location, Surgeon, Product, Size, Lot/SN, Unit price, PO (multi-piece cases expanded one row per piece). |

### Application-log columns (every month tab)
`Application Date · Initials · Alexian (patient ID, Alexian WCC only) · Location · Surgeon · Product · Size (cm) · Lot · Unit (SN) · Lot/SN (all pieces, e.g. LYU-260196_36003, comma-separated) · Price · Purchase Order [Received / Date Recv / Outstanding / PO Number / PO Match] · Sales · Commission · Submitted · Acct # · REF · Expiry · Notes · Case Sheet (link)`

### Row conventions
- **One row per case**, even for multi-piece cases (first lot/SN in Lot and Unit columns; all pieces in Lot/SN).
- **Row color:** light green = ready to submit to SFDC (PO received, not yet submitted) · light grey = submitted and showing in SFDC · no fill = submitted, not in SFDC yet.
- **Submitted column values:** blank → "Submitted" → "In SFDC".
- **PO Match:** "Certain" or "Unconfirmed" until AJ confirms a PO-to-patient match.
- **Price:** consistent per account — same product/size at the same account always carries the same price; a case missing a price takes the account's rate. Romano Lombard Grafix = $95/cm². St. Joseph Elgin Grafix 2x3 $900, 3x4 $1,210.50.
- **Patient ID (Alexian column)** used only for Alexian Brothers Wound Center — their POs match by patient ID.
- **Notes** carry application number ("App #5 of 6"), PO receipt date, case-sheet creation date, SFDC invoice/transaction once submitted.

---

## 3. Core workflow: application → PO → SFDC

1. **Application happens.** Source is one of: a photo of a handwritten case sheet, sticker photos + case details (Claude builds the case sheet PDF from `tools/case_sheet.py`), a "Grafted" tap in the Patients app, or a shipping-report export ("A JOSEPH PETERSON" rows).
2. **File the row** on the current month tab with all columns above; link the case sheet; mark Purchase Order "Outstanding". The patient's application count comes from the Patients app (newest sheet = total count; earlier sheets numbered backwards).
3. **PO arrives** (photo/PDF/email). Match to the row by account + patient (Alexian by patient ID; multi-unit POs count every unit as matched). Fill PO Number, Date Recv, Received; set PO Match Certain/Unconfirmed; put the PO number in the case sheet banner. Row turns light green. POs can also arrive first — they're added as blank rows and matched when the application comes in.
4. **AJ submits to SFDC** and says so in chat → Submitted = "Submitted", fill cleared.
5. **SFDC export/screenshot shows the invoice** → Submitted = "In SFDC", row light grey, invoice number in Notes.

### Account-specific rules
- **Ascension Saint Alexius (OR):** always bill & replace; orders flow to SFDC automatically; SFDC may show a different (replacement) Stravix serial than the case sheet — keep the case-sheet serial, not a billing error. OR shelf target 7 each of Stravix 3x6, Grafix 5x5, Stravix 2x4; restocks are not projected.
- **Advocate Good Samaritan WCC:** orders on its own POs, no case sheets; Grafix shipments are WCC stock — count as sales when shipped and offset projected visits in Projections ("WCC stock" section). Dr. Ayesh's patients come every two weeks. Dr. Luetkehans (podiatry) uses Downers Grove office.
- **Alexian Brothers Wound Center (WCC):** all patients weekly; cap 8 applications; POs match by patient ID; keep 5–10 of each Grafix size, 10–15 of 16mm. Alexian Ambulatory also weekly.
- **OASIS:** sold by the box, priced per box (Wound Matrix 10 sheets; Burn, Ultra 5 sheets; XL 1). One number per box (no per-sheet serial). A box ordered on a PO posts to SFDC on its own next day → row is "Submitted (pending SFDC)", not ready-to-submit. No OASIS projected until stock levels are given.
- **Stravix:** OR-based, no IVR/coverage needed. Upcoming OR cases projected as 3x6 upside / 2x4 downside until the implant record arrives.
- **Donated product** (e.g. Ankle and Foot Surgery LLC / Dr. Ziffra; Dr. McKanna 11/4): no sale, not in projections.
- **WFNU** ("Waiting For Notification of Use"): product left at a clinic without a PO; billed only if used — not a sale until then.
- **Transfers:** a consignment unit used at a different account gets a Transport Log (reason "Needed same day"; transported 7 AM, received 8 AM on the day used) and goes on the Transfers tab.

### Period rules
- Sales periods follow the company 2026 period calendar: P10 Sep 27–Oct 31, P11 Nov 1–Nov 28, P12 Nov 29–Dec 31. P11 = 19 shipping days.
- Period close: final number = In SFDC + Submitted. Rows still outstanding roll to the next period's month tab.
- Same-day ship cutoff is 3:30 PM; an order placed near cutoff on the last shipping day stays period-TBD until confirmed.
- Commission: 7% below quota, 10% at/above. Payout schedule on the README (Periods and payouts).

---

## 4. Projections (Projections / P11 Projections tab)

- **Current sales** = whole current-month tab (In SFDC + Submitted + outstanding).
- **Office/wound-care patients** projected weekly: current graft size price × remaining weeks in the period, starting the week after the current week. This week's scheduled visits count until they're filed; once filed they're not projected again.
- **OR cases** count once. Stravix = current cases + upcoming cases AJ lists.
- **Upside** = current plan. **Downside** = Grafix patients above 16mm with >4 applications drop one size from next week; 16mm patients with >4 applications get one last application; Dr. Bach's Good Sam patients start next week then every two weeks.
- A patient never goes back up a size — use the latest (smallest) size.
- "Potential new starts" section = approved patients not yet grafting (suggested size).
- Always include all accounts.

---

## 5. The phone apps (Claude-hosted; not reachable from the work laptop)

| App | Role | Data |
|---|---|---|
| **Chi West Graft Tracker** (main app) | Home screen: period pace (shipping day X of Y, % of month, % to target), To-do (ready to submit / waiting on PO / not in SFDC / stock alerts), November cases, projections, full year & payouts, Excel download. One-tap **Submitted** / **In SFDC** marks and notes for Claude; upload area that shows case sheets/POs only while unfiled. | db `actions`, `uploads`, `todos` |
| **New Case Sheet** / **Upload PO** | One-tap upload pages on the home screen. | db `uploads` each |
| **Chi West Patients** | Patient roster behind the tracker: one record per patient (account, provider, initials, name hidden until tapped, status pending/active/approved/hold/potential/OR/done/denied, product, size, frequency, application count, next appointment, Stravix case date, dated notes, files). Schedule by day (By time / By clinic), "+ Event" for lunches/in-services/OR cases, "Grafted" button uploads stickers/case sheet and bumps the count, clinic To-do (pending IVRs, follow-ups, stock to order), "No follow-up" flag, Caroline/Pete coverage (`cover` field, cleared every Friday night). No sales data in this app. | db `patients`, `inbox`, `uploads`, `events`, `meta/clinic` |
| **Chi West Targets** | 181 prospects by hospital/specialty, tiered; Today / Plan / Log / Activity tabs; outreach touches, lunches booked, objections; route plans for AJ, Caroline and Pete with nearby drop-ins; Sunday week email to associates. | db `outreach`, `plans/<day>~<name>`, `settings/prefs` |

Status workflow in Patients app: new patients enter as **Pending** → **Potential New Start** (approved, size suggested) → **Grafting**; patients missing from a clinic's weekly schedule go **On hold** ("verify with Chris"); OR cases sit in one **OR** category showing product + size (no application number); Stravix OR patients later seen in the office get one record under the office account.

---

## 6. Automation (scheduled tasks)

| Task | When | Does |
|---|---|---|
| Graft Tracker Update (GitHub) | Every morning | Pulls all unfiled uploads from the apps (Graft Tracker taps first, then its uploads, then the two upload pages, then the Patients app) plus `inbox/*` in the repo; files them into the tracker and patient list; rebuilds projections, stock, README/web pages; pushes everything to GitHub; refreshes the apps. |
| Gmail "Graft Inbox" Apps Script | On each forwarded email | Copies emails sent to ajpeterson113+grafts@gmail.com into `inbox/email/` in the repo. A line AJ typed above the forward is an instruction; forwarded text is information. |

"File the uploads" (said in chat) runs the same filing pass on demand. After every tracker change: Excel attached in chat, Graft Tracker app refreshed, GitHub Excel and README updated.

---

## 7. People and accounts (quick reference)

- **Associates:** Caroline and Pete cover visits AJ can't make; weekly "Chi West Coverage <dates>" email lists cases by day as "time Account – Dr. X, initials, product size". Keep each account's day with one person.
- **Shana:** reimbursement specialist (claims/IVR problems). **Eboni:** scheduler, St. Joe's WCC. **Chris:** verifies patient status at clinics.
- **Accounts (P11 forecast order):** Alexian WCC, Good Sam, St. Joseph Elgin, St. Alexius (OR), NW Suburban Foot & Ankle (HE = Hoffman Estates, EG = Elk Grove), Elmhurst, Sherman, Central DuPage Foot & Ankle (Downers Grove), Alexian Ambulatory, Romano Orthopaedics – Lombard (Dr. Rojas; Dr. McGann follow-ups), Suburban Foot & Ankle Center (Elmhurst main, Bloomingdale, Hoffman Estates), UChicago GlenOaks, Ankle and Foot Surgery LLC (Dr. Ziffra, 800 Biesterfield Ste 750).
- **Surgeons to keep straight:** St. Alexius OR "Dr. Singh" = Sutpal Singh DPM. "Dr. V" at Alexian WCC is not Palmieri. Good Sam podiatrist = Luetkehans. Drs. Greg and Adam Bryniczka operate at St. Alexius, Sherman, St. Joseph Elgin.

---

## 8. How Copilot should use this

- **Reading the tracker:** open `Graft_Tracker_BIOT_Chicago_West.xlsx` from the GitHub raw link; the current month tab is the live log; `Pieces` is the flat per-unit list for pivots; `Projections` for the forecast.
- **Answering "what's outstanding":** rows on the current month tab with Purchase Order = "Outstanding" (waiting on PO) or green fill (ready to submit). The README's To-do block has the same counts.
- **Reconciling SFDC exports:** match on PO number first, then lot/serial (Pieces tab), then account + date + product/size. St. Alexius serials may differ (bill & replace).
- **Drafting case sheets, coverage emails, transport logs:** follow the formats above; keep patient names out (initials only) unless the file is the patient list Excel.
- **Do not** edit the generated web pages, and do not assume laptop edits reach the apps — AJ relays changes in chat.
