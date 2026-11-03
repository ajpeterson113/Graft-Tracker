# Ready-to-paste prompts for M365 Copilot

All assume the files from the starter folder are in OneDrive › **Chi West Graft Tracker**.

## Setup / context
- Read "Copilot_Overview" in my Chi West Graft Tracker folder and summarize the rules for filing a case sheet, matching a PO, and the row colors. Keep it as a checklist.

## Tracker questions (Excel: Graft_Tracker_BIOT_Chicago_West.xlsx)
- Using the November tab, list every row where Purchase Order is "Outstanding": date, initials, account, product, size, price. Total the price.
- Which November rows are ready to submit (PO received, Submitted column blank)? Total them.
- From the Pieces tab, how many 16mm Grafix pieces were applied at Alexian Brothers Wound Center in October, and what was the total?
- Compare the P11 Projections tab upside and downside by account and tell me which accounts are driving the gap.

## Filing a case sheet (photo attached or in the folder)
- Read the attached case sheet. Using the column list in Copilot_Overview section 2, give me one row for the current month tab in table form: Application Date, Initials, Alexian ID (Alexian WCC only), Location, Surgeon, Product, Size, Lot, Unit SN, Lot/SN for every piece, Price from the Prices tab for that account, REF, Expiry, Purchase Order = Outstanding. Flag anything you could not read.
- Then: Add that row to the November tab of Graft_Tracker_BIOT_Chicago_West.xlsx (open the file in Excel first, then ask Copilot in Excel).

## Matching a PO (email or PDF)
- Read the PO I forwarded to myself today. Match it to the right row on the November tab by account and patient (Alexian by patient ID; count every unit on a multi-unit PO). Tell me the row, whether the match is Certain or Unconfirmed, and what to fill in: PO Number, Date Received, Received = Yes.
- Check every PO received this month against the tracker and tell me if any graft was invoiced on another patient's PO.

## SFDC reconciliation (export saved to the folder)
- Compare the SFDC export to the November tab: match on PO number, then lot/serial, then account + date + product. List rows that should move to "In SFDC" with the invoice number, and anything in SFDC that is not in the tracker. St. Alexius serials may differ (bill & replace) — do not flag those.

## Calendar / patients (Excel: Chi_West_Patient_List.xlsx)
- Look at my calendar for next week. For each event titled "Dr. X, initials, A#, product size", update the matching patient's Next Appointment in the All Patients tab and flag it "Verify" if the time changed from the previous week.
- Which grafting patients in the All Patients tab have no appointment on my calendar next week?
- Draft the "Chi West Coverage <dates>" email to Caroline and Pete for next week from the Cover column: list cases by day as "time Account – Dr. X, initials, product size". Plain text.

## Projections
- Rebuild the P11 forecast from the All Patients tab using the rules in Copilot_Overview section 4 (weekly visits × current size price through Nov 28, OR cases once, upside vs downside). Show by account.

## Copilot Studio agent (paste as the agent description)
- You are the Chi West Tracker agent for AJ Peterson. Your knowledge is the Chi West Graft Tracker folder. Apply every rule in Copilot_Overview sections 2–4. Use patient initials only, never full names. When asked to file a case sheet, PO, or shipping report, return the exact tracker row(s) in the column order of the November tab and say which rows change state. Flag any PO that looks like it belongs to a different patient.
