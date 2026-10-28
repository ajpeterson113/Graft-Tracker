# Graft Inbox

Forward any email to **ajpeterson113+grafts@gmail.com**: a PO, an email about a case, or anything that needs doing. Within 5 minutes, [Code.gs](./Code.gs) (a Google Apps Script that runs in AJ's Gmail) copies it into [inbox/email](../../inbox/email). Claude files everything there at the morning update, or whenever AJ says "check the inbox":

- **POs** get matched to their cases.
- **Case details** go into the tracker and the Patients app.
- **Anything that needs doing** becomes an action item in the To do tab of the Graft Tracker app.

To give Claude an instruction, type a line above the forwarded message, for example "this PO is for J.L. 10/20" or "remind me Friday".

In Gmail, each copied email gets the **Graft Inbox** label and is archived. The script only copies email sent from AJ's own addresses: his Google account, plus `ALLOWED_FROM` (by default anything @smith-nephew.com). Email from any other sender gets the **Graft Inbox Blocked** label instead, and AJ gets one email about that sender.

## Setup (once, about 5 minutes, on a computer)

1. Sign in to Google as ajpeterson113@gmail.com and open [script.google.com](https://script.google.com). Click **New project** and name it *Graft Inbox*.
2. Delete everything in `Code.gs`. Copy all of [Code.gs](./Code.gs) (use the copy button at the top right of the file on GitHub) and paste it in. Click **Save**.
3. Click the gear icon (**Project Settings**). Under **Script properties**, click **Add script property**:
   - `GITHUB_TOKEN` = your GitHub token. This is the token you keep for this project. You can also make a new fine-grained token with access to Graft-Tracker only and **Contents: Read and write**.
   - Optional: `ALLOWED_FROM` = your work email. You only need this if it doesn't end in @smith-nephew.com. Separate several addresses with commas.

   Click **Save script properties**.
4. Go back to the editor (the `< >` icon). Pick **setup** in the function menu at the top and click **Run**. Google asks for access to Gmail. It warns that the app isn't verified because it's your own script. Click **Advanced**, then **Go to Graft Inbox**, then **Allow**.
5. Forward one test email to ajpeterson113+grafts@gmail.com. Within 5 minutes it gets the **Graft Inbox** label in Gmail and shows up in [inbox/email](../../inbox/email).

Tip: save the address as a contact named "Graft Inbox" in Outlook so it autocompletes.

## If something stops working

- **You get an email saying Graft Inbox couldn't copy to GitHub.** The GitHub token probably expired. Make a new one and paste it over `GITHUB_TOKEN` in Script properties. The waiting emails go through on the next run.
- **You get an email saying Graft Inbox ignored an email from your work address.** Add that address to `ALLOWED_FROM`, then forward the email again.
- **A forward never arrives.** Check Gmail's Spam folder and mark the email *Not spam*.
- **To turn the inbox off,** run **stop** in the script editor. Running **setup** turns it back on.
