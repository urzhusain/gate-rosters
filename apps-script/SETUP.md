# Setup (about 15 minutes, once)

1. Go to sheets.google.com and create a new blank spreadsheet. Name it **Security Control**.
2. In the spreadsheet: **Extensions > Apps Script**.
3. Delete the sample code. Paste the whole of `Code.gs`. Click Save.
4. In the function dropdown choose **setup** and click **Run**. Allow the permissions Google asks for (it needs Sheets and Drive). When it finishes, the sheet has new tabs.
5. Open the **Settings** tab. Change the three PINs (supervisor_pin, control_pin, head_pin). Keep them different.
6. Back in Apps Script: **Deploy > New deployment > type: Web app**. Execute as: **Me**. Who has access: **Anyone**. Click Deploy and copy the **Web app URL**.
7. Send that URL to Claude. It goes into `supervisor/config.js` and the app goes live.
