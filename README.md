# MVU Wise Daily Report

## Run on Windows / VS Code

Open this folder in VS Code Terminal and run:

```powershell
npm install
node server.js
```

Then open:

`http://localhost:5173/upload.html`

## Daily workflow

1. On first setup, upload Village Mapping and Week Off Master.
2. Upload the two-date DetailedReport Excel file.
3. Click Generate Report.
4. The Result page shows the complete MVU-wise report.

Village Mapping and Week Off Master stay locked after upload. Click **Hard Reset** (password `1122`) when those masters need to be uploaded again. After a successful Hard Reset, their upload panels become available below Daily Detailed Report.

## Report calculations

- Cases Received = first-date Total Received + second-date Total Received.
- Attended Cases = first-date Attend + second-date Attended.
- Today Pending Cases = second-date Pending only.
- Overall Attend % = Attended Cases / Cases Received × 100.
- Dist. Head Quarter rows are retained.
- District subtotal rows are not displayed.
- One Overall Grand Total is displayed.
- TA, Enquiry and WT rows are filtered before aggregation.
- Open/Appointed/Reassign tickets are Pending.
- Visited Farmer is Attend; other SubStatus values are Not Attend.
- District/Block/Village information is taken from the saved masters, with fallback repair from report data where possible.

## Result page

- District filter, Refresh, Export Excel and Upload Page controls remain outside the report table.
- Copy PNG captures the report header, summary and full table, while website controls are excluded.
- The report always loads from the server and does not use stale sessionStorage data.

## Data

Persistent application data is stored in `data/mvu-data.json`.
