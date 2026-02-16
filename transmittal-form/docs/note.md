- [DONE] in the transmittal id the generated one like 202601-0001 , if the user encoded this means this should not appear again, This should be unique data that no one can replicate so that in the history when I tried to browse again I can see and potentially edit some encoded things inside.
  - Server now generates the transmittal number atomically inside a DB transaction (no duplicates possible). The Transmittal ID field is read-only in the UI.
- [DONE] System will have a gsheet setup where we can link a g-sheet so one there is a created row of what transmittal document.
  - In the History tab, users can paste a Google Sheets URL to link. On every new transmittal creation a summary row is auto-appended (Transmittal No., Date, Project, Recipient, etc.). Headers are written automatically on first use. Google Sheets OAuth scope added.
- [DONE] Internet detector so that there will a pop up when there is no internet so user will be alerted
  - A fullscreen modal with a "No Internet Connection" message appears instantly when the browser goes offline and auto-dismisses when connectivity is restored. Rendered globally via the root layout.

## PLAN

- [DONE] Dropdown for the Original Copies, Photocopies, To Sign are the choices for the dropdown
- [DONE] Transmittal Documents should have Quantity ADD or Subtract, a Plus and Minus button so that user dont need to type but can select those button to add value or number
  - Added +/- quantity controls directly in the table rows while still allowing manual typing. Quantity never goes below 1.
- [DONE] Document # Reference Number should be auto generated from the BROWSE DRIVE and you need to unify the formats of it. Also the remarks should be fixed.
  - Drive imports now auto-generate a standardized document number format and all Drive import flows use a unified remark: "Imported from Google Drive".

- [DONT DO THIS] Connection and integration or merging from this stack to Firebase for server hosting and using of database that we can use, maybe removing PRISMA and converting to Firestore

NOTES ni BOSS

- [] btw, meron sana ako ipa fix sa iyo - ung Notary automation na ginawa ko. I think, same issue with the PCF ng accounting. Since, nag palit sa gemini ng model, need update. I will provide you link later. And meron pang isa - summary of emails (hindi na na cacapture ung accurate title ng document. You can ask Belle from receipt - regarding the Notary app, and the Summary of emails. BOth app needs to have an updated Gemini model.

- [] For Doc/Ref #, walang format. But it should be extracted from the document number itself, eg. Official Receipts issued by the BIR - duon makukuha ung ref #. As for the remarks, meron din nilalagay ang team na note kung original, certified or photocopy ba un for transmitted na document nila
- [] I needed something to add in the system. you see we import files to the transmittal items we use AI in here for each you see we need the tagging of the documents # / reference # based on the file content of the file we can see anything related like receipt or any documents na may marecognize na number (eg title number, invoice/receipt number, eCAR number) in each documents. I want that each process will undergo with ai searching data like this and use that as a reference or document number.
