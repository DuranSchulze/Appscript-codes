## Setup Instructions (First Time)

1. Open the bound Google Spreadsheet
2. Go to **Extensions → Apps Script** and paste the script (or it should already be there)
3. In Apps Script editor, go to **Project Settings → Google Cloud Platform project** and ensure the **Google Drive API** is enabled in the linked GCP project
4. Run **📁 Client Document Monitor → 🔧 Initial Setup** — this creates all sheets
5. Run **Configure Drives → 📄 Set Client Documents Drive ID** — paste the folder ID from the Drive URL
6. Run **Configure Drives → ⚖️ Set Scanned Pleadings Drive ID** — paste the folder ID
7. Run **🔑 Set Gemini API Key** — paste your key from [https://aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
8. Run **🤖 Select AI Model** — fetches the live model list; pick the `★ RECOMMENDED` one
9. Run **🩺 Diagnostics → 🧪 Test Gemini AI Integration** to verify the full AI pipeline works
10. Run **🩺 Diagnostics → 👥 Test User Tracking** to verify email extraction works
11. Run **🔍 Full Scanning → 🔄 Scan Both Drives** for the initial population
12. Run **⏰ Setup Daily Schedule** to enable automatic daily scans
