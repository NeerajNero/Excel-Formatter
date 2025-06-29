📦 Serial Number Multi-Sheet Excel Exporter
A lightweight React + Bootstrap web app that lets you:

✅ Paste tab/comma-separated data copied from Excel

✅ Extract serial numbers from any column (you choose the index)

✅ Add each dataset as a separate worksheet with a custom name (e.g., "Batch A", "Lot 1")

✅ Export everything as a single Excel file with multiple sheets

🚀 Features
Paste serial number rows (copied from Excel or CSV)

Define which column contains the Serial No.

Assign a name to each sheet (e.g. "Batch A")

Add unlimited sheets to memory

Export a final Excel file with all your custom sheets

Built using React and Bootstrap

Excel file generation handled by SheetJS (xlsx)

🖥️ Demo Flow
Paste your data rows

Enter the sheet name (e.g. "Batch A")

Set the serial number column index (0-based)

Click "Add to Workbook"

Repeat steps 1–4 as needed

Click "Export Excel File" to download the .xlsx

📦 Tech Stack
React

Bootstrap 5

SheetJS (xlsx) for Excel export

📸 Screenshots
Add a few screenshots of your UI once hosted or running locally

🛠️ Getting Started
1. Clone this repo
bash
Copy
Edit
git clone https://github.com/your-username/serial-multi-sheet-exporter.git
cd serial-multi-sheet-exporter
2. Install dependencies
bash
Copy
Edit
npm install
# or
yarn install
3. Start the app
bash
Copy
Edit
npm run dev
