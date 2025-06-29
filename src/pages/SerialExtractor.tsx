import * as XLSX from "xlsx";
import { useState } from "react";

export default function MultiSheetBuilder() {
  const [rawText, setRawText] = useState("");
  const [columnIndex, setColumnIndex] = useState(1);
  const [sheetName, setSheetName] = useState("");
  const [sheetDataList, setSheetDataList] = useState<
    { sheetName: string; data: object[] }[]
  >([]);

  const handleAddSheet = () => {
    const lines = rawText.split("\n").filter((line) => line.trim() !== "");

    const serials = lines
      .map((line) => {
        const cols = line.includes("\t") ? line.split("\t") : line.split(",");
        return cols[columnIndex]?.trim();
      })
      .filter(Boolean);

    if (!sheetName.trim()) {
      alert("Please enter a valid sheet name.");
      return;
    }

    if (serials.length === 0) {
      alert("No valid serial numbers found.");
      return;
    }

    const data = serials.map((serial) => ({
      "Availability, Serial No.": "Yes",
      "Serial No.": serial,
      "Availability, Lot No.": "Yes",
      "Lot No.": "",
      "Availability, Package No.": "Yes",
      "Package No.": "",
      "Quantity (Base)": 1,
      "Qty. to Handle (Base)": 1,
      "Appl.-to Item Entry": 0,
      "License key": "",
      "Bin Code": "",
    }));

    setSheetDataList((prev) => [...prev, { sheetName, data }]);
    setRawText(""); // clear for next sheet
    setSheetName(""); // reset name
  };

  const handleExport = () => {
    if (sheetDataList.length === 0) {
      alert("No sheets added to export.");
      return;
    }

    const workbook = XLSX.utils.book_new();

    sheetDataList.forEach(({ sheetName, data }) => {
      const worksheet = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    });

    XLSX.writeFile(workbook, "multi_sheet_serials.xlsx");
  };

  return (
    <div className="container py-5">
      <div className="card shadow">
        <div className="card-body">
          <h2 className="card-title mb-3">Multi-Sheet Excel Builder</h2>
          <p className="text-muted">
            Paste serial number rows, name the sheet, and add it. When done, export all sheets.
          </p>

          <div className="mb-3">
            <label className="form-label fw-bold">Paste Excel Rows</label>
            <textarea
              className="form-control"
              rows={8}
              value={rawText}
              onChange={(e) => setRawText(e.target.value)}
              placeholder="Paste Excel-like rows (tab/comma separated)"
            />
          </div>

          <div className="row g-3 mb-4">
            <div className="col-md-6">
              <label className="form-label fw-bold">Sheet Name</label>
              <input
                className="form-control"
                type="text"
                value={sheetName}
                onChange={(e) => setSheetName(e.target.value)}
                placeholder="e.g., Batch A"
              />
            </div>

            <div className="col-md-6">
              <label className="form-label fw-bold">
                Column Index (0-based for Serial No.)
              </label>
              <input
                className="form-control"
                type="number"
                value={columnIndex}
                onChange={(e) => setColumnIndex(Number(e.target.value))}
              />
            </div>
          </div>

          <div className="d-flex gap-2 mb-4">
            <button className="btn btn-outline-primary w-100" onClick={handleAddSheet}>
              ➕ Add to Workbook
            </button>
            <button className="btn btn-success w-100" onClick={handleExport}>
              📤 Export Excel File
            </button>
          </div>

          {sheetDataList.length > 0 && (
            <div className="alert alert-info">
              ✅ <strong>{sheetDataList.length}</strong> sheet(s) added:
              <ul className="mb-0">
                {sheetDataList.map((s, i) => (
                  <li key={i}>
                    <strong>{s.sheetName}</strong> – {s.data.length} entries
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
