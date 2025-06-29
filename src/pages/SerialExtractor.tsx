import * as XLSX from "xlsx";
import { useState } from "react";

export default function SerialExtractorBootstrap() {
  const [rawText, setRawText] = useState("");
  const [columnIndex, setColumnIndex] = useState(1); // default: second column

  const handleExport = () => {
    const lines = rawText.split("\n").filter((line) => line.trim() !== "");

    const serials = lines
      .map((line) => {
        const cols = line.includes("\t") ? line.split("\t") : line.split(",");
        return cols[columnIndex]?.trim();
      })
      .filter(Boolean);

    if (serials.length === 0) {
      alert("No valid serial numbers found. Please check your input and column index.");
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

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Serials");
    XLSX.writeFile(workbook, "custom_serial_data.xlsx");
  };

  return (
    <div className="container py-5">
      <div className="card shadow">
        <div className="card-body">
          <h2 className="card-title mb-3">Serial Number Extractor</h2>
          <p className="text-muted mb-4">
            Paste tab or comma-separated data copied from Excel. Select the column index containing serial numbers.
          </p>

          <div className="mb-3">
            <label htmlFor="excelInput" className="form-label fw-semibold">
              Excel Data
            </label>
            <textarea
              id="excelInput"
              className="form-control"
              rows={10}
              placeholder="Paste your Excel rows here..."
              value={rawText}
              onChange={(e) => setRawText(e.target.value)}
            ></textarea>
          </div>

          <div className="mb-3">
            <label htmlFor="columnIndex" className="form-label fw-semibold">
              Serial No. Column Index <small className="text-muted">(0-based)</small>
            </label>
            <input
              type="number"
              className="form-control w-25"
              id="columnIndex"
              value={columnIndex}
              onChange={(e) => setColumnIndex(Number(e.target.value))}
            />
          </div>

          <button
            className="btn btn-primary w-20"
            onClick={handleExport}
          >
            Generate Excel
          </button>
        </div>
      </div>
    </div>
  );
}
