import * as XLSX from "xlsx";
import { useState } from "react";

interface RowData {
  "Availability, Serial No.": string;
  "Serial No.": string;
  "Availability, Lot No.": string;
  "Lot No.": string;
  "Availability, Package No.": string;
  "Package No.": string;
  "Quantity (Base)": number;
  "Qty. to Handle (Base)": number;
  "Appl.-to Item Entry": number;
  "License key": string;
  "Bin Code": string;
}

interface SheetData {
  sheetName: string;
  data: RowData[];
}

export default function MultiSheetBuilder() {
  const [rawText, setRawText] = useState<string>("");
  const [columnIndex, setColumnIndex] = useState<number>(1);
  const [sheetName, setSheetName] = useState<string>("");
  const [lotOnlyMode, setLotOnlyMode] = useState<boolean>(false);
  const [hasHeader, setHasHeader] = useState<boolean>(false);
  const [sheetDataList, setSheetDataList] = useState<SheetData[]>([]);
  const [editingSheetIndex, setEditingSheetIndex] = useState<number | null>(null);

  const buildSheetData = (entries: string[]): RowData[] => {
    return entries.map((value): RowData => ({
      "Availability, Serial No.": "Yes",
      "Serial No.": lotOnlyMode ? "" : value,
      "Availability, Lot No.": "Yes",
      "Lot No.": lotOnlyMode ? value : "",
      "Availability, Package No.": "Yes",
      "Package No.": "",
      "Quantity (Base)": 1,
      "Qty. to Handle (Base)": 1,
      "Appl.-to Item Entry": 0,
      "License key": "",
      "Bin Code": "",
    }));
  };

  const handleAddOrUpdateSheet = () => {
    const lines = rawText.split("\n").filter((line) => line.trim() !== "");

    const slicedLines = hasHeader ? lines.slice(1) : lines;

    const entries = slicedLines
      .map((line) => {
        const cols = line.includes("\t") ? line.split("\t") : line.split(",");
        return cols[columnIndex]?.trim();
      })
      .filter((entry): entry is string => Boolean(entry));

    if (entries.length === 0) {
      alert("No valid entries found.");
      return;
    }

    const finalSheetName = sheetName.trim() || `Sheet ${sheetDataList.length + 1}`;

    const newSheet: SheetData = {
      sheetName: finalSheetName,
      data: buildSheetData(entries),
    };

    if (editingSheetIndex !== null) {
      const updatedList = [...sheetDataList];
      updatedList[editingSheetIndex] = newSheet;
      setSheetDataList(updatedList);
      setEditingSheetIndex(null);
    } else {
      setSheetDataList((prev) => [...prev, newSheet]);
    }

    setRawText("");
    setSheetName("");
  };

  const handleExport = () => {
    if (sheetDataList.length === 0) {
      alert("No sheets to export.");
      return;
    }

    const workbook = XLSX.utils.book_new();
    sheetDataList.forEach(({ sheetName, data }) => {
      const worksheet = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    });
    XLSX.writeFile(workbook, "multi_sheet_serials.xlsx");
  };

  const handleCopySheet = (index: number) => {
    const sheet = sheetDataList[index];
    if (!sheet || sheet.data.length === 0) return;

    const headers = Object.keys(sheet.data[0]) as (keyof RowData)[];
    const rows = sheet.data.map((row) =>
      headers.map((h) => row[h] ?? "").join("\t")
    );

    navigator.clipboard.writeText([headers.join("\t"), ...rows].join("\n"));
    alert(`✅ Sheet "${sheet.sheetName}" copied!`);
  };

  const handleEditSheet = (index: number) => {
    const sheet = sheetDataList[index];
    if (!sheet || sheet.data.length === 0) return;

    const isLotOnly = sheet.data[0]["Serial No."] === "";

    const columnData = sheet.data.map((row) =>
      isLotOnly ? row["Lot No."] : row["Serial No."]
    );

    setRawText(columnData.join("\n"));
    setSheetName(sheet.sheetName);
    setLotOnlyMode(isLotOnly);
    setEditingSheetIndex(index);
  };

  const cancelEdit = () => {
    setEditingSheetIndex(null);
    setRawText("");
    setSheetName("");
  };

  const handleDeleteSheet = (index: number) => {
    const confirmDelete = window.confirm(
      `Are you sure you want to delete "${sheetDataList[index].sheetName}"?`
    );
    if (!confirmDelete) return;

    setSheetDataList((prev) => prev.filter((_, i) => i !== index));
  };

  return (
    <div className="container py-5">
      <div className="card shadow">
        <div className="card-body">
          <h2 className="card-title mb-3">Multi-Sheet Excel Builder</h2>
          <p className="text-muted">
            Paste serial or lot numbers, add/edit/export/delete sheets, or copy output.
          </p>

          <div className="mb-3">
            <label className="form-label fw-bold">Paste Serial/Lot Numbers</label>
            <textarea
              className="form-control"
              rows={8}
              value={rawText}
              onChange={(e) => setRawText(e.target.value)}
              placeholder="Paste rows here (tab/comma-separated if needed)"
            />
          </div>

          <div className="row g-3 mb-3">
            <div className="col-md-4">
              <label className="form-label fw-bold">Sheet Name</label>
              <input
                className="form-control"
                type="text"
                value={sheetName}
                onChange={(e) => setSheetName(e.target.value)}
                placeholder="Optional (e.g., Batch A)"
              />
            </div>

            <div className="col-md-4">
              <label className="form-label fw-bold">Column Index</label>
              <input
                className="form-control"
                type="number"
                value={columnIndex}
                onChange={(e) => setColumnIndex(Number(e.target.value))}
              />
            </div>

            <div className="col-md-4 d-flex align-items-end justify-content-between">
              <div className="form-check">
                <input
                  className="form-check-input"
                  type="checkbox"
                  checked={lotOnlyMode}
                  onChange={() => setLotOnlyMode((prev) => !prev)}
                  id="lotOnlyCheck"
                />
                <label className="form-check-label fw-bold" htmlFor="lotOnlyCheck">
                  Lot Number Mode
                </label>
              </div>
              <div className="form-check ms-3">
                <input
                  className="form-check-input"
                  type="checkbox"
                  checked={hasHeader}
                  onChange={() => setHasHeader((prev) => !prev)}
                  id="hasHeaderCheck"
                />
                <label className="form-check-label fw-bold" htmlFor="hasHeaderCheck">
                  First row is header
                </label>
              </div>
            </div>
          </div>

          <div className="d-flex gap-2 mb-4">
            <button
              className="btn btn-primary w-100"
              onClick={handleAddOrUpdateSheet}
            >
              {editingSheetIndex !== null ? "💾 Update Sheet" : "➕ Add Sheet"}
            </button>
            {editingSheetIndex !== null && (
              <button className="btn btn-secondary w-100" onClick={cancelEdit}>
                ❌ Cancel Edit
              </button>
            )}
            <button className="btn btn-success w-100" onClick={handleExport}>
              📤 Export Excel
            </button>
          </div>

          {sheetDataList.length > 0 && (
            <div className="alert alert-info">
              ✅ <strong>{sheetDataList.length}</strong> sheet(s) added:
              <ul className="mb-0 list-unstyled">
                {sheetDataList.map((s, i) => (
                  <li key={i} className="mb-2">
                    <strong>{s.sheetName}</strong> – {s.data.length} rows
                    <div className="btn-group ms-2">
                      <button
                        className="btn btn-sm btn-outline-secondary"
                        onClick={() => handleCopySheet(i)}
                      >
                        📋 Copy
                      </button>
                      <button
                        className="btn btn-sm btn-outline-primary"
                        onClick={() => handleEditSheet(i)}
                      >
                        ✏️ Edit
                      </button>
                      <button
                        className="btn btn-sm btn-outline-danger"
                        onClick={() => handleDeleteSheet(i)}
                      >
                        🗑️ Delete
                      </button>
                    </div>
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
