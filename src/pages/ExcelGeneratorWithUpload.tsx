import * as XLSX from "xlsx";
import React, { useState } from "react";

// The required output structure for each row
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

// The structure for each sheet held in state
interface SheetData {
  sheetName: string;
  data: RowData[];
}

// NEW: Structure for our notification state
interface Notification {
    message: string;
    type: 'success' | 'error' | 'warning';
}

// Helper function to build a single RowData object from a serial number
const buildRowFromSerial = (serial: string): RowData => ({
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
});

export default function MultiSheetBuilder() {
  // --- STATE MANAGEMENT ---
  const [sheetDataList, setSheetDataList] = useState<SheetData[]>([]);
  
  // States for manual text input
  const [rawText, setRawText] = useState<string>("");
  const [columnIndex, setColumnIndex] = useState<number>(0);
  const [sheetName, setSheetName] = useState<string>("");
  const [lotOnlyMode, setLotOnlyMode] = useState<boolean>(false);
  const [hasHeader, setHasHeader] = useState<boolean>(false);
  const [editingSheetIndex, setEditingSheetIndex] = useState<number | null>(null);

  // States for improved user experience
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [notification, setNotification] = useState<Notification | null>(null);

  // States for configurable column names
  const [partNumberCol, setPartNumberCol] = useState<string>("Part number");
  const [qtyCol, setQtyCol] = useState<string>("Qty");
  const [serialCol, setSerialCol] = useState<string>("serail number");

  // --- FILE PROCESSING LOGIC ---
  const handleFileProcess = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setNotification(null);
    setSheetDataList([]);

    setTimeout(() => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target?.result as ArrayBuffer);
                const workbook = XLSX.read(data, { type: "array" });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                let json = XLSX.utils.sheet_to_json<any>(worksheet);

                // Group data by the user-defined part number column
                const groupedByPartNumber = json.reduce((acc, row) => {
                    const partNumber = row[partNumberCol];
                    if (!partNumber || partNumber === partNumberCol) return acc;
                    if (!acc[partNumber]) acc[partNumber] = [];
                    acc[partNumber].push(row);
                    return acc;
                }, {} as Record<string, any[]>);

                let totalDuplicatesFound = 0;
                
                const newSheets: SheetData[] = Object.entries(groupedByPartNumber)
                    .reduce((acc: SheetData[], [partNumber, rows]) => {
                        const typedRows = rows as Record<string, any>[];
                        const hasQuantityGreaterThanOne = typedRows.some(row => row[qtyCol] > 1);

                        if (!hasQuantityGreaterThanOne) {
                            const uniqueSerials = new Set<string>();
                            const serials = typedRows.map(row => row[serialCol]).filter(Boolean);
                            
                            const uniqueSerialData = serials
                                .filter((serial: string) => {
                                    if (uniqueSerials.has(serial)) {
                                        totalDuplicatesFound++;
                                        return false;
                                    }
                                    uniqueSerials.add(serial);
                                    return true;
                                })
                                .map(buildRowFromSerial);

                            if (uniqueSerialData.length > 0) {
                                acc.push({ sheetName: String(partNumber), data: uniqueSerialData });
                            }
                        }
                        return acc;
                    }, []);

                setSheetDataList(newSheets);
                let successMessage = `${newSheets.length} sheets created successfully.`;
                if(totalDuplicatesFound > 0) {
                    successMessage += ` Found and removed ${totalDuplicatesFound} duplicate serial number(s).`
                }
                setNotification({ message: successMessage, type: 'success' });

            } catch (error) {
                console.error("Error processing file:", error);
                setNotification({ message: "Failed to process file. Check column names and file format.", type: 'error' });
            } finally {
                setIsLoading(false);
                event.target.value = "";
            }
        };
        reader.readAsArrayBuffer(file);
    }, 50);
  };

  // --- MANUAL INPUT LOGIC ---
  const buildSheetDataFromText = (entries: string[]): RowData[] => {
    return entries.map((value): RowData => ({
      ...buildRowFromSerial(lotOnlyMode ? "" : value),
      "Lot No.": lotOnlyMode ? value : "",
    }));
  };

  const handleAddOrUpdateSheet = () => {
    const lines = rawText.split("\n").filter(line => line.trim() !== "");
    const slicedLines = hasHeader ? lines.slice(1) : lines;
    const entries = slicedLines
      .map(line => (line.includes("\t") ? line.split("\t") : line.split(","))[columnIndex]?.trim())
      .filter(Boolean);

    if (entries.length === 0) {
      setNotification({ message: "No valid entries found in the text area.", type: 'error' });
      return;
    }

    const finalSheetName = sheetName.trim() || `Sheet ${sheetDataList.length + 1}`;
    const newSheet: SheetData = { sheetName: finalSheetName, data: buildSheetDataFromText(entries) };

    if (editingSheetIndex !== null) {
      const updatedList = [...sheetDataList];
      updatedList[editingSheetIndex] = newSheet;
      setSheetDataList(updatedList);
      setEditingSheetIndex(null);
    } else {
      setSheetDataList(prev => [...prev, newSheet]);
    }
    setNotification({ message: `Sheet "${finalSheetName}" was added/updated.`, type: 'success' });
    setRawText("");
    setSheetName("");
  };

  // --- SHARED UTILITY FUNCTIONS ---
  const handleExport = () => {
    if (sheetDataList.length === 0) {
      setNotification({ message: "No sheets to export.", type: 'error' });
      return;
    }
    const workbook = XLSX.utils.book_new();
    const summaryData = sheetDataList.map(s => ({ "Part Number": s.sheetName, "Serial Count": s.data.length }));
    const summaryWorksheet = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(workbook, summaryWorksheet, "Summary");

    sheetDataList.forEach(({ sheetName, data }) => {
      const worksheet = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    });

    XLSX.writeFile(workbook, "multi_sheet_output.xlsx");
    setNotification({ message: "Excel file with summary has been exported!", type: 'success' });
  };

  // NEW: Function to copy a sheet's content to the clipboard
  const handleCopySheet = (index: number) => {
    const sheet = sheetDataList[index];
    if (!sheet || sheet.data.length === 0) {
      setNotification({ message: "Cannot copy an empty sheet.", type: 'warning' });
      return;
    }

    const headers = Object.keys(sheet.data[0]) as (keyof RowData)[];
    const headerString = headers.join("\t");
    const rowsString = sheet.data.map(row =>
        headers.map(h => row[h] ?? "").join("\t")
    ).join("\n");

    const clipboardText = [headerString, rowsString].join("\n");
    navigator.clipboard.writeText(clipboardText).then(() => {
        setNotification({ message: `Copied data for "${sheet.sheetName}" to clipboard.`, type: 'success' });
    }).catch(err => {
        setNotification({ message: "Failed to copy data.", type: 'error' });
        console.error('Clipboard error:', err);
    });
  };

  const handleEditSheet = (index: number) => {
    const sheet = sheetDataList[index];
    if (!sheet) return;
    const isLot = sheet.data.some(d => d["Lot No."] !== "");
    setRawText(sheet.data.map(row => isLot ? row["Lot No."] : row["Serial No."]).join("\n"));
    setSheetName(sheet.sheetName);
    setLotOnlyMode(isLot);
    setEditingSheetIndex(index);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const cancelEdit = () => {
    setEditingSheetIndex(null);
    setRawText("");
    setSheetName("");
  };

  const handleDeleteSheet = (index: number) => {
    if (window.confirm(`Delete "${sheetDataList[index].sheetName}"?`)) {
      const sheetName = sheetDataList[index].sheetName;
      setSheetDataList(prev => prev.filter((_, i) => i !== index));
      setNotification({ message: `Sheet "${sheetName}" deleted.`, type: 'warning' });
    }
  };

  // --- JSX / RENDER ---
  return (
    <div className="container py-5">
      {notification && (
        <div 
          className={`alert alert-${notification.type === 'error' ? 'danger' : notification.type} alert-dismissible fade show`} 
          role="alert"
        >
          {notification.message}
          <button type="button" className="btn-close" onClick={() => setNotification(null)} aria-label="Close"></button>
        </div>
      )}

      <div className="card shadow mb-4">
        <div className="card-body">
          <h2 className="card-title">1. Automated Processing from File</h2>
          <p className="text-muted">Upload an Excel/CSV file to automatically generate sheets. This will replace any existing sheets.</p>
          <div className="row g-3 mb-3">
            <div className="col-md-4"><label className="form-label fw-bold">Part Number Column</label><input type="text" className="form-control" value={partNumberCol} onChange={e => setPartNumberCol(e.target.value)} /></div>
            <div className="col-md-4"><label className="form-label fw-bold">Quantity Column</label><input type="text" className="form-control" value={qtyCol} onChange={e => setQtyCol(e.target.value)} /></div>
            <div className="col-md-4"><label className="form-label fw-bold">Serial Number Column</label><input type="text" className="form-control" value={serialCol} onChange={e => setSerialCol(e.target.value)} /></div>
          </div>
          <div className="mb-3">
            <label htmlFor="file-upload" className="form-label fw-bold">Upload File</label>
            <input id="file-upload" className="form-control" type="file" accept=".xlsx, .xls, .csv" onChange={handleFileProcess} disabled={isLoading} />
          </div>
          {isLoading && <div className="text-primary fw-bold">Processing your file, please wait...</div>}
        </div>
      </div>

      <div className="card shadow">
        <div className="card-body">
            <h2 className="card-title">2. Manual Sheet Builder</h2>
            <p className="text-muted">Paste data to add or edit sheets manually.</p>
            <div className="mb-3"><label className="form-label fw-bold">Paste Serial/Lot Numbers</label><textarea className="form-control" rows={8} value={rawText} onChange={(e) => setRawText(e.target.value)} placeholder="Paste rows here..."/></div>
            <div className="row g-3 mb-3">
                <div className="col-md-4"><label className="form-label fw-bold">Sheet Name</label><input className="form-control" type="text" value={sheetName} onChange={(e) => setSheetName(e.target.value)} placeholder="Optional"/></div>
                <div className="col-md-4"><label className="form-label fw-bold">Column Index (0-based)</label><input className="form-control" type="number" value={columnIndex} onChange={(e) => setColumnIndex(Number(e.target.value))}/></div>
                <div className="col-md-4 d-flex align-items-end justify-content-between">
                    <div className="form-check"><input className="form-check-input" type="checkbox" checked={lotOnlyMode} onChange={() => setLotOnlyMode((p) => !p)} id="lotOnlyCheck"/><label className="form-check-label fw-bold" htmlFor="lotOnlyCheck">Lot Mode</label></div>
                    <div className="form-check ms-3"><input className="form-check-input" type="checkbox" checked={hasHeader} onChange={() => setHasHeader((p) => !p)} id="hasHeaderCheck"/><label className="form-check-label fw-bold" htmlFor="hasHeaderCheck">Has Header</label></div>
                </div>
            </div>
            <div className="d-flex gap-2 mb-4">
                <button className="btn btn-primary w-100" onClick={handleAddOrUpdateSheet}>{editingSheetIndex !== null ? "💾 Update Sheet" : "➕ Add Sheet"}</button>
                {editingSheetIndex !== null && (<button className="btn btn-secondary w-100" onClick={cancelEdit}>❌ Cancel Edit</button>)}
            </div>
        </div>
      </div>

      {sheetDataList.length > 0 && (
        <div className="card shadow mt-4">
          <div className="card-body">
            <div className="d-flex justify-content-between align-items-center mb-3"><h2 className="card-title mb-0">Generated Sheets</h2><button className="btn btn-success" onClick={handleExport}>📤 Export All w/ Summary</button></div>
            <ul className="list-group">
              {sheetDataList.map((s, i) => (
                <li key={i} className="list-group-item d-flex justify-content-between align-items-center">
                  <span><strong>{s.sheetName}</strong> – {s.data.length} rows</span>
                  <div className="btn-group">
                    {/* NEW: Copy Button Added */}
                    <button className="btn btn-sm btn-outline-secondary" onClick={() => handleCopySheet(i)}>📋 Copy</button>
                    <button className="btn btn-sm btn-outline-primary" onClick={() => handleEditSheet(i)}>✏️ Edit</button>
                    <button className="btn btn-sm btn-outline-danger" onClick={() => handleDeleteSheet(i)}>🗑️ Delete</button>
                  </div>
                </li>
              ))}
            </ul>
          </div>
        </div>
      )}
    </div>
  );
}