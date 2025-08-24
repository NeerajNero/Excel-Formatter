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

// Structure for our notification state
interface Notification {
    message: string;
    type: 'success' | 'error' | 'warning';
}

// A type definition for a row of data from the uploaded file
type DataRow = { [key: string]: any };

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

// Helper function to get the character type pattern of a string
const getSequence = (s: string) => {
    return s.split('').map(char => isNaN(parseInt(char)) ? 'L' : 'N').join('');
};


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

  // States for the interactive column selection workflow
  const [fileHeaders, setFileHeaders] = useState<string[]>([]);
  const [parsedData, setParsedData] = useState<DataRow[]>([]);
  const [selectedPartNumberCol, setSelectedPartNumberCol] = useState<string>("");
  const [selectedInvoiceCol, setSelectedInvoiceCol] = useState<string>("");
  const [selectedQtyCol, setSelectedQtyCol] = useState<string>("");
  const [selectedSerialCol, setSelectedSerialCol] = useState<string>("");

  // State for processing mode
  const [processingMode, setProcessingMode] = useState<'invoice' | 'consolidate'>('invoice');


  // --- FILE PROCESSING LOGIC ---

  // Step 1: Read the file, manually separating header from data
  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setNotification(null);
    setSheetDataList([]);
    setFileHeaders([]);
    setParsedData([]);
    setSelectedPartNumberCol("");
    setSelectedInvoiceCol("");
    setSelectedQtyCol("");
    setSelectedSerialCol("");
    setProcessingMode('invoice');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const dataAsArray = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
            
            const nonEmptyRows = dataAsArray.filter(row => row && row.length > 0 && row.some(cell => cell !== null && cell !== ''));
            
            if (nonEmptyRows.length < 2) {
              throw new Error("File appears to be empty or has no data rows.");
            }

            const headers: string[] = nonEmptyRows[0];
            const dataRows = nonEmptyRows.slice(1);

            const jsonData: DataRow[] = dataRows.map(rowArray => {
                const rowObject: DataRow = {};
                headers.forEach((header, index) => {
                    if (typeof header === 'string') {
                        rowObject[header] = rowArray[index];
                    }
                });
                return rowObject;
            });
            
            setFileHeaders(headers);
            setParsedData(jsonData);

            // Auto-select common column names
            const defaultPartCol = headers.find(h => h.toLowerCase().includes('part number'));
            const defaultInvoiceCol = headers.find(h => h.toLowerCase().includes('invoice'));
            const defaultBoeCol = headers.find(h => h.toLowerCase().includes('boe'));
            if (defaultPartCol) setSelectedPartNumberCol(defaultPartCol);
            if (defaultInvoiceCol) setSelectedInvoiceCol(defaultInvoiceCol);
            else if (defaultBoeCol) setSelectedInvoiceCol(defaultBoeCol);


            setNotification({ message: `File loaded. Please confirm column selections and process.`, type: 'success' });

        } catch (error) {
            console.error("Error reading file:", error);
            const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
            setNotification({ message: `Failed to read file: ${errorMessage}`, type: 'error' });
        } finally {
            setIsLoading(false);
            event.target.value = "";
        }
    };
    reader.readAsArrayBuffer(file);
  };

  // Step 2: Process the data after user selections
  const processData = () => {
    const isReadyForInvoice = processingMode === 'invoice' && selectedPartNumberCol && selectedInvoiceCol && selectedQtyCol && selectedSerialCol;
    const isReadyForConsolidation = processingMode === 'consolidate' && selectedPartNumberCol && selectedQtyCol && selectedSerialCol;

    if (!isReadyForInvoice && !isReadyForConsolidation) {
        setNotification({ message: "Please select all required columns for the chosen mode.", type: 'error' });
        return;
    }

    setIsLoading(true);
    setNotification(null);

    setTimeout(() => {
        try {
            const getGroupKey = (row: DataRow) => {
                if (processingMode === 'consolidate') {
                    return row[selectedPartNumberCol];
                }
                const partKey = row[selectedPartNumberCol];
                const invoiceKey = row[selectedInvoiceCol];
                if (partKey !== undefined && invoiceKey !== undefined) {
                    return `${String(partKey)} - ${String(invoiceKey)}`;
                }
                return null;
            };

            const groupedData = parsedData.reduce((acc: {[key: string]: DataRow[]}, row: DataRow) => {
                const groupKey = getGroupKey(row);
                if (groupKey === null) return acc;
                const groupKeyStr = String(groupKey);
                if (!acc[groupKeyStr]) acc[groupKeyStr] = [];
                acc[groupKeyStr].push(row);
                return acc;
            }, {});

            const finalSheets: SheetData[] = [];
            const mismatchedSerials: { group: string; serial: string; reason: string }[] = [];

            Object.entries(groupedData).forEach(([groupKey, rows]: [string, DataRow[]]) => {
                const validRows = rows.filter(row => row[selectedQtyCol] <= 1);
                const serials = validRows.map(row => String(row[selectedSerialCol])).filter(Boolean);
                
                if (serials.length > 1) {
                    const masterSerial = serials[0];
                    const masterLength = masterSerial.length;
                    const masterSequence = getSequence(masterSerial);

                    for (let i = 1; i < serials.length; i++) {
                        const currentSerial = serials[i];
                        if (currentSerial.length !== masterLength) {
                            mismatchedSerials.push({ group: groupKey, serial: currentSerial, reason: 'Length Mismatch' });
                        } else if (getSequence(currentSerial) !== masterSequence) {
                            mismatchedSerials.push({ group: groupKey, serial: currentSerial, reason: 'Sequence Mismatch' });
                        }
                    }
                }
                
                if(serials.length > 0) {
                    finalSheets.push({
                        sheetName: groupKey,
                        data: serials.map(s => buildRowFromSerial(s))
                    });
                }
            });

            setSheetDataList(finalSheets);

            if (mismatchedSerials.length > 0) {
                const errorList = mismatchedSerials.map(e => `\n- Group [${e.group}]: Serial "${e.serial}" (${e.reason})`).join('');
                setNotification({ message: `Validation Warning! ${mismatchedSerials.length} serial(s) did not match the expected format. Details: ${errorList}`, type: 'warning' });
            } else {
                 setNotification({ message: `${finalSheets.length} sheets created successfully. All serials validated.`, type: 'success' });
            }

        } catch (error) {
            console.error("Error processing data:", error);
            setNotification({ message: "An error occurred during processing.", type: 'error' });
        } finally {
            setIsLoading(false);
        }
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

    // NEW: Validation for manual entries
    const mismatchedSerials: { serial: string; reason: string }[] = [];
    const seenSerials = new Set<string>();
    let duplicatesFound = 0;

    if (entries.length > 1) {
        const masterSerial = entries[0];
        const masterLength = masterSerial.length;
        const masterSequence = getSequence(masterSerial);

        for(const entry of entries) {
            if (seenSerials.has(entry)) {
                duplicatesFound++;
                mismatchedSerials.push({ serial: entry, reason: 'Duplicate' });
            } else {
                seenSerials.add(entry);
            }

            if (entry.length !== masterLength) {
                mismatchedSerials.push({ serial: entry, reason: 'Length Mismatch' });
            } else if (getSequence(entry) !== masterSequence) {
                mismatchedSerials.push({ serial: entry, reason: 'Sequence Mismatch' });
            }
        }
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

    if (mismatchedSerials.length > 0) {
        const errorList = mismatchedSerials.map(e => `\n- "${e.serial}" (${e.reason})`).join('');
        setNotification({ message: `Sheet "${finalSheetName}" created/updated with validation warnings: ${errorList}`, type: 'warning' });
    } else {
        setNotification({ message: `Sheet "${finalSheetName}" was added/updated successfully.`, type: 'success' });
    }
    
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
    const summaryData = sheetDataList.map(s => ({ "Sheet Name": s.sheetName, "Serial Count": s.data.length }));
    const summaryWorksheet = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(workbook, summaryWorksheet, "Summary");

    sheetDataList.forEach(({ sheetName, data }) => {
      const safeSheetName = sheetName.substring(0, 31);
      const worksheet = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName);
    });

    XLSX.writeFile(workbook, "multi_sheet_output.xlsx");
    setNotification({ message: "Excel file with summary has been exported!", type: 'success' });
  };

  const handleCopySheet = (index: number) => {
    const sheet = sheetDataList[index];
    if (!sheet || sheet.data.length === 0) {
      setNotification({ message: "Cannot copy an empty sheet.", type: 'warning' });
      return;
    }

    const headers = Object.keys(sheet.data[0]) as (keyof RowData)[];
    const headerString = headers.join("\t");
    const rowsString = sheet.data.map(row =>
        headers.map(h => row[h as keyof RowData] ?? "").join("\t")
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
    const manualSection = document.getElementById("manual-builder");
    manualSection?.scrollIntoView({ behavior: 'smooth' });
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
          style={{ whiteSpace: 'pre-wrap' }}
        >
          {notification.message}
          <button type="button" className="btn-close" onClick={() => setNotification(null)} aria-label="Close"></button>
        </div>
      )}

      <div className="card shadow mb-4">
        <div className="card-body">
          <h2 className="card-title">1. Automated Processing from File</h2>
          <p className="text-muted">Upload your file, choose a mode, map the columns, and process.</p>
          <div className="mb-3">
            <label htmlFor="file-upload" className="form-label fw-bold">Step 1: Upload Your File</label>
            <input id="file-upload" className="form-control" type="file" accept=".xlsx, .xls, .csv" onChange={handleFileChange} disabled={isLoading} />
          </div>
          
          {isLoading && <div className="text-primary fw-bold">Loading File...</div>}

          {fileHeaders.length > 0 && !isLoading && (
            <div className="card bg-light p-3">
              <h5 className="card-title">Step 2: Choose Processing Mode</h5>
              <div className="form-check">
                <input className="form-check-input" type="radio" name="mode" id="invoiceMode" checked={processingMode === 'invoice'} onChange={() => setProcessingMode('invoice')} />
                <label className="form-check-label fw-bold" htmlFor="invoiceMode">Separate by Part Number & Invoice/BOE</label>
              </div>
              <div className="form-check mb-3">
                <input className="form-check-input" type="radio" name="mode" id="consolidateMode" checked={processingMode === 'consolidate'} onChange={() => setProcessingMode('consolidate')} />
                <label className="form-check-label fw-bold" htmlFor="consolidateMode">Consolidate all Serials by Part Number only</label>
              </div>

              <h5 className="card-title mt-2">Step 3: Map Your Columns</h5>
              <div className="row g-3 mb-3">
                  <div className="col-md-4">
                      <label className="form-label fw-bold">Part Number Column</label>
                      <select className="form-select" value={selectedPartNumberCol} onChange={e => setSelectedPartNumberCol(e.target.value)}>
                          <option value="" disabled>-- Select --</option>
                          {fileHeaders.map((h, i) => <option key={`part-${i}`} value={h}>{h}</option>)}
                      </select>
                  </div>
                  {processingMode === 'invoice' && (
                    <div className="col-md-4">
                        <label className="form-label fw-bold">Invoice/BOE Column</label>
                        <select className="form-select" value={selectedInvoiceCol} onChange={e => setSelectedInvoiceCol(e.target.value)}>
                            <option value="" disabled>-- Select --</option>
                            {fileHeaders.map((h, i) => <option key={`inv-${i}`} value={h}>{h}</option>)}
                        </select>
                    </div>
                  )}
                  <div className="col-md-4">
                      <label className="form-label fw-bold">Quantity Column</label>
                      <select className="form-select" value={selectedQtyCol} onChange={e => setSelectedQtyCol(e.target.value)}>
                          <option value="" disabled>-- Select --</option>
                          {fileHeaders.map((h, i) => <option key={`qty-${i}`} value={h}>{h}</option>)}
                      </select>
                  </div>
                  <div className="col-md-4">
                      <label className="form-label fw-bold">Serial Number Column</label>
                      <select className="form-select" value={selectedSerialCol} onChange={e => setSelectedSerialCol(e.target.value)}>
                          <option value="" disabled>-- Select --</option>
                          {fileHeaders.map((h, i) => <option key={`sn-${i}`} value={h}>{h}</option>)}
                      </select>
                  </div>
              </div>
              
              <button className="btn btn-primary w-100" onClick={processData}>
                {isLoading ? 'Processing...' : 'Process File'}
              </button>
            </div>
          )}
        </div>
      </div>

      {sheetDataList.length > 0 && (
        <div className="card shadow mt-4">
          <div className="card-body">
            <div className="d-flex justify-content-between align-items-center mb-3"><h2 className="card-title mb-0">Generated Sheets</h2><button className="btn btn-success" onClick={handleExport}>📤 Export All w/ Summary</button></div>
            <ul className="list-group">
              {sheetDataList.map((s, i) => (
                <li key={i} className="list-group-item d-flex justify-content-between align-items-center">
                  <span style={{ wordBreak: 'break-all' }}><strong>{s.sheetName}</strong> – {s.data.length} rows</span>
                  <div className="btn-group">
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

      <div id="manual-builder" className="card shadow mt-4">
        <div className="card-body">
            <h2 className="card-title">2. Manual Sheet Builder</h2>
            <p className="text-muted">Paste data to add or edit sheets manually. Data is automatically validated.</p>
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
    </div>
  );
}