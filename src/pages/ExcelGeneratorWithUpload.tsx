import * as XLSX from "xlsx";
import React, { useState, useRef } from "react";

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

type MismatchedSerial = { group: string; serial: string; reason: string };

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
  const [isDragging, setIsDragging] = useState<boolean>(false);

  // States for the interactive column selection workflow
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [fileHeaders, setFileHeaders] = useState<string[]>([]);
  const [parsedData, setParsedData] = useState<DataRow[]>([]);
  const [selectedPartNumberCol, setSelectedPartNumberCol] = useState<string>("");
  const [selectedInvoiceCol, setSelectedInvoiceCol] = useState<string>("");
  const [selectedQtyCol, setSelectedQtyCol] = useState<string>("");
  const [selectedSerialCol, setSelectedSerialCol] = useState<string>("");

  // States for processing mode and validation
  const [processingMode, setProcessingMode] = useState<'invoice' | 'consolidate'>('invoice');
  const [validateSerials, setValidateSerials] = useState<boolean>(true);
  const [mismatchedSerials, setMismatchedSerials] = useState<MismatchedSerial[]>([]);
  const [showMismatchedModal, setShowMismatchedModal] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- FILE PROCESSING LOGIC ---
  const parseSheet = (wb: XLSX.WorkBook, sheetNameToParse: string) => {
    try {
        const worksheet = wb.Sheets[sheetNameToParse];
        if (!worksheet) throw new Error(`Sheet "${sheetNameToParse}" not found in workbook.`);

        const dataAsArray = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        const nonEmptyRows = dataAsArray.filter(row => row && row.length > 0 && row.some(cell => cell !== null && cell !== ''));
        
        if (nonEmptyRows.length < 2) throw new Error("Selected sheet appears to be empty or has no data rows.");

        const headers: string[] = nonEmptyRows[0];
        const dataRows = nonEmptyRows.slice(1);

        const jsonData: DataRow[] = dataRows.map(rowArray => {
            const rowObject: DataRow = {};
            headers.forEach((header, index) => {
                if (typeof header === 'string') rowObject[header] = rowArray[index];
            });
            return rowObject;
        });

        setFileHeaders(headers);
        setParsedData(jsonData);

        loadMapping(headers); // Attempt to load and apply saved mapping

        setNotification({ message: `Sheet "${sheetNameToParse}" loaded. Please confirm column selections.`, type: 'success' });
    } catch (error) {
        console.error("Error parsing sheet:", error);
        const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
        setNotification({ message: `Failed to parse sheet: ${errorMessage}`, type: 'error' });
    }
  };

  const processUploadedFile = (file: File) => {
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
    setSheetNames([]);
    setWorkbook(null);

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const wb = XLSX.read(data, { type: "array" });
            setWorkbook(wb);
            const names = wb.SheetNames;
            setSheetNames(names);

            if (names.length === 1) {
                parseSheet(wb, names[0]);
            } else {
                setNotification({ message: `Workbook loaded with ${names.length} sheets. Please select a sheet to process.`, type: 'success' });
            }
        } catch (error) {
            console.error("Error reading file:", error);
            const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
            setNotification({ message: `Failed to read file: ${errorMessage}`, type: 'error' });
        } finally {
            setIsLoading(false);
        }
    };
    reader.readAsArrayBuffer(file);
  };
  
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) processUploadedFile(file);
    if (event.target) event.target.value = "";
  };
  
  const handleDragEnter = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); };
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); };
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) processUploadedFile(file);
  };

  const processData = () => {
    const isReadyForInvoice = processingMode === 'invoice' && selectedPartNumberCol && selectedInvoiceCol && selectedQtyCol && selectedSerialCol;
    const isReadyForConsolidation = processingMode === 'consolidate' && selectedPartNumberCol && selectedQtyCol && selectedSerialCol;

    if (!isReadyForInvoice && !isReadyForConsolidation) {
        setNotification({ message: "Please select all required columns for the chosen mode.", type: 'error' });
        return;
    }

    setIsLoading(true);
    setNotification(null);
    setShowMismatchedModal(false);
    setMismatchedSerials([]);

    setTimeout(() => {
        try {
            const getGroupKey = (row: DataRow) => {
                if (processingMode === 'consolidate') return row[selectedPartNumberCol];
                const partKey = row[selectedPartNumberCol];
                const invoiceKey = row[selectedInvoiceCol];
                if (partKey !== undefined && invoiceKey !== undefined) return `${String(partKey)} - ${String(invoiceKey)}`;
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
            const validationErrors: MismatchedSerial[] = [];

            Object.entries(groupedData).forEach(([groupKey, rows]: [string, DataRow[]]) => {
                const validRows = rows.filter(row => row[selectedQtyCol] <= 1);
                const serials = validRows.map(row => String(row[selectedSerialCol])).filter(Boolean);
                
                if (serials.length > 0) {
                    if (validateSerials && serials.length > 1) {
                        const masterSerial = serials[0];
                        const masterLength = masterSerial.length;
                        const masterSequence = getSequence(masterSerial);
                        for (let i = 1; i < serials.length; i++) {
                            const currentSerial = serials[i];
                            if (currentSerial.length !== masterLength) {
                                validationErrors.push({ group: groupKey, serial: currentSerial, reason: 'Length Mismatch' });
                            } else if (getSequence(currentSerial) !== masterSequence) {
                                validationErrors.push({ group: groupKey, serial: currentSerial, reason: 'Sequence Mismatch' });
                            }
                        }
                    }
                    finalSheets.push({ sheetName: groupKey, data: serials.map(s => buildRowFromSerial(s)) });
                }
            });

            setSheetDataList(finalSheets);

            if (validationErrors.length > 0) {
                setMismatchedSerials(validationErrors);
                setNotification({ message: `Processing complete with ${validationErrors.length} validation warning(s).`, type: 'warning' });
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

  const saveMapping = () => {
    const mapping = {
      part: selectedPartNumberCol,
      invoice: selectedInvoiceCol,
      qty: selectedQtyCol,
      serial: selectedSerialCol
    };
    localStorage.setItem('excelAppMappings', JSON.stringify(mapping));
    setNotification({ message: 'Column mapping saved!', type: 'success' });
  };

  const loadMapping = (currentHeaders: string[]) => {
    const saved = localStorage.getItem('excelAppMappings');
    if (saved) {
      const mapping = JSON.parse(saved);
      if (currentHeaders.includes(mapping.part)) setSelectedPartNumberCol(mapping.part);
      if (currentHeaders.includes(mapping.invoice)) setSelectedInvoiceCol(mapping.invoice);
      if (currentHeaders.includes(mapping.qty)) setSelectedQtyCol(mapping.qty);
      if (currentHeaders.includes(mapping.serial)) setSelectedSerialCol(mapping.serial);
      setNotification({ message: 'Saved column mapping loaded!', type: 'success' });
    }
  };

  const buildSheetDataFromText = (entries: string[]): RowData[] => {
    return entries.map((value): RowData => ({ ...buildRowFromSerial(lotOnlyMode ? "" : value), "Lot No.": lotOnlyMode ? value : "" }));
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
  const dropZoneStyle: React.CSSProperties = { border: '2px dashed #ccc', borderRadius: '8px', padding: '2rem', textAlign: 'center', cursor: 'pointer', transition: 'border-color 0.2s, background-color 0.2s' };
  const dropZoneDraggingStyle: React.CSSProperties = { borderColor: '#0d6efd', backgroundColor: '#f0f8ff' };
  
  return (
    <div className="container py-5">
      {notification && (
        <div className={`alert alert-${notification.type === 'error' ? 'danger' : notification.type} alert-dismissible fade show`} role="alert">
          {notification.message}
          {mismatchedSerials.length > 0 && <button className="btn btn-sm btn-light ms-3" onClick={() => setShowMismatchedModal(true)}>Show Details</button>}
          <button type="button" className="btn-close" onClick={() => setNotification(null)} aria-label="Close"></button>
        </div>
      )}
      {showMismatchedModal && (
        <div className="modal show" style={{ display: 'block', backgroundColor: 'rgba(0,0,0,0.5)' }} tabIndex={-1}>
          <div className="modal-dialog modal-lg modal-dialog-scrollable">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title">Validation Mismatch Details</h5>
                <button type="button" className="btn-close" onClick={() => setShowMismatchedModal(false)}></button>
              </div>
              <div className="modal-body">
                <table className="table table-sm table-striped">
                    <thead><tr><th>Group</th><th>Mismatched Serial</th><th>Reason</th></tr></thead>
                    <tbody>{mismatchedSerials.map((e, i) => <tr key={i}><td>{e.group}</td><td>{e.serial}</td><td>{e.reason}</td></tr>)}</tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}
      
      <div className="row g-4">
        {/* --- LEFT COLUMN --- */}
        <div className="col-lg-6">
          <div className="card shadow h-100">
            <div className="card-body d-flex flex-column">
              <h2 className="card-title">1. Automated Processing from File</h2>
              <div 
                style={isDragging ? {...dropZoneStyle, ...dropZoneDraggingStyle} : dropZoneStyle}
                onDragEnter={handleDragEnter} onDragLeave={handleDragLeave} onDragOver={handleDragOver} onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <input ref={fileInputRef} id="file-upload" type="file" accept=".xlsx, .xls, .csv" onChange={handleFileChange} disabled={isLoading} style={{ display: 'none' }} />
                <p className="mb-0">Drag & drop file here, or click to select</p>
              </div>
              
              {isLoading && <div className="text-primary fw-bold mt-3">Loading...</div>}

              {sheetNames.length > 1 && !isLoading && (
                 <div className="mt-3">
                    <label className="form-label fw-bold">Select Sheet to Process</label>
                    <select className="form-select" onChange={(e) => workbook && parseSheet(workbook, e.target.value)}>
                        <option value="">-- Please select a sheet --</option>
                        {sheetNames.map(name => <option key={name} value={name}>{name}</option>)}
                    </select>
                </div>
              )}
              
              {parsedData.length > 0 && !isLoading && (
                <div className="mt-3">
                  <h6>Data Preview (First 5 Rows)</h6>
                  <div className="table-responsive" style={{ maxHeight: '150px' }}>
                    <table className="table table-sm table-bordered table-striped">
                      <thead className="table-dark"><tr>{fileHeaders.map((h, i) => <th key={i}>{h}</th>)}</tr></thead>
                      <tbody>{parsedData.slice(0, 5).map((row, i) => <tr key={i}>{fileHeaders.map((h, j) => <td key={j}>{row[h]}</td>)}</tr>)}</tbody>
                    </table>
                  </div>
                </div>
              )}

              {fileHeaders.length > 0 && !isLoading && (
                <div className="card bg-light p-3 mt-3">
                  <h5 className="card-title">Step 2: Choose Processing Mode</h5>
                  <div className="form-check"><input className="form-check-input" type="radio" name="mode" id="invoiceMode" checked={processingMode === 'invoice'} onChange={() => setProcessingMode('invoice')} /><label className="form-check-label fw-bold" htmlFor="invoiceMode">Separate by Part Number & Invoice/BOE</label></div>
                  <div className="form-check mb-3"><input className="form-check-input" type="radio" name="mode" id="consolidateMode" checked={processingMode === 'consolidate'} onChange={() => setProcessingMode('consolidate')} /><label className="form-check-label fw-bold" htmlFor="consolidateMode">Consolidate all Serials by Part Number only</label></div>
                  
                  <div className="d-flex justify-content-between align-items-center mt-2">
                    <h5 className="card-title mb-0">Step 3: Map Your Columns</h5>
                    <div><button className="btn btn-sm btn-outline-secondary me-2" onClick={() => saveMapping()}>Save</button><button className="btn btn-sm btn-outline-secondary" onClick={() => loadMapping(fileHeaders)}>Load</button></div>
                  </div>
                  <div className="row g-3 my-2">
                      <div className="col-md-6"><label className="form-label fw-bold">Part Number Column</label><select className="form-select" value={selectedPartNumberCol} onChange={e => setSelectedPartNumberCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`part-${i}`} value={h}>{h}</option>)}</select></div>
                      {processingMode === 'invoice' && (<div className="col-md-6"><label className="form-label fw-bold">Invoice/BOE Column</label><select className="form-select" value={selectedInvoiceCol} onChange={e => setSelectedInvoiceCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`inv-${i}`} value={h}>{h}</option>)}</select></div>)}
                      <div className="col-md-6"><label className="form-label fw-bold">Quantity Column</label><select className="form-select" value={selectedQtyCol} onChange={e => setSelectedQtyCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`qty-${i}`} value={h}>{h}</option>)}</select></div>
                      <div className="col-md-6"><label className="form-label fw-bold">Serial Number Column</label><select className="form-select" value={selectedSerialCol} onChange={e => setSelectedSerialCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`sn-${i}`} value={h}>{h}</option>)}</select></div>
                  </div>
                  <div className="form-check form-switch mb-3"><input className="form-check-input" type="checkbox" role="switch" id="validateSerialsSwitch" checked={validateSerials} onChange={(e) => setValidateSerials(e.target.checked)} /><label className="form-check-label fw-bold" htmlFor="validateSerialsSwitch">Validate serial number sequence and length</label></div>
                  <button className="btn btn-primary w-100" onClick={processData}>{isLoading ? 'Processing...' : 'Process File'}</button>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* --- RIGHT COLUMN --- */}
        <div className="col-lg-6">
          <div className="d-flex flex-column gap-4">
            {sheetDataList.length > 0 && (
              <div className="card shadow">
                <div className="card-body">
                  <div className="d-flex justify-content-between align-items-center mb-3"><h2 className="card-title mb-0">Generated Sheets</h2><button className="btn btn-success" onClick={handleExport}>📤 Export All w/ Summary</button></div>
                  <ul className="list-group" style={{ maxHeight: '40vh', overflowY: 'auto' }}>
                    {sheetDataList.map((s, i) => (
                      <li key={i} className="list-group-item d-flex justify-content-between align-items-center">
                        <span style={{ wordBreak: 'break-all' }}><strong>{s.sheetName}</strong> – {s.data.length} rows</span>
                        <div className="btn-group"><button className="btn btn-sm btn-outline-secondary" onClick={() => handleCopySheet(i)}>📋 Copy</button><button className="btn btn-sm btn-outline-primary" onClick={() => handleEditSheet(i)}>✏️ Edit</button><button className="btn btn-sm btn-outline-danger" onClick={() => handleDeleteSheet(i)}>🗑️ Delete</button></div>
                      </li>
                    ))}
                  </ul>
                </div>
              </div>
            )}

            <div id="manual-builder" className="card shadow">
              <div className="card-body">
                  <h2 className="card-title">2. Manual Sheet Builder</h2>
                  <p className="text-muted">Paste data to add or edit sheets manually.</p>
                  <div className="mb-3"><label className="form-label fw-bold">Paste Serial/Lot Numbers</label><textarea className="form-control" rows={5} value={rawText} onChange={(e) => setRawText(e.target.value)} placeholder="Paste rows here..."/></div>
                  <div className="row g-3 mb-3">
                      <div className="col-md-6"><label className="form-label fw-bold">Sheet Name</label><input className="form-control" type="text" value={sheetName} onChange={(e) => setSheetName(e.target.value)} placeholder="Optional"/></div>
                      <div className="col-md-6"><label className="form-label fw-bold">Column Index</label><input className="form-control" type="number" value={columnIndex} onChange={(e) => setColumnIndex(Number(e.target.value))}/></div>
                      <div className="col-12 d-flex justify-content-start">
                          <div className="form-check me-4"><input className="form-check-input" type="checkbox" checked={lotOnlyMode} onChange={() => setLotOnlyMode((p) => !p)} id="lotOnlyCheck"/><label className="form-check-label fw-bold" htmlFor="lotOnlyCheck">Lot Mode</label></div>
                          <div className="form-check"><input className="form-check-input" type="checkbox" checked={hasHeader} onChange={() => setHasHeader((p) => !p)} id="hasHeaderCheck"/><label className="form-check-label fw-bold" htmlFor="hasHeaderCheck">Has Header</label></div>
                      </div>
                  </div>
                  <div className="d-flex gap-2">
                      <button className="btn btn-primary w-100" onClick={handleAddOrUpdateSheet}>{editingSheetIndex !== null ? "💾 Update Sheet" : "➕ Add Sheet"}</button>
                      {editingSheetIndex !== null && (<button className="btn btn-secondary w-100" onClick={cancelEdit}>❌ Cancel Edit</button>)}
                  </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}