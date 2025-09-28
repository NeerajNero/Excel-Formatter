import * as XLSX from "xlsx";
import React, { useMemo, useState, useRef } from "react";

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
  type: "success" | "error" | "warning";
}

// A type definition for a row of data from the uploaded file
type DataRow = { [key: string]: any };

type MismatchedSerial = { group: string; serial: string; reason: string };

// Helper: copy text with navigator.clipboard or hidden textarea fallback
async function copyTextSafe(text: string) {
  try {
    if (navigator.clipboard && (window as any).isSecureContext) {
      await navigator.clipboard.writeText(text);
      return;
    }
  } catch (_) {}
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.setAttribute("readonly", "");
  ta.style.position = "absolute";
  ta.style.left = "-99999px";
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand("copy");
  } finally {
    document.body.removeChild(ta);
  }
} // Clipboard secure-context requirement + fallback pattern [web:30][web:88].

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
}); // Output row builder for export [web:21].

// Helper function to get the character type pattern of a string
const getSequence = (s: string) =>
  s.split("").map(ch => (isNaN(parseInt(ch)) ? "L" : "N")).join(""); // L/N mask for serial validation [web:12].

// Normalize quantity values robustly
const normalizeQty = (q: any) => {
  if (typeof q === "number") return q;
  const v = String(q ?? "").trim();
  if (v === "") return NaN;
  return Number(v);
}; // Trimmed Number parsing avoids NaN pitfalls for "1 " and similar values [web:114][web:117].

export default function MultiSheetBuilder() {
  // --- STATE MANAGEMENT ---
  const [sheetDataList, setSheetDataList] = useState<SheetData[]>([]); // Preserve across uploads [web:49].

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
  const [processingMode, setProcessingMode] = useState<"invoice" | "consolidate">("invoice");
  const [validateSerials, setValidateSerials] = useState<boolean>(true);
  const [mismatchedSerials, setMismatchedSerials] = useState<MismatchedSerial[]>([]);
  const [showMismatchedModal, setShowMismatchedModal] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Derived readiness for process button
  const isReadyForInvoice = useMemo(
    () => processingMode === "invoice" && !!selectedPartNumberCol && !!selectedInvoiceCol && !!selectedQtyCol && !!selectedSerialCol,
    [processingMode, selectedPartNumberCol, selectedInvoiceCol, selectedQtyCol, selectedSerialCol]
  ); // Basic mapping readiness for invoice mode [web:110].

  const isReadyForConsolidation = useMemo(
    () => processingMode === "consolidate" && !!selectedPartNumberCol && !!selectedQtyCol && !!selectedSerialCol,
    [processingMode, selectedPartNumberCol, selectedQtyCol, selectedSerialCol]
  ); // Basic mapping readiness for consolidation [web:110].

  const processButtonDisabled = isLoading || parsedData.length === 0 || (!isReadyForInvoice && !isReadyForConsolidation); // Require data and mappings before processing [web:17][web:110].

  // --- FILE PROCESSING LOGIC ---
  const parseSheet = (wb: XLSX.WorkBook, sheetNameToParse: string) => {
    try {
      const worksheet = wb.Sheets[sheetNameToParse];
      if (!worksheet) throw new Error(`Sheet "${sheetNameToParse}" not found in workbook.`);

      const dataAsArray = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false }) as any[][]; // Skip blank rows [web:8][web:142].
      const nonEmptyRows = dataAsArray.filter(
        (row) => row && row.length > 0 && row.some((cell) => cell !== null && cell !== "")
      );

      if (nonEmptyRows.length < 2)
        throw new Error("Selected sheet appears to be empty or has no data rows.");

      // Disambiguate duplicate/blank headers
      const rawHeaders: any[] = nonEmptyRows[0];
      const seen = new Map<string, number>();
      const uniqueHeaders = rawHeaders.map((h) => {
        let name = String(h ?? "").trim();
        if (name === "") name = "__EMPTY";
        const count = seen.get(name) ?? 0;
        seen.set(name, count + 1);
        return count === 0 ? name : `${name}_${count}`;
      }); // Duplicate header disambiguation [web:106][web:149].

      const dataRows = nonEmptyRows.slice(1);

      const jsonData: DataRow[] = dataRows.map((rowArray) => {
        const rowObject: DataRow = {};
        uniqueHeaders.forEach((header, index) => {
          const cell = rowArray[index];
          rowObject[header] = (cell === undefined || cell === null) ? "" : cell;
        });
        return rowObject;
      }); // Ensure keys exist even when first value is blank [web:107][web:148].

      setFileHeaders(uniqueHeaders);
      setParsedData(jsonData);

      loadMapping(uniqueHeaders); // Attempt to load and apply saved mapping [web:21].

      setNotification({
        message: `Sheet "${sheetNameToParse}" loaded. Please confirm column selections.`,
        type: "success",
      });
    } catch (error) {
      console.error("Error parsing sheet:", error);
      const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
      setNotification({ message: `Failed to parse sheet: ${errorMessage}`, type: "error" });
    }
  }; // Import pipeline with header disambiguation [web:17][web:8].

  const processUploadedFile = (file: File) => {
    setIsLoading(true);
    setNotification(null);

    // Preserve previously created sheets; reset preview-only state
    setFileHeaders([]);
    setParsedData([]);

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
          setNotification({
            message: `Workbook loaded with ${names.length} sheets. Please select a sheet to process.`,
            type: "success",
          });
        }
      } catch (error) {
        console.error("Error reading file:", error);
        const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
        setNotification({ message: `Failed to read file: ${errorMessage}`, type: "error" });
      } finally {
        setIsLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  }; // Preserve accumulated results across uploads [web:49].

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) processUploadedFile(file);
    if (event.target) event.target.value = "";
  };

  const handleDragEnter = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); };
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); e.stopPropagation(); };
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault(); e.stopPropagation(); setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) processUploadedFile(file);
  };

  const verifyMappingsExistInHeaders = () => {
    const needed = processingMode === "invoice"
      ? [selectedPartNumberCol, selectedInvoiceCol, selectedQtyCol, selectedSerialCol]
      : [selectedPartNumberCol, selectedQtyCol, selectedSerialCol];

    const missing = needed.filter(h => !fileHeaders.includes(h));
    if (missing.length) {
      setNotification({ message: `Column(s) not found in file: ${missing.join(", ")}`, type: "error" });
      return false;
    }
    return true;
  }; // Guard mapping vs headers [web:103][web:21].

  // Unified processing: consumes current parsedData state and mappings
  const processData = async () => {
    const ready = (processingMode === "invoice" ? isReadyForInvoice : isReadyForConsolidation);
    if (!ready) {
      setNotification({ message: "Please select all required columns for the chosen mode.", type: "error" });
      return;
    }
    if (parsedData.length === 0) {
      setNotification({ message: "No data rows to process.", type: "error" });
      return;
    }
    if (!verifyMappingsExistInHeaders()) return;

    setIsLoading(true);
    setNotification(null);
    setShowMismatchedModal(false);
    setMismatchedSerials([]);

    try {
      // Build groups with required-field guards
      const groupedData = parsedData.reduce((acc: Record<string, DataRow[]>, row: DataRow) => {
        const part = row[selectedPartNumberCol];
        const inv = processingMode === "invoice" ? row[selectedInvoiceCol] : undefined;
        const serial = row[selectedSerialCol];
        if (part == null || String(part).trim() === "") return acc;
        if (serial == null || String(serial).trim() === "") return acc;
        if (processingMode === "invoice" && (inv == null || String(inv).trim() === "")) return acc;
        const groupKey = processingMode === "invoice"
          ? `${String(part)} - ${String(inv)}`
          : String(part);
        if (!acc[groupKey]) acc[groupKey] = [];
        acc[groupKey].push(row);
        return acc;
      }, {});

      const finalSheets: SheetData[] = [];
      const validationErrors: MismatchedSerial[] = [];

      Object.entries(groupedData).forEach(([groupKey, rows]) => {
        const validRows = rows.filter((row) => {
          const qn = normalizeQty(row[selectedQtyCol]);
          return !Number.isNaN(qn) && qn <= 1;
        });
        const serials = validRows.map(row => String(row[selectedSerialCol])).filter(v => v.trim().length > 0);

        if (serials.length > 0) {
          if (validateSerials && serials.length > 1) {
            const masterSerial = serials[0];
            const masterLength = masterSerial.length;
            const masterSequence = getSequence(masterSerial);
            for (let i = 1; i < serials.length; i++) {
              const currentSerial = serials[i];
              if (currentSerial.length !== masterLength) {
                validationErrors.push({ group: groupKey, serial: currentSerial, reason: "Length Mismatch" });
              } else if (getSequence(currentSerial) !== masterSequence) {
                validationErrors.push({ group: groupKey, serial: currentSerial, reason: "Sequence Mismatch" });
              }
            }
          }
          finalSheets.push({ sheetName: groupKey, data: serials.map(s => buildRowFromSerial(s)) });
        }
      });

      // Compute uniqued names synchronously based on latest state and update state
      let uniqued: SheetData[] = [];
      setSheetDataList(prev => {
        const existingNames = new Set(prev.map(s => s.sheetName));
        uniqued = finalSheets.map(s => {
          let name = s.sheetName;
          let suffix = 1;
          while (existingNames.has(name)) name = `${s.sheetName} (${suffix++})`;
          existingNames.add(name);
          return { ...s, sheetName: name };
        });
        return [...prev, ...uniqued];
      }); // Append with duplicate name auto-suffix [web:52][web:45].

      // Clipboard copies after state update (side effect)
      for (const { data } of uniqued) {
        if (!data || data.length === 0) continue;
        const headers = Object.keys(data[0]) as (keyof RowData)[];
        const headerString = headers.join("\t");
        const rowsString = data.map(row => headers.map(h => row[h] ?? "").join("\t")).join("\n");
        try { await copyTextSafe([headerString, rowsString].join("\n")); } catch {}
      } // Avoid async in setState; perform afterwards [web:84][web:30].

      if (validationErrors.length > 0) {
        setMismatchedSerials(validationErrors);
        setNotification({ message: `Processing complete with ${validationErrors.length} validation warning(s).`, type: "warning" });
      } else {
        setNotification({ message: `${finalSheets.length} sheets created successfully. All serials validated.`, type: "success" });
      }
    } catch (error) {
      console.error("Error processing data:", error);
      setNotification({ message: "An error occurred during processing.", type: "error" });
    } finally {
      setIsLoading(false);
    }
  }; // Robust grouping, sync state update, async clipboard after [web:45][web:84].

  const saveMapping = () => {
    const mapping = {
      part: selectedPartNumberCol,
      invoice: selectedInvoiceCol,
      qty: selectedQtyCol,
      serial: selectedSerialCol,
    };
    localStorage.setItem("excelAppMappings", JSON.stringify(mapping));
    setNotification({ message: "Column mapping saved!", type: "success" });
  }; // Persist mapping [web:21].

  const loadMapping = (currentHeaders: string[]) => {
    const saved = localStorage.getItem("excelAppMappings");
    if (saved) {
      const mapping = JSON.parse(saved);
      if (currentHeaders.includes(mapping.part)) setSelectedPartNumberCol(mapping.part);
      if (currentHeaders.includes(mapping.invoice)) setSelectedInvoiceCol(mapping.invoice);
      if (currentHeaders.includes(mapping.qty)) setSelectedQtyCol(mapping.qty);
      if (currentHeaders.includes(mapping.serial)) setSelectedSerialCol(mapping.serial);
      setNotification({ message: "Saved column mapping loaded!", type: "success" });
    }
  }; // Load mapping if compatible [web:21].

  const buildSheetDataFromText = (entries: string[]): RowData[] => {
    return entries.map((value): RowData => ({ ...buildRowFromSerial(lotOnlyMode ? "" : value), "Lot No.": lotOnlyMode ? value : "" }));
  }; // Manual raw-add shaping [web:21].

  // Build parsedData rows for manual entries that match current column mapping
  const buildParsedRowsFromManual = (entries: string[]): DataRow[] => {
    // 1) Require mappings needed by the chosen mode
    if (processingMode === "invoice") {
      if (!selectedPartNumberCol || !selectedInvoiceCol || !selectedQtyCol || !selectedSerialCol) {
        throw new Error("Please map Part, Invoice/BOE, Quantity, and Serial before processing manual entries.");
      }
    } else {
      if (!selectedPartNumberCol || !selectedQtyCol || !selectedSerialCol) {
        throw new Error("Please map Part, Quantity, and Serial before processing manual entries.");
      }
    }

    // 2) Verify mapping keys exist in current headers
    const needed = processingMode === "invoice"
      ? [selectedPartNumberCol, selectedInvoiceCol, selectedQtyCol, selectedSerialCol]
      : [selectedPartNumberCol, selectedQtyCol, selectedSerialCol];
    const missing = needed.filter(h => !fileHeaders.includes(h));
    if (missing.length > 0) {
      throw new Error(`Column(s) not found in current file: ${missing.join(", ")}`);
    } // Prevent undefined lookups [web:103][web:21].

    // 3) Interpret each line as a single item; force quantity 1
    const partValue = sheetName?.trim() || "Manual";
    const invoiceValue = processingMode === "invoice" ? "Manual" : undefined;

    return entries
      .map(v => String(v ?? "").trim())
      .filter(v => v.length > 0)
      .map((value) => {
        const row: DataRow = {};
        row[selectedPartNumberCol] = partValue;
        if (processingMode === "invoice") row[selectedInvoiceCol] = invoiceValue;
        row[selectedQtyCol] = 1;

        // 4) Always populate the serial column so grouping/validation can run
        row[selectedSerialCol] = value;

        return row;
      });
  }; // Manual rows ready for unified pipeline [web:17].

  // Unified runner so manual can reuse the same validation/grouping pipeline
  const runProcessingPipeline = (sourceParsedData: DataRow[]) => {
    setIsLoading(true);
    setNotification(null);
    setShowMismatchedModal(false);
    setMismatchedSerials([]);

    const prev = parsedData;
    setParsedData(sourceParsedData);

    // Defer processData to next tick to allow state to commit
    setTimeout(() => {
      processData().finally(() => {
        // Restore original parsedData after processing for preview consistency
        setParsedData(prev);
      });
    }, 0);
  }; // Defer avoids stale state reads [web:45].

  // UPDATED: Raw manual add now also copies to clipboard
  const handleAddOrUpdateSheet = async () => {
    const lines = rawText.split("\n").filter(line => line.trim() !== "");
    const slicedLines = hasHeader ? lines.slice(1) : lines;
    const entries = slicedLines
      .map(line => (line.includes("\t") ? line.split("\t") : line.split(","))[columnIndex]?.trim())
      .filter(Boolean) as string[];

    if (entries.length === 0) {
      setNotification({ message: "No valid entries found in the text area.", type: "error" });
      return;
    }

    const finalSheetName = sheetName.trim() || `Sheet ${sheetDataList.length + 1}`;
    const newSheet: SheetData = { sheetName: finalSheetName, data: buildSheetDataFromText(entries) };

    if (editingSheetIndex !== null) {
      const updatedList = [...sheetDataList];
      updatedList[editingSheetIndex] = newSheet;
      setSheetDataList(updatedList);
    } else {
      setSheetDataList(prev => [...prev, newSheet]);
    }

    // Build TSV and copy to clipboard after requesting state update
    try {
      const headers = Object.keys(newSheet.data[0] ?? {}) as (keyof RowData)[];
      const headerString = headers.join("\t");
      const rowsString = newSheet.data
        .map(row => headers.map(h => row[h] ?? "").join("\t"))
        .join("\n");
      const clipboardText = [headerString, rowsString].join("\n");
      await copyTextSafe(clipboardText);
      setNotification({ message: `Sheet "${finalSheetName}" added/updated and copied to clipboard.`, type: "success" });
    } catch {
      setNotification({ message: `Sheet "${finalSheetName}" added/updated, but copy failed.`, type: "warning" });
    }

    setRawText("");
    setSheetName("");
    setEditingSheetIndex(null);
  }; // New: raw manual add auto-copies like validated flow [web:30][web:88].

  // Process manual entries through the same pipeline with validation and auto-copy
  const handleProcessManual = () => {
    const lines = rawText.split("\n").filter(line => line.trim() !== "");
    const slicedLines = hasHeader ? lines.slice(1) : lines;
    const entries = slicedLines
      .map(line => (line.includes("\t") ? line.split("\t") : line.split(","))[columnIndex]?.trim())
      .filter(Boolean) as string[];

    if (entries.length === 0) {
      setNotification({ message: "No valid entries found in the text area.", type: "error" });
      return;
    }

    try {
      const manualParsed = buildParsedRowsFromManual(entries);
      runProcessingPipeline(manualParsed);
      setRawText("");
      setSheetName("");
      setEditingSheetIndex(null);
    } catch (err: any) {
      setNotification({ message: err?.message || "Failed to process manual entries.", type: "error" });
    }
  }; // Manual validated path auto-copies via unified pipeline [web:17].

  const handleExport = () => {
    if (sheetDataList.length === 0) {
      setNotification({ message: "No sheets to export.", type: "error" });
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
    setNotification({ message: "Excel file with summary has been exported!", type: "success" });
  }; // Enforce Excel 31-char sheet-name limit on export [web:40].

  const handleCopySheet = async (index: number) => {
    const sheet = sheetDataList[index];
    if (!sheet || sheet.data.length === 0) {
      setNotification({ message: "Cannot copy an empty sheet.", type: "warning" });
      return;
    }

    const headers = Object.keys(sheet.data[0]) as (keyof RowData)[];
    const headerString = headers.join("\t");
    const rowsString = sheet.data.map(row => headers.map(h => row[h as keyof RowData] ?? "").join("\t")).join("\n");

    const clipboardText = [headerString, rowsString].join("\n");
    try {
      await copyTextSafe(clipboardText);
      setNotification({ message: `Copied data for "${sheet.sheetName}" to clipboard.`, type: "success" });
    } catch (err) {
      setNotification({ message: "Failed to copy data.", type: "error" });
      console.error("Clipboard error:", err);
    }
  }; // Manual copy helper [web:30].

  const handleEditSheet = (index: number) => {
    const sheet = sheetDataList[index];
    if (!sheet) return;
    const isLot = sheet.data.some(d => d["Lot No."] !== "");
    setRawText(sheet.data.map(row => (isLot ? row["Lot No."] : row["Serial No."])).join("\n"));
    setSheetName(sheet.sheetName);
    setLotOnlyMode(isLot);
    setEditingSheetIndex(index);
    const manualSection = document.getElementById("manual-builder");
    manualSection?.scrollIntoView({ behavior: "smooth" });
  }; // Edit flow [web:52].

  const cancelEdit = () => {
    setEditingSheetIndex(null);
    setRawText("");
    setSheetName("");
  }; // Cancel edit [web:52].

  const handleDeleteSheet = (index: number) => {
    if (window.confirm(`Delete "${sheetDataList[index].sheetName}"?`)) {
      const sName = sheetDataList[index].sheetName;
      setSheetDataList(prev => prev.filter((_, i) => i !== index));
      setNotification({ message: `Sheet "${sName}" deleted.`, type: "warning" });
    }
  }; // Delete [web:52].

  // --- JSX / RENDER ---
  const dropZoneStyle: React.CSSProperties = { border: "2px dashed #ccc", borderRadius: "8px", padding: "2rem", textAlign: "center", cursor: "pointer", transition: "border-color 0.2s, background-color 0.2s" };
  const dropZoneDraggingStyle: React.CSSProperties = { borderColor: "#0d6efd", backgroundColor: "#f0f8ff" };

  return (
    <div className="container py-5">
      {notification && (
        <div className={`alert alert-${notification.type === "error" ? "danger" : notification.type} alert-dismissible fade show`} role="alert">
          {notification.message}
          {mismatchedSerials.length > 0 && <button className="btn btn-sm btn-light ms-3" onClick={() => setShowMismatchedModal(true)}>Show Details</button>}
          <button type="button" className="btn-close" onClick={() => setNotification(null)} aria-label="Close"></button>
        </div>
      )}

      {showMismatchedModal && (
        <div className="modal show" style={{ display: "block", backgroundColor: "rgba(0,0,0,0.5)" }} tabIndex={-1}>
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
              <h2 className="card-title">1. Unified Processing</h2>

              <div
                style={isDragging ? { ...dropZoneStyle, ...dropZoneDraggingStyle } : dropZoneStyle}
                onDragEnter={e => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); }}
                onDragLeave={e => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); }}
                onDragOver={e => { e.preventDefault(); e.stopPropagation(); }}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <input ref={fileInputRef} id="file-upload" type="file" accept=".xlsx, .xls, .csv" onChange={handleFileChange} disabled={isLoading} style={{ display: "none" }} />
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
                  <div className="table-responsive" style={{ maxHeight: "150px" }}>
                    <table className="table table-sm table-bordered table-striped">
                      <thead className="table-dark"><tr>{fileHeaders.map((h, i) => <th key={i}>{h}</th>)}</tr></thead>
                      <tbody>{parsedData.slice(0, 5).map((row, i) => <tr key={i}>{fileHeaders.map((h, j) => <td key={j}>{row[h]}</td>)}</tr>)}</tbody>
                    </table>
                  </div>
                </div>
              )}

              {fileHeaders.length > 0 && !isLoading && (
                <div className="card bg-light p-3 mt-3">
                  <h5 className="card-title">Step 1: Choose Mode</h5>
                  <div className="form-check"><input className="form-check-input" type="radio" name="mode" id="invoiceMode" checked={processingMode === "invoice"} onChange={() => setProcessingMode("invoice")} /><label className="form-check-label fw-bold" htmlFor="invoiceMode">Separate by Part Number & Invoice/BOE</label></div>
                  <div className="form-check mb-3"><input className="form-check-input" type="radio" name="mode" id="consolidateMode" checked={processingMode === "consolidate"} onChange={() => setProcessingMode("consolidate")} /><label className="form-check-label fw-bold" htmlFor="consolidateMode">Consolidate all Serials by Part Number only</label></div>

                  <div className="d-flex justify-content-between align-items-center mt-2">
                    <h5 className="card-title mb-0">Step 2: Map Columns</h5>
                    <div><button className="btn btn-sm btn-outline-secondary me-2" onClick={() => saveMapping()}>Save</button><button className="btn btn-sm btn-outline-secondary" onClick={() => loadMapping(fileHeaders)}>Load</button></div>
                  </div>
                  <div className="row g-3 my-2">
                    <div className="col-md-6"><label className="form-label fw-bold">Part Number Column</label><select className="form-select" value={selectedPartNumberCol} onChange={e => setSelectedPartNumberCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`part-${i}`} value={h}>{h}</option>)}</select></div>
                    {processingMode === "invoice" && (<div className="col-md-6"><label className="form-label fw-bold">Invoice/BOE Column</label><select className="form-select" value={selectedInvoiceCol} onChange={e => setSelectedInvoiceCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`inv-${i}`} value={h}>{h}</option>)}</select></div>)}
                    <div className="col-md-6"><label className="form-label fw-bold">Quantity Column</label><select className="form-select" value={selectedQtyCol} onChange={e => setSelectedQtyCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`qty-${i}`} value={h}>{h}</option>)}</select></div>
                    <div className="col-md-6"><label className="form-label fw-bold">Serial Number Column</label><select className="form-select" value={selectedSerialCol} onChange={e => setSelectedSerialCol(e.target.value)}><option value="" disabled>-- Select --</option>{fileHeaders.map((h, i) => <option key={`sn-${i}`} value={h}>{h}</option>)}</select></div>
                  </div>

                  <div className="form-check form-switch mb-3"><input className="form-check-input" type="checkbox" role="switch" id="validateSerialsSwitch" checked={validateSerials} onChange={(e) => setValidateSerials(e.target.checked)} /><label className="form-check-label fw-bold" htmlFor="validateSerialsSwitch">Validate serial number sequence and length</label></div>

                  <button className="btn btn-primary w-100" disabled={processButtonDisabled} onClick={processData}>
                    {isLoading ? "Processing..." : "Process File"}
                  </button>
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
                  <ul className="list-group" style={{ maxHeight: "40vh", overflowY: "auto" }}>
                    {sheetDataList.map((s, i) => (
                      <li key={i} className="list-group-item d-flex justify-content-between align-items-center">
                        <span style={{ wordBreak: "break-all" }}><strong>{s.sheetName}</strong> – {s.data.length} rows</span>
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
                <p className="text-muted">Paste data to add or process with validation and auto-copy.</p>
                <div className="mb-3"><label className="form-label fw-bold">Paste Serial/Lot Numbers</label><textarea className="form-control" rows={5} value={rawText} onChange={(e) => setRawText(e.target.value)} placeholder="Paste rows here..."/></div>
                <div className="row g-3 mb-3">
                  <div className="col-md-6"><label className="form-label fw-bold">Sheet/Part Name</label><input className="form-control" type="text" value={sheetName} onChange={(e) => setSheetName(e.target.value)} placeholder="Optional (used as Part Number in manual processing)"/></div>
                  <div className="col-md-6"><label className="form-label fw-bold">Column Index</label><input className="form-control" type="number" value={columnIndex} onChange={(e) => setColumnIndex(Number(e.target.value))}/></div>
                  <div className="col-12 d-flex justify-content-start">
                    <div className="form-check me-4"><input className="form-check-input" type="checkbox" checked={lotOnlyMode} onChange={() => setLotOnlyMode((p) => !p)} id="lotOnlyCheck"/><label className="form-check-label fw-bold" htmlFor="lotOnlyCheck">Lot Mode</label></div>
                    <div className="form-check"><input className="form-check-input" type="checkbox" checked={hasHeader} onChange={() => setHasHeader((p) => !p)} id="hasHeaderCheck"/><label className="form-check-label fw-bold" htmlFor="hasHeaderCheck">Has Header</label></div>
                  </div>
                </div>
                <div className="d-flex gap-2">
                  <button className="btn btn-primary w-100" onClick={handleProcessManual}>⚙️ Process Manual (validated)</button>
                  <button className="btn btn-outline-primary w-100" onClick={handleAddOrUpdateSheet}>{editingSheetIndex !== null ? "💾 Update Sheet (raw)" : "➕ Add Sheet (raw)"}</button>
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
