import { useState } from "react";

interface RawRow {
  [key: string]: string;
}

export default function ThirdExtractor() {
  const [rawText, setRawText] = useState<string>("");
  const [parsedData, setParsedData] = useState<RawRow[]>([]);

  const handleParse = () => {
    const lines = rawText.trim().split("\n");
    const headers = lines[0].split("\t");

    const rows: RawRow[] = lines.slice(1).map((line) => {
      const values = line.split("\t");
      const row: RawRow = {};
      headers.forEach((key, i) => {
        row[key.trim()] = values[i]?.trim() || "";
      });
      return row;
    });

    setParsedData(rows);
  };

  const cleanDollar = (value?: string): number => {
    return parseFloat(value?.replace(/[$,\s]/g, "").trim() || "");
  };

  const formatUSD = (num: number): string => {
    return isNaN(num) ? "" : `$ ${num.toFixed(2)}`;
  };

  const formatPlain = (num: number): string => {
    return isNaN(num) ? "" : num.toFixed(2);
  };

  const copyTemplate1 = () => {
    const header = [
      "SKU CODE", "OEM PART NO", "Variant", "DESCRIPTION",
      "REQ QTY", "QUOTE PRICE $", "Total $", "PURPOSE"
    ].join("\t");

    const rows = parsedData.map(row => {
      const qty = parseInt(row["REQ QTY"] || "");
      const unitPrice = cleanDollar(row["QUOTE PRICE"]);
      const total = qty * unitPrice;
      const purpose = row["END USER NAME"] || "";
      return [
        row["SKU CODE"],
        row["Main OEM Part No"],
        row["VARIANT CODE"],
        row["DESCRIPTION"],
        qty,
        formatUSD(unitPrice),
        formatUSD(total),
        purpose
      ].join("\t");
    });

    navigator.clipboard.writeText([header, ...rows].join("\n"));
  };

  const copyTemplate2 = () => {
    const header = [
      "PO#", "PO date", "Status", "Supplier", "Item", "Item Code",
      "Variant", "Description", "Qty", "U/P", "Total", "Purpose"
    ].join("\t");

    const rows = parsedData.map(row => {
      const qty = parseInt(row["REQ QTY"] || "");
      const unitPrice = cleanDollar(row["QUOTE PRICE"]);
      const total = qty * unitPrice;
      const purpose = row["END USER NAME"] || "";

      return [
        "",
        "",
        row["Status"] || "",
        "Armortec",
        row["SKU CODE"],
        row["ITEM CODE"],
        row["VARIANT CODE"],
        row["DESCRIPTION"],
        qty,
        formatPlain(unitPrice),
        formatPlain(total),
        purpose
      ].join("\t");
    });

    navigator.clipboard.writeText([header, ...rows].join("\n"));
  };

  return (
    <div className="container py-4">
      <h1 className="mb-4">📋 Third Extractor Tool</h1>

      <div className="mb-3">
        <label htmlFor="excelInput" className="form-label fw-bold">
          Paste Data (Tab-Separated)
        </label>
        <textarea
          id="excelInput"
          className="form-control"
          rows={8}
          placeholder="Paste your data here"
          value={rawText}
          onChange={(e) => setRawText(e.target.value)}
        />
      </div>

      <button className="btn btn-primary mb-4" onClick={handleParse}>
        Generate Templates
      </button>

      {parsedData.length > 0 && (
        <div className="row g-4">
          {/* Template 1 */}
          <div className="col-md-6">
            <div className="d-flex justify-content-between align-items-center mb-2">
              <h4>🧾 Template 1</h4>
              <button className="btn btn-sm btn-outline-secondary" onClick={copyTemplate1}>
                Copy Template 1
              </button>
            </div>
            <div className="table-responsive">
              <table className="table table-bordered table-sm table-striped">
                <thead className="table-light">
                  <tr>
                    <th>SKU CODE</th>
                    <th>OEM PART NO</th>
                    <th>Variant</th>
                    <th>DESCRIPTION</th>
                    <th>REQ QTY</th>
                    <th>QUOTE PRICE $</th>
                    <th>Total $</th>
                    <th>PURPOSE</th>
                  </tr>
                </thead>
                <tbody>
                  {parsedData.map((row, i) => {
                    const qty = parseInt(row["REQ QTY"] || "");
                    const unitPrice = cleanDollar(row["QUOTE PRICE"]);
                    const total = qty * unitPrice;
                    const purpose = row["END USER NAME"] || "";
                    return (
                      <tr key={i}>
                        <td>{row["SKU CODE"]}</td>
                        <td>{row["Main OEM Part No"]}</td>
                        <td>{row["VARIANT CODE"]}</td>
                        <td>{row["DESCRIPTION"]}</td>
                        <td>{qty}</td>
                        <td>{formatUSD(unitPrice)}</td>
                        <td>{formatUSD(total)}</td>
                        <td>{purpose}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Template 2 */}
          <div className="col-md-6">
            <div className="d-flex justify-content-between align-items-center mb-2">
              <h4>📦 Template 2</h4>
              <button className="btn btn-sm btn-outline-secondary" onClick={copyTemplate2}>
                Copy Template 2
              </button>
            </div>
            <div className="table-responsive">
              <table className="table table-bordered table-sm table-striped">
                <thead className="table-light">
                  <tr>
                    <th>PO#</th>
                    <th>PO date</th>
                    <th>Status</th>
                    <th>Supplier</th>
                    <th>Item</th>
                    <th>Item Code</th>
                    <th>Variant</th>
                    <th>Description</th>
                    <th>Qty</th>
                    <th>U/P</th>
                    <th>Total</th>
                    <th>Purpose</th>
                  </tr>
                </thead>
                <tbody>
                  {parsedData.map((row, i) => {
                    const qty = parseInt(row["REQ QTY"] || "");
                    const unitPrice = cleanDollar(row["QUOTE PRICE"]);
                    const total = qty * unitPrice;
                    const purpose = row["END USER NAME"] || "";
                    return (
                      <tr key={i}>
                        <td></td>
                        <td></td>
                        <td>{row["Status"] || ""}</td>
                        <td>Armortec</td>
                        <td>{row["SKU CODE"]}</td>
                        <td>{row["ITEM CODE"]}</td>
                        <td>{row["VARIANT CODE"]}</td>
                        <td>{row["DESCRIPTION"]}</td>
                        <td>{qty}</td>
                        <td>{formatPlain(unitPrice)}</td>
                        <td>{formatPlain(total)}</td>
                        <td>{purpose}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
