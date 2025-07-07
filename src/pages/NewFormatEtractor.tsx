import { useState } from "react";

interface RawRow {
  [key: string]: string;
}

export default function SecondExtractor() {
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

  // Utility to fetch either Buy Price or Quote Price
  const getBuyPrice = (row: RawRow): string => {
    return row["Buy Price"]?.trim() || row["Quote Price"]?.trim() || "";
  };

  const cleanNumber = (value?: string): number => {
    return parseFloat(value?.replace(/[,₹\s]/g, "").trim() || "");
  };

  const toNumber = (value?: string): string => {
    const num = cleanNumber(value);
    return isNaN(num) ? "" : num.toFixed(2);
  };

  const add5Percent = (value?: string): string => {
    const num = cleanNumber(value);
    return isNaN(num) ? "" : (num * 1.05).toFixed(2);
  };

  const calculateMargin = (cost?: string, sale?: string): string => {
    const costNum = cleanNumber(cost);
    const saleNum = cleanNumber(sale);
    if (isNaN(costNum) || isNaN(saleNum) || costNum === 0) return "";
    return (((saleNum - costNum) / costNum) * 100).toFixed(2);
  };

  const copyTable1 = () => {
    const header = [
      "Qty", "Item", "Mustek Buy Price", "Quinta 5%", "End Price", "Margin",
      "Customer", "Vendor", "PO #"
    ].join("\t");

    const rows = parsedData.map((row) => {
      return [
        row["Req Qty"] || "",
        row["SKU Code"] || "",
        toNumber(getBuyPrice(row)),
        add5Percent(getBuyPrice(row)),
        toNumber(row["Sell Price"]),
        calculateMargin(getBuyPrice(row), row["Sell Price"]),
        row["End User Name"] || "",
        row["Vendor"] || "",
        ""
      ].join("\t");
    });

    navigator.clipboard.writeText([header, ...rows].join("\n"));
  };

  const copyTable2 = () => {
    const header = [
      "PO No", "PO dt", "Customer", "Part number", "Variant/Brand", "Product type",
      "Qty", "U/P", "Total", "ETD", "BU", "Supplier"
    ].join("\t");

    const rows = parsedData.map((row) => {
      const qty = parseInt(row["Req Qty"] || "");
      const unitPrice = cleanNumber(getBuyPrice(row));
      const total = isNaN(qty) || isNaN(unitPrice) ? "" : (qty * unitPrice).toFixed(2);

      return [
        "", "", row["End User Name"] || "", row["SKU Code"] || "", "", "",
        isNaN(qty) ? "" : qty,
        isNaN(unitPrice) ? "" : unitPrice.toFixed(2),
        total, "", "", row["Vendor"] || ""
      ].join("\t");
    });

    navigator.clipboard.writeText([header, ...rows].join("\n"));
  };

  return (
    <div className="container py-4">
      <h1 className="mb-4">📋 Excel Formatter Tool</h1>

      <div className="mb-3">
        <label htmlFor="excelInput" className="form-label fw-bold">
          Paste Excel Table Data
        </label>
        <textarea
          id="excelInput"
          className="form-control"
          rows={8}
          placeholder="Paste your Excel table here (Tab-separated)"
          value={rawText}
          onChange={(e) => setRawText(e.target.value)}
        />
      </div>

      <button className="btn btn-primary mb-4" onClick={handleParse}>
        Generate Tables
      </button>

      {parsedData.length > 0 && (
        <div className="row g-4">
          {/* Table 1 */}
          <div className="col-md-6">
            <div className="d-flex justify-content-between align-items-center mb-2">
              <h4>💰 Price Calculation Table</h4>
              <button className="btn btn-sm btn-outline-secondary" onClick={copyTable1}>
                Copy Table
              </button>
            </div>
            <div className="table-responsive">
              <table className="table table-bordered table-sm table-striped">
                <thead className="table-light">
                  <tr>
                    <th>Qty</th>
                    <th>Item</th>
                    <th>Mustek Buy Price</th>
                    <th>Quinta 5%</th>
                    <th>End Price</th>
                    <th>Margin</th>
                    <th>Customer</th>
                    <th>Vendor</th>
                    <th>PO #</th>
                  </tr>
                </thead>
                <tbody>
                  {parsedData.map((row, i) => (
                    <tr key={i}>
                      <td>{row["Req Qty"] || ""}</td>
                      <td>{row["SKU Code"] || ""}</td>
                      <td>{toNumber(getBuyPrice(row))}</td>
                      <td>{add5Percent(getBuyPrice(row))}</td>
                      <td>{toNumber(row["Sell Price"])}</td>
                      <td>{calculateMargin(getBuyPrice(row), row["Sell Price"])}</td>
                      <td>{row["End User Name"] || ""}</td>
                      <td>{row["Vendor"] || ""}</td>
                      <td></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Table 2 */}
          <div className="col-md-6">
            <div className="d-flex justify-content-between align-items-center mb-2">
              <h4>📦 PO Format Table</h4>
              <button className="btn btn-sm btn-outline-secondary" onClick={copyTable2}>
                Copy Table
              </button>
            </div>
            <div className="table-responsive">
              <table className="table table-bordered table-sm table-striped">
                <thead className="table-light">
                  <tr>
                    <th>PO No</th>
                    <th>PO dt</th>
                    <th>Customer</th>
                    <th>Part number</th>
                    <th>Variant/Brand</th>
                    <th>Product type</th>
                    <th>Qty</th>
                    <th>U/P</th>
                    <th>Total</th>
                    <th>ETD</th>
                    <th>BU</th>
                    <th>Supplier</th>
                  </tr>
                </thead>
                <tbody>
                  {parsedData.map((row, i) => {
                    const qty = parseInt(row["Req Qty"] || "");
                    const unitPrice = cleanNumber(getBuyPrice(row));
                    const total = isNaN(qty) || isNaN(unitPrice)
                      ? ""
                      : (qty * unitPrice).toFixed(2);

                    return (
                      <tr key={i}>
                        <td></td>
                        <td></td>
                        <td>{row["End User Name"] || ""}</td>
                        <td>{row["SKU Code"] || ""}</td>
                        <td></td>
                        <td></td>
                        <td>{isNaN(qty) ? "" : qty}</td>
                        <td>{isNaN(unitPrice) ? "" : unitPrice.toFixed(2)}</td>
                        <td>{total}</td>
                        <td></td>
                        <td></td>
                        <td>{row["Vendor"] || ""}</td>
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
