import { useState, useRef, useMemo } from "react";
import * as ExcelJS from "exceljs";
import ParseWithStyles, { CellStyleData } from "../utils/ParseWithStyles";

export interface CellHyperlinkValue {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

export const ExcelReader: React.FC = () => {
  const [fileData, setFileData] = useState<ExcelJS.Workbook | null>(null);
  const [cellStyles, setCellStyles] = useState<CellStyleData[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [fileName, setFileName] = useState<string>("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const paginatedCellStyles = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return cellStyles.slice(startIndex, startIndex + itemsPerPage);
  }, [cellStyles, currentPage]);

  const totalPages = Math.ceil(cellStyles.length / itemsPerPage);

  const handlePageChange = (newPage: number) => {
    setCurrentPage(newPage);
  };

  const PaginationControls = () => (
    <div
      style={{
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        marginTop: "10px",
        gap: "10px",
      }}
    >
      <button
        onClick={() => handlePageChange(currentPage - 1)}
        disabled={currentPage === 1}
        style={{
          padding: "5px 10px",
          cursor: currentPage === 1 ? "not-allowed" : "pointer",
        }}
      >
        Previous
      </button>

      <span>
        Page {currentPage} of {totalPages}
      </span>

      <button
        onClick={() => handlePageChange(currentPage + 1)}
        disabled={currentPage === totalPages}
        style={{
          padding: "5px 10px",
          cursor: currentPage === totalPages ? "not-allowed" : "pointer",
        }}
      >
        Next
      </button>
    </div>
  );

  const renderCellValue = (
    value: ExcelJS.CellValue,
    cellStyle: CellStyleData["style"]
  ): React.ReactNode => {
    if (value === null || value === undefined) return "";

    if (typeof value === "string") {
      const styleProps: React.CSSProperties = {
        fontWeight: cellStyle?.font?.bold ? "bold" : "normal",
        fontStyle: cellStyle?.font?.italic ? "italic" : "normal",
        textDecoration: cellStyle?.font?.underline ? "underline" : "none",
        color: cellStyle?.font?.color
          ? `rgb(${cellStyle.font.color})`
          : "#f9f9f9",
        fontSize: cellStyle?.font?.size
          ? `${cellStyle.font.size}pt`
          : "inherit",
        fontFamily: cellStyle?.font?.name || "inherit",
        backgroundColor: cellStyle?.fill?.color
          ? `rgb(${cellStyle.fill.color})`
          : "transparent",
      };

      return <span style={styleProps}>{value}</span>;
    }

    if (typeof value === "number")
      return (
        <span style={{ color: "#f9f9f9", textAlign: "right" }}>
          {value.toString()}
        </span>
      );

    if (typeof value === "boolean")
      return (
        <span
          style={{
            color: value ? "green" : "red",
            fontWeight: "bold",
          }}
        >
          {value.toString()}
        </span>
      );

    if (value instanceof Date)
      return (
        <span style={{ color: "purple" }}>{value.toLocaleDateString()}</span>
      );

    if ("richText" in value) {
      return value.richText.map((rt, index) => (
        <span
          key={index}
          style={{
            fontWeight: rt.font?.bold ? "bold" : "normal",
            fontStyle: rt.font?.italic ? "italic" : "normal",
            textDecoration: rt.font?.underline ? "underline" : "none",
            color: rt.font?.color ? `rgb(${rt.font.color.argb})` : "#000",
            fontSize: rt.font?.size ? `${rt.font.size}pt` : "inherit",
            fontFamily: rt.font?.name || "inherit",
            verticalAlign:
              rt.font?.vertAlign === "superscript"
                ? "super"
                : rt.font?.vertAlign === "subscript"
                ? "sub"
                : "baseline",
          }}
        >
          {rt.text}
        </span>
      ));
    }

    if ("hyperlink" in value && "text" in value) {
      return (
        <a
          href={value.hyperlink}
          target="_blank"
          rel="noopener noreferrer"
          style={{ color: "blue", textDecoration: "underline" }}
        >
          {value.hyperlink}
        </a>
      );
    }

    // Fallback for complex types with error-like styling
    return (
      <span
        style={{
          color: "red",
          fontStyle: "italic",
          backgroundColor: "#f8f8f8",
        }}
      >
        {JSON.stringify(value)}
      </span>
    );
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setIsLoading(true);
      setFileName(file.name);

      try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());

        const stylesData = await ParseWithStyles.parseExcelStyles(file);

        const filteredStyles = ParseWithStyles.filterStyles(
          stylesData,
          (cellData) => cellData.style?.font?.bold === true
        );

        setCellStyles(stylesData);
        setFileData(workbook);
      } catch (error) {
        console.error("Error parsing Excel file:", error);
        alert(
          "Error parsing Excel file. Please try again with a different file."
        );
      } finally {
        setIsLoading(false);
      }
    }
  };

  const handleButtonClick = () => {
    fileInputRef.current?.click();
  };

  return (
    <div className="excel-reader-content">
      <div
        style={{
          display: "flex",
          justifyContent: "center",
          alignContent: "center",
        }}
      >
        <input
          ref={fileInputRef}
          type="file"
          onChange={handleFileUpload}
          accept=".xlsx, .xls"
          disabled={isLoading}
          id="file-upload"
        />
        <button
          onClick={handleButtonClick}
          disabled={isLoading}
          className="file-upload-label"
        >
          {isLoading ? "Uploading..." : "Choose Excel File"}
        </button>
      </div>

      <div
        className="file-info"
        style={{
          display: "flex",
          justifyContent: "center",
          alignContent: "center",
        }}
      >
        {fileName && !isLoading && <>{fileName}</>}
      </div>

      {isLoading && (
        <div className="loader-container">
          <div className="loader"></div>
          <p className="loader-text">Parsing Excel file, please wait...</p>
        </div>
      )}

      {fileData && !isLoading && (
        <div className="results-container">
          <div className="results-header">
            <h3>Parsed Styles</h3>
            <span className="results-count">
              {cellStyles.length} cells found
            </span>
          </div>

          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th>Cell</th>
                  <th>Value</th>
                  <th>Style</th>
                </tr>
              </thead>
              <tbody>
                {paginatedCellStyles.map((data, index) => (
                  <tr key={index}>
                    <td>{data.cell}</td>
                    <td>{renderCellValue(data.value, data.style)}</td>
                    <td>
                      {data.style ? (
                        <details>
                          <summary>View Style</summary>
                          <pre
                            style={{
                              backgroundColor: "rgb(45, 47, 48)",
                              border: "1px solid #e0e0e0",
                              borderRadius: "4px",
                              padding: "10px",
                              color: "rgb(247, 239, 239)",
                              maxHeight: "200px",
                              overflowY: "auto",
                              fontFamily: "monospace",
                              fontSize: "0.8em",
                              userSelect: "text",
                              cursor: "text",
                            }}
                          >
                            <code>
                              {JSON.stringify(
                                data.style,
                                (_, value) => {
                                  if (
                                    typeof value === "object" &&
                                    value !== null
                                  ) {
                                    if (
                                      Array.isArray(value) &&
                                      value.length > 10
                                    ) {
                                      return `[${value.length} items]`;
                                    }
                                    if (Object.keys(value).length > 10) {
                                      return `{${
                                        Object.keys(value).length
                                      } keys}`;
                                    }
                                  }
                                  return value;
                                },
                                2
                              )}
                            </code>
                          </pre>
                        </details>
                      ) : (
                        <span style={{ color: "#888", fontStyle: "italic" }}>
                          No style
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            {cellStyles.length > itemsPerPage && <PaginationControls />}

            {cellStyles.length > itemsPerPage && (
              <p className="results-count">
                Showing first {itemsPerPage} of {cellStyles.length} cells
              </p>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
