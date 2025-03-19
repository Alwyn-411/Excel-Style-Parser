import * as ExcelJS from "exceljs";

export interface FontStyle {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  subscript?: boolean;
  superscript?: boolean;
}

export interface FillStyle {
  type?: string;
  color?: string;
}

export interface AlignmentStyle {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
}

export interface BorderStyle {
  style?: string;
  color?: string;
}

export interface CellStyleData {
  cell: string;
  value: ExcelJS.CellValue;
  rowIndex: number;
  colIndex: number;
  style?: {
    font?: FontStyle;
    fill?: FillStyle;
    alignment?: AlignmentStyle;
    border?: {
      top?: BorderStyle;
      bottom?: BorderStyle;
      left?: BorderStyle;
      right?: BorderStyle;
    };
  };
}

export class ParseWithStyles {
  static async parseExcelStyles(
    file: File,
    sheetIndex: number = 0
  ): Promise<CellStyleData[]> {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());

      const worksheet = workbook.worksheets[sheetIndex];

      if (!worksheet) throw new Error(`Sheet at index ${sheetIndex} not found`);

      const stylesData: CellStyleData[] = [];

      worksheet.eachRow((row, rowIndex) => {
        row.eachCell((cell, colIndex) => {
          const cellStyleData: CellStyleData = {
            cell: this.getCellReference(rowIndex, colIndex),
            value: cell.value,
            rowIndex,
            colIndex,
            style: this.extractCellStyle(cell),
          };

          stylesData.push(cellStyleData);
        });
      });

      return stylesData;
    } catch (error) {
      console.error("Excel parsing error:", error);
      throw error;
    }
  }

  private static extractCellStyle(cell: ExcelJS.Cell): CellStyleData["style"] {
    const style: CellStyleData["style"] = {};

    if (cell.font) {
      style.font = {
        name: cell.font.name,
        size: cell.font.size,
        color: cell.font.color?.argb,
        bold: cell.font.bold,
        italic: cell.font.italic,
        subscript: cell.font.vertAlign === "subscript",
        superscript: cell.font.vertAlign === "superscript",
        underline:
          typeof cell.font.underline === "boolean"
            ? cell.font.underline
            : undefined,
      };
    }

    if (cell.fill && cell.fill.type === "pattern") {
      style.fill = {
        type: cell.fill.type,
        color: (cell.fill as ExcelJS.FillPattern).fgColor?.argb,
      };
    }

    if (cell.alignment) {
      style.alignment = {
        horizontal: cell.alignment.horizontal,
        vertical: cell.alignment.vertical,
        wrapText: cell.alignment.wrapText,
      };
    }

    if (cell.border) {
      style.border = {
        top: this.extractBorderStyle(cell.border.top as ExcelJS.Border),
        bottom: this.extractBorderStyle(cell.border.bottom as ExcelJS.Border),
        left: this.extractBorderStyle(cell.border.left as ExcelJS.Border),
        right: this.extractBorderStyle(cell.border.right as ExcelJS.Border),
      };
    }

    return style;
  }

  private static extractBorderStyle(
    border?: ExcelJS.Border
  ): BorderStyle | undefined {
    if (!border) return undefined;
    return {
      style: border.style,
      color: border.color?.argb,
    };
  }

  private static getCellReference(rowIndex: number, colIndex: number): string {
    const columnLetter = this.getColumnLetter(colIndex);
    return `${columnLetter}${rowIndex}`;
  }

  private static getColumnLetter(colIndex: number): string {
    let letter = "";
    while (colIndex > 0) {
      colIndex--;
      letter = String.fromCharCode(65 + (colIndex % 26)) + letter;
      colIndex = Math.floor(colIndex / 26);
    }
    return letter;
  }

  static filterStyles(
    stylesData: CellStyleData[],
    filterFn?: (cellData: CellStyleData) => boolean
  ): CellStyleData[] {
    return filterFn ? stylesData.filter(filterFn) : stylesData;
  }
}

export default ParseWithStyles;
