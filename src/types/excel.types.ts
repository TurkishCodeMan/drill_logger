export type CellType = 'text' | 'number' | 'date' | 'select';

export interface ColumnGroup {
  name: string;
  columns: ExcelColumn[];
}

export interface ExcelColumn {
  name: string;
  type: CellType;
  group?: string;
  required?: boolean;
  options?: string[];
  columnIndex?: number;
  validation?: {
    min?: number;
    max?: number;
    pattern?: string;
    message?: string;
  };
}

export interface ExcelRow {
  [key: string]: string | number | boolean | Date;
}

export interface ExcelSheet {
  name: string;
  columns: ExcelColumn[];
  columnGroups: ColumnGroup[];
  rows: ExcelRow[];
  currentRowIndex?: number;
  headerRow: number;
  subHeaderRow: number;
  design: {
    headerColor: string;
    fontStyle: string;
    rowColors?: {
      even: string;
      odd: string;
    };
    borders?: {
      color: string;
      style: string;
    };
  };
}

export interface ExcelFile {
  sheets: {
    [key: string]: ExcelSheet;
  };
  metadata?: {
    author?: string;
    lastModified?: Date;
    version?: string;
  };
}

export interface ExcelGroup {
  name: string;
  color: string;
  columns: string[];
  startColumn: number;
  endColumn: number;
}

export interface ExcelConfig {
  sheetName: string;
  groupRow: number;
  columnRow: number;
  dataStartRow: number;
}

export interface ExcelData {
  groups: ExcelGroup[];
  rows: Record<string, any>[];
  dropdownFields?: string[];
}

export interface ExcelUploadResponse {
  success: boolean;
  message: string;
  data?: ExcelData;
  error?: string;
} 