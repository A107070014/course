export type RuleType =
  | 'direct'
  | 'combine'
  | 'constant'
  | 'split'
  | 'math'
  | 'date_format'
  | 'lookup'
  | 'case_trim'
  | 'ai_transform';

export interface MappingRule {
  id: string;
  targetColumn: string;
  ruleType: RuleType;
  sourceColumns: string[];
  defaultValue?: string;
  separator?: string;
  splitIndex?: number;
  formula?: string;
  datePattern?: string;
  lookupMap?: Record<string, string>;
  trimWhitespace?: boolean;
  textCase?: 'none' | 'upper' | 'lower';
  explanation?: string;
  enabled: boolean;
}

export interface ExcelSheetData {
  fileName: string;
  fileSize: number;
  sheetNames: string[];
  activeSheet: string;
  headers: string[];
  rows: Record<string, any>[];
  totalRows: number;
  headerRowIndex?: number;
  rawBuffer?: ArrayBuffer;
  detectedStyles?: {
    headerRowNumber: number;
    headerBgColor?: string;
    headerFontColor?: string;
    fontName?: string;
    sampleRowCount: number;
    columnWidthsCount: number;
    hasFormulas?: boolean;
    hasFooter?: boolean;
    hasZebraStriping?: boolean;
  };
}

export interface ConversionTemplate {
  id: string;
  name: string;
  description: string;
  createdAt: string;
  targetHeaders: string[];
  rules: MappingRule[];
  postFilters?: {
    skipEmptyRows?: boolean;
    distinctBy?: string;
  };
}

export interface ConversionStats {
  sourceRowCount: number;
  targetRowCount: number;
  mappedColumnsCount: number;
  totalTargetColumnsCount: number;
  emptyCellRate: number;
  executionTimeMs: number;
}
