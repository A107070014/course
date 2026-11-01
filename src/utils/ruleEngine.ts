import { MappingRule, ConversionStats } from '../types/excel';

/**
 * Execute transformations on source rows based on mapping rules to produce Excel B rows.
 */
export function executeTransformation(
  sourceRows: Record<string, any>[],
  rules: MappingRule[]
): {
  targetHeaders: string[];
  targetRows: Record<string, any>[];
  stats: ConversionStats;
} {
  const startTime = performance.now();
  const enabledRules = rules.filter((r) => r.enabled);
  const targetHeaders = enabledRules.map((r) => r.targetColumn);

  const targetRows: Record<string, any>[] = [];
  let emptyCellsCount = 0;
  let totalCellsCount = 0;

  for (let rowIndex = 0; rowIndex < sourceRows.length; rowIndex++) {
    const srcRow = sourceRows[rowIndex];
    const tgtRow: Record<string, any> = {};

    for (const rule of enabledRules) {
      let cellValue = evaluateRule(rule, srcRow, rowIndex);

      // Clean & format
      if (rule.trimWhitespace && typeof cellValue === 'string') {
        cellValue = cellValue.trim();
      }
      if (rule.textCase === 'upper' && typeof cellValue === 'string') {
        cellValue = cellValue.toUpperCase();
      } else if (rule.textCase === 'lower' && typeof cellValue === 'string') {
        cellValue = cellValue.toLowerCase();
      }

      tgtRow[rule.targetColumn] = cellValue !== undefined && cellValue !== null ? cellValue : '';

      totalCellsCount++;
      if (tgtRow[rule.targetColumn] === '') {
        emptyCellsCount++;
      }
    }

    targetRows.push(tgtRow);
  }

  const endTime = performance.now();

  const stats: ConversionStats = {
    sourceRowCount: sourceRows.length,
    targetRowCount: targetRows.length,
    mappedColumnsCount: enabledRules.filter(
      (r) => r.ruleType !== 'constant' && r.sourceColumns && r.sourceColumns.length > 0
    ).length,
    totalTargetColumnsCount: targetHeaders.length,
    emptyCellRate: totalCellsCount > 0 ? (emptyCellsCount / totalCellsCount) * 100 : 0,
    executionTimeMs: Math.round(endTime - startTime),
  };

  return { targetHeaders, targetRows, stats };
}

/**
 * Evaluate single cell transformation rule against a source row.
 */
function evaluateRule(rule: MappingRule, row: Record<string, any>, rowIndex: number): any {
  switch (rule.ruleType) {
    case 'direct': {
      const col = rule.sourceColumns?.[0];
      if (!col || row[col] === undefined) {
        return rule.defaultValue !== undefined ? rule.defaultValue : '';
      }
      const val = row[col];
      return val !== '' ? val : rule.defaultValue || '';
    }

    case 'combine': {
      const separator = rule.separator !== undefined ? rule.separator : ' ';
      const parts: string[] = [];
      for (const col of rule.sourceColumns || []) {
        const val = row[col];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          parts.push(String(val).trim());
        }
      }
      if (parts.length === 0) {
        return rule.defaultValue || '';
      }
      return parts.join(separator);
    }

    case 'constant': {
      return rule.defaultValue !== undefined ? rule.defaultValue : '';
    }

    case 'split': {
      const col = rule.sourceColumns?.[0];
      const val = col ? String(row[col] || '') : '';
      const sep = rule.separator || ' ';
      const idx = rule.splitIndex || 0;
      const parts = val.split(sep);
      return parts[idx] !== undefined ? parts[idx].trim() : rule.defaultValue || '';
    }

    case 'math': {
      return evaluateMathFormula(rule, row);
    }

    case 'date_format': {
      const col = rule.sourceColumns?.[0];
      const rawDate = col ? row[col] : '';
      if (!rawDate) return rule.defaultValue || '';
      return formatDate(String(rawDate), rule.datePattern || 'YYYY-MM-DD');
    }

    case 'lookup': {
      const col = rule.sourceColumns?.[0];
      const val = col ? String(row[col] || '').trim() : '';
      if (rule.lookupMap && rule.lookupMap[val] !== undefined) {
        return rule.lookupMap[val];
      }
      return rule.defaultValue !== undefined && rule.defaultValue !== ''
        ? rule.defaultValue
        : val;
    }

    case 'case_trim': {
      const col = rule.sourceColumns?.[0];
      const val = col ? String(row[col] || '') : '';
      return val.trim();
    }

    default: {
      const col = rule.sourceColumns?.[0];
      return col && row[col] !== undefined ? row[col] : rule.defaultValue || '';
    }
  }
}

/**
 * Safely compute math expressions between source columns
 */
function evaluateMathFormula(rule: MappingRule, row: Record<string, any>): any {
  if (rule.formula) {
    let expr = rule.formula;
    // Replace column names with actual numbers
    for (const col of rule.sourceColumns || []) {
      const num = parseFloat(String(row[col]).replace(/[^0-9.-]/g, '')) || 0;
      // replace occurrences of column name
      expr = expr.split(`[${col}]`).join(String(num));
      expr = expr.split(col).join(String(num));
    }

    // Also support placeholder A, B, C...
    (rule.sourceColumns || []).forEach((col, idx) => {
      const letter = String.fromCharCode(65 + idx); // 'A', 'B', 'C'
      const num = parseFloat(String(row[col]).replace(/[^0-9.-]/g, '')) || 0;
      const regex = new RegExp(`\\b${letter}\\b`, 'g');
      expr = expr.replace(regex, String(num));
    });

    try {
      // Evaluate safe mathematical expression
      if (/^[0-9+\-*/().\s]+$/.test(expr)) {
        // eslint-disable-next-line no-new-func
        const result = Function(`"use strict"; return (${expr});`)();
        return isNaN(result) ? (rule.defaultValue || '0') : String(Math.round(result * 100) / 100);
      }
    } catch {
      // fallback
    }
  }

  // Default math: multiply all source columns
  if (rule.sourceColumns && rule.sourceColumns.length > 0) {
    let product = 1;
    let found = false;
    for (const col of rule.sourceColumns) {
      const raw = row[col];
      if (raw !== undefined && raw !== null && raw !== '') {
        const num = parseFloat(String(raw).replace(/[^0-9.-]/g, ''));
        if (!isNaN(num)) {
          product *= num;
          found = true;
        }
      }
    }
    return found ? String(Math.round(product * 100) / 100) : rule.defaultValue || '0';
  }

  return rule.defaultValue || '0';
}

/**
 * Flexible date parsing and formatting
 */
export function formatDate(input: string, pattern = 'YYYY-MM-DD'): string {
  if (!input) return '';
  const str = input.trim();

  let year: number | null = null;
  let month: number | null = null;
  let day: number | null = null;

  // Check ROC (Taiwan) year, e.g. 113/05/20 or 113.05.20 or 113年5月20日
  const rocMatch = str.match(/^(\d{2,3})[./年\-](\d{1,2})[./月\-](\d{1,2})/);
  if (rocMatch && parseInt(rocMatch[1], 10) < 1900) {
    year = parseInt(rocMatch[1], 10) + 1911;
    month = parseInt(rocMatch[2], 10);
    day = parseInt(rocMatch[3], 10);
  } else {
    // Normal date parsing
    const dateObj = new Date(str);
    if (!isNaN(dateObj.getTime())) {
      year = dateObj.getFullYear();
      month = dateObj.getMonth() + 1;
      day = dateObj.getDate();
    } else {
      // Regex fallback for YYYY/MM/DD or YYYY-MM-DD
      const m = str.match(/(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})/);
      if (m) {
        year = parseInt(m[1], 10);
        month = parseInt(m[2], 10);
        day = parseInt(m[3], 10);
      }
    }
  }

  if (!year || !month || !day) {
    return str; // return original if unparseable
  }

  const yyyy = String(year);
  const mm = String(month).padStart(2, '0');
  const dd = String(day).padStart(2, '0');

  return pattern
    .replace('YYYY', yyyy)
    .replace('MM', mm)
    .replace('DD', dd)
    .replace('M', String(month))
    .replace('D', String(day));
}
