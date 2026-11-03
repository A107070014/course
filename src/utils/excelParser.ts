import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { ExcelSheetData } from '../types/excel';

function deepClone<T>(obj: T): T | undefined {
  if (obj === undefined || obj === null) return undefined;
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch {
    return undefined;
  }
}

function normalizeHeader(str: any): string {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[\s_()（）[\]]/g, '')
    .trim()
    .toLowerCase();
}

// 統一黑色框線標準定義 (Solid Black Borders)
const BLACK_CELL_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
};

const BLACK_HEADER_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'medium', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
};

/**
 * 100% 純前端 Excel 解析函式（無需任何後端）
 * 規範：第 1 列嚴格為表格 Header，由 ExcelJS 負責樣式與二進位緩衝區操作。
 */
export async function parseExcelFile(
  file: File,
  sheetIndexOrName?: number | string,
  forcedHeaderRowIndex?: number
): Promise<ExcelSheetData> {
  let buffer = await file.arrayBuffer();

  // Read via SheetJS for fast multi-format table extraction
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellDates: true,
    cellText: false,
    raw: false,
  });

  const sheetNames = workbook.SheetNames;
  if (!sheetNames || sheetNames.length === 0) {
    throw new Error('此 Excel 檔案中沒有發現任何工作表 (Sheet)。');
  }

  let activeSheetName = sheetNames[0];
  if (typeof sheetIndexOrName === 'string' && sheetNames.includes(sheetIndexOrName)) {
    activeSheetName = sheetIndexOrName;
  } else if (typeof sheetIndexOrName === 'number' && sheetNames[sheetIndexOrName]) {
    activeSheetName = sheetNames[sheetIndexOrName];
  }

  // If the uploaded file is .xls or .csv, convert to a real .xlsx ArrayBuffer
  const isXlsx = file.name.toLowerCase().endsWith('.xlsx');
  if (!isXlsx) {
    try {
      const xlsxBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
      buffer = xlsxBuffer as ArrayBuffer;
    } catch (e) {
      console.warn('Failed to convert non-xlsx file to xlsx buffer:', e);
    }
  }

  const worksheet = workbook.Sheets[activeSheetName];
  const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    blankrows: false,
  });

  if (!jsonData || jsonData.length === 0) {
    return {
      fileName: file.name,
      fileSize: file.size,
      sheetNames,
      activeSheet: activeSheetName,
      headers: [],
      rows: [],
      totalRows: 0,
      headerRowIndex: 0,
      rawBuffer: buffer,
    };
  }

  // 第 1 列嚴格為 Header（索引 0）
  const headerRowIndex = forcedHeaderRowIndex !== undefined ? forcedHeaderRowIndex : 0;

  const rawHeaders = (jsonData[headerRowIndex] || []).map((h, idx) =>
    h !== undefined && h !== null && String(h).trim() !== ''
      ? String(h).trim()
      : `未命名欄位_${idx + 1}`
  );

  const rows: Record<string, any>[] = [];
  for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
    const rowValues = jsonData[i] || [];
    const hasValues = rowValues.some(
      (v) => v !== undefined && v !== null && String(v).trim() !== ''
    );
    if (!hasValues) continue;

    const rowObj: Record<string, any> = {};
    rawHeaders.forEach((colName, colIdx) => {
      const val = rowValues[colIdx];
      rowObj[colName] = formatCellValue(val);
    });
    rows.push(rowObj);
  }

  // 利用 ExcelJS 深層檢視樣式
  let detectedStyles: ExcelSheetData['detectedStyles'] = undefined;
  try {
    const inspectWb = new ExcelJS.Workbook();
    await inspectWb.xlsx.load(buffer.slice(0));
    const ws = inspectWb.getWorksheet(activeSheetName) || inspectWb.worksheets[0];

    if (ws) {
      const hRow = ws.getRow(headerRowIndex + 1);
      const c1 = hRow.getCell(1);

      let headerBg = '';
      if (c1.fill && (c1.fill as any).fgColor) {
        const fg = (c1.fill as any).fgColor;
        headerBg = fg.argb ? `#${fg.argb.slice(-6)}` : '';
      }

      let headerFontColor = '';
      if (c1.font && c1.font.color) {
        const col = c1.font.color as any;
        headerFontColor = col.argb ? `#${col.argb.slice(-6)}` : '';
      }

      const sample1 = ws.getRow(headerRowIndex + 2);
      const sample2 = ws.getRow(headerRowIndex + 3);
      const hasZebra =
        sample1.actualCellCount > 0 &&
        sample2.actualCellCount > 0 &&
        JSON.stringify(sample1.getCell(1).fill) !== JSON.stringify(sample2.getCell(1).fill);

      detectedStyles = {
        headerRowNumber: headerRowIndex + 1,
        headerBgColor: headerBg || '#1F4E79',
        headerFontColor: headerFontColor || '#FFFFFF',
        fontName: c1.font?.name || '微軟正黑體',
        sampleRowCount: Math.max(0, ws.rowCount - (headerRowIndex + 1)),
        columnWidthsCount: ws.columns?.filter((c) => c.width && c.width > 0).length || ws.columnCount,
        hasZebraStriping: hasZebra,
      };
    }
  } catch (err) {
    console.warn('ExcelJS inspection non-fatal notice:', err);
  }

  return {
    fileName: file.name,
    fileSize: file.size,
    sheetNames,
    activeSheet: activeSheetName,
    headers: rawHeaders,
    rows,
    totalRows: rows.length,
    headerRowIndex,
    rawBuffer: buffer,
    detectedStyles,
  };
}

function formatCellValue(val: any): any {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    const h = val.getHours();
    const min = val.getMinutes();
    if (h === 0 && min === 0) {
      return `${y}-${m}-${d}`;
    }
    return `${y}-${m}-${d} ${String(h).padStart(2, '0')}:${String(min).padStart(2, '0')}`;
  }
  return String(val).trim();
}

/**
 * 100% 純前端 ExcelJS 匯出引擎：
 * 1. 外框線與內部框線一律為黑色 (Thin/Medium Black Borders FF000000)
 * 2. 第 1 列 Header 的字體一律為粗體字 (Bold: true)
 * 3. 第 2 列起依序填入轉換資料，完整繼承原範本樣式與數字幣別/日期格式
 */
export async function exportToXLSXWithTemplate(
  headers: string[],
  rows: Record<string, any>[],
  templateData?: ExcelSheetData | null,
  filename = 'excel_b_converted.xlsx'
): Promise<void> {
  const wb = new ExcelJS.Workbook();

  if (templateData && templateData.rawBuffer) {
    try {
      // 1. 載入 Excel B 範本
      await wb.xlsx.load(templateData.rawBuffer.slice(0));
      const ws = wb.getWorksheet(templateData.activeSheet) || wb.worksheets[0];

      if (ws) {
        // 第 1 列為表頭
        const headerRowNumber = 1;
        const headerRow = ws.getRow(headerRowNumber);
        const colCount = Math.max(ws.columnCount, headers.length);

        // 強制設定第一列 Header：粗體字 (bold: true) + 黑色框線 (black borders)
        headerRow.height = Math.max(headerRow.height || 28, 28);
        for (let c = 1; c <= colCount; c++) {
          const hCell = headerRow.getCell(c);
          hCell.font = {
            ...(hCell.font || { name: 'Microsoft JhengHei', size: 11 }),
            bold: true, // 使用者指定：第 1 列 Header 字體必須為粗體字
          };
          hCell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'medium', color: { argb: 'FF000000' } }, // 表頭底部黑線
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
        }
        headerRow.commit();

        // 收集欄位資訊
        const templateCols: { colIndex: number; text: string; normalized: string }[] = [];
        for (let c = 1; c <= colCount; c++) {
          const cell = headerRow.getCell(c);
          const text = (cell.text || String(cell.value || '')).trim();
          templateCols.push({
            colIndex: c,
            text,
            normalized: normalizeHeader(text),
          });
        }

        // 擷取範例列樣式（第 2 列，以及若有第 3 列的斑馬紋）
        interface CellStyleSnapshot {
          font?: any;
          fill?: any;
          border?: any;
          alignment?: any;
          numFmt?: string;
          protection?: any;
        }
        interface RowStyleSnapshot {
          height?: number;
          cells: Record<number, CellStyleSnapshot>;
        }

        const sampleStyles: RowStyleSnapshot[] = [];
        const sampleRowNumbers = [2];
        if (ws.rowCount >= 3 && ws.getRow(3).actualCellCount > 0) {
          sampleRowNumbers.push(3);
        }

        sampleRowNumbers.forEach((rNum) => {
          const row = ws.getRow(rNum);
          const cells: Record<number, CellStyleSnapshot> = {};

          templateCols.forEach(({ colIndex }) => {
            const cell = row.getCell(colIndex);
            cells[colIndex] = {
              font: deepClone(cell.font),
              fill: deepClone(cell.fill),
              // 外框線強制為黑色
              border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } },
              },
              alignment: deepClone(cell.alignment),
              numFmt: cell.numFmt,
              protection: deepClone(cell.protection),
            };
          });

          sampleStyles.push({
            height: row.height || 22,
            cells,
          });
        });

        // 若原範本完全沒有第 2 列，則基於第 1 列建立預設資料列樣式
        if (sampleStyles.length === 0 || !sampleStyles[0].cells[1]?.font) {
          const cells: Record<number, CellStyleSnapshot> = {};
          templateCols.forEach(({ colIndex }) => {
            cells[colIndex] = {
              font: { name: headerRow.getCell(colIndex).font?.name || '微軟正黑體', size: 10 },
              border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } },
              },
            };
          });
          sampleStyles.push({ height: 22, cells });
        }

        // 第 2 列起依序填入轉換資料
        rows.forEach((rowObj, rIdx) => {
          const rowNumber = 2 + rIdx;
          const targetRow = ws.getRow(rowNumber);
          const stylePattern = sampleStyles[rIdx % sampleStyles.length];

          if (stylePattern.height) {
            targetRow.height = stylePattern.height;
          }

          templateCols.forEach(({ colIndex, text, normalized }) => {
            const cell = targetRow.getCell(colIndex);

            let val: any = undefined;
            if (rowObj[text] !== undefined) {
              val = rowObj[text];
            } else {
              const matchKey = Object.keys(rowObj).find(
                (k) => normalizeHeader(k) === normalized
              );
              if (matchKey && rowObj[matchKey] !== undefined) {
                val = rowObj[matchKey];
              } else if (headers[colIndex - 1] && rowObj[headers[colIndex - 1]] !== undefined) {
                val = rowObj[headers[colIndex - 1]];
              }
            }

            const style = stylePattern.cells[colIndex];
            const strVal = val !== undefined && val !== null ? String(val).trim() : '';

            const isNumeric =
              strVal !== '' &&
              /^-?\d+(\.\d+)?$/.test(strVal) &&
              !(strVal.startsWith('0') && strVal.length > 1 && !strVal.startsWith('0.'));

            if (isNumeric) {
              cell.value = Number(strVal);
            } else {
              cell.value = strVal;
            }

            // 套用字體、底色、黑色外框線與數值格式
            if (style) {
              if (style.font) cell.font = deepClone(style.font);
              if (style.fill) cell.fill = deepClone(style.fill);
              // 強制套用黑色外框線
              cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } },
              };
              if (style.alignment) {
                cell.alignment = deepClone(style.alignment);
              } else {
                cell.alignment = isNumeric
                  ? { vertical: 'middle', horizontal: 'right' }
                  : { vertical: 'middle', horizontal: 'left' };
              }
              if (style.numFmt) {
                cell.numFmt = style.numFmt;
              } else if (isNumeric) {
                cell.numFmt = '#,##0';
              }
            } else {
              cell.font = { name: 'Microsoft JhengHei', size: 10 };
              cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } },
              };
            }
          });

          targetRow.commit();
        });

        // 清理原範本多餘的未填列
        const expectedRowCount = 1 + rows.length;
        if (ws.rowCount > expectedRowCount) {
          ws.spliceRows(expectedRowCount + 1, ws.rowCount - expectedRowCount);
        }

        // 下載 Excel 檔案
        const outBuffer = await wb.xlsx.writeBuffer();
        triggerDownload(
          outBuffer,
          filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`,
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        return;
      }
    } catch (err) {
      console.warn('ExcelJS template cloning failed, using styled fallback:', err);
    }
  }

  // Fallback: 純前端 ExcelJS 生成標準企業樣式
  const ws = wb.addWorksheet('轉換結果');
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  // 第 1 列 Header：粗體字 (bold: true) + 黑色外框線 (black border)
  const headerRow = ws.addRow(headers);
  headerRow.height = 28;
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E3A5F' },
    };
    cell.font = {
      name: 'Microsoft JhengHei',
      size: 11,
      bold: true, // 粗體字
      color: { argb: 'FFFFFFFF' },
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'medium', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };
  });

  // 資料列：黑色外框線
  rows.forEach((row, rIdx) => {
    const rowValues = headers.map((h) => {
      const val = row[h];
      const strVal = val !== undefined && val !== null ? String(val).trim() : '';
      if (
        strVal !== '' &&
        /^-?\d+(\.\d+)?$/.test(strVal) &&
        !(strVal.startsWith('0') && strVal.length > 1 && !strVal.startsWith('0.'))
      ) {
        return Number(strVal);
      }
      return strVal;
    });

    const dataRow = ws.addRow(rowValues);
    dataRow.height = 22;

    const isEven = rIdx % 2 === 0;
    dataRow.eachCell((cell) => {
      cell.font = {
        name: 'Microsoft JhengHei',
        size: 10,
        color: { argb: 'FF1E293B' },
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: isEven ? 'FFFFFFFF' : 'FFF8FAFC' },
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };

      if (typeof cell.value === 'number') {
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      }
    });
  });

  // 自動欄寬
  ws.columns.forEach((column, colIdx) => {
    const headerName = headers[colIdx] || '';
    let maxLen = headerName.length * 2.2;
    for (let r = 0; r < Math.min(rows.length, 50); r++) {
      const cellVal = String(rows[r][headerName] || '');
      maxLen = Math.max(maxLen, cellVal.length * 1.5);
    }
    column.width = Math.min(Math.max(Math.ceil(maxLen), 14), 50);
  });

  const outBuffer = await wb.xlsx.writeBuffer();
  triggerDownload(
    outBuffer,
    filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
}

export async function exportToXLSX(
  headers: string[],
  rows: Record<string, any>[],
  filename = 'excel_b_converted.xlsx'
) {
  return exportToXLSXWithTemplate(headers, rows, null, filename);
}

function triggerDownload(buffer: ExcelJS.Buffer, filename: string, mimeType: string) {
  const blob = new Blob([buffer], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function exportToCSV(
  headers: string[],
  rows: Record<string, any>[],
  filename = 'excel_b_converted.csv'
) {
  const csvContent = [
    headers.map((h) => `"${String(h).replace(/"/g, '""')}"`).join(','),
    ...rows.map((row) =>
      headers
        .map((h) => `"${String(row[h] !== undefined ? row[h] : '').replace(/"/g, '""')}"`)
        .join(',')
    ),
  ].join('\r\n');

  const bom = '\uFEFF';
  const blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename.endsWith('.csv') ? filename : `${filename}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * 產生高品質範本 Excel B（純 ExcelJS 生成）
 * 1. 第 1 列為表頭：粗體字 (bold: true) + 黑色外框線 (Black border)
 * 2. 第 2、3 列為範例資料列：黑色外框線 + 幣別 $#,##0 與日期 yyyy-mm-dd
 */
export async function generateStyledSampleBBuffer(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('DeliveryBatch');

  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const headers = [
    '物流單號',
    '收件客戶',
    '聯絡電話',
    '完整收件地址',
    '商品明細與件數',
    '代收貨款金額',
    '配送溫層',
    '出貨日期',
    '處理狀態',
  ];

  const widths = [16, 14, 16, 42, 30, 18, 14, 16, 14];

  // 第 1 列：表頭（粗體字 + 黑色框線）
  const headerRow = ws.getRow(1);
  headerRow.height = 30;
  headerRow.values = headers;
  headerRow.eachCell((cell, colNum) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E79' },
    };
    cell.font = {
      name: 'Microsoft JhengHei',
      size: 11,
      bold: true, // 粗體字
      color: { argb: 'FFFFFFFF' },
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
    // 黑色外框線
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'medium', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };
    ws.getColumn(colNum).width = widths[colNum - 1] || 18;
  });

  // 第 2 列：範例資料 1（黑色框線）
  const sampleRow1 = ws.getRow(2);
  sampleRow1.height = 24;
  sampleRow1.values = [
    'TC-8801',
    '林冠宇',
    '0912-345-678',
    '台北市大安區敦化南路二段 216 號 12 樓',
    '極致降噪無線耳機 Pro x2',
    3400,
    '常溫',
    '2026-04-10',
    '待分揀',
  ];
  sampleRow1.eachCell((cell, colNum) => {
    cell.font = { name: 'Microsoft JhengHei', size: 10, color: { argb: 'FF000000' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    // 黑色外框線
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };
    if (colNum === 6) {
      cell.numFmt = '$#,##0';
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (colNum === 7 || colNum === 9) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    } else if (colNum === 8) {
      cell.numFmt = 'yyyy-mm-dd';
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
    }
  });

  // 第 3 列：範例資料 2（斑馬紋 + 黑色框線）
  const sampleRow2 = ws.getRow(3);
  sampleRow2.height = 24;
  sampleRow2.values = [
    'TC-8802',
    '陳雅婷',
    '0928-876-543',
    '新北市板橋區文化路一段 188 巷 8 號',
    '人體工學護脊辦公椅 x1',
    6800,
    '常溫',
    '2026-04-10',
    '待分揀',
  ];
  sampleRow2.eachCell((cell, colNum) => {
    cell.font = { name: 'Microsoft JhengHei', size: 10, color: { argb: 'FF000000' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    // 黑色外框線
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };
    if (colNum === 6) {
      cell.numFmt = '$#,##0';
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (colNum === 7 || colNum === 9) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    } else if (colNum === 8) {
      cell.numFmt = 'yyyy-mm-dd';
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}

export function getDemoSampleDatasets(): {
  sampleA: ExcelSheetData;
  sampleB: ExcelSheetData;
} {
  const sampleARows = [
    {
      '訂單編號': 'ORD-2026-8801',
      '訂購人姓名': '林冠宇',
      '收件人手機': '0912-345-678',
      '寄送縣市': '台北市',
      '行政區域': '大安區',
      '街道門牌': '敦化南路二段 216 號 12 樓',
      '商品品名': '極致降噪無線耳機 Pro',
      '數量': '2',
      '單價': '3400',
      '付款方式': '信用卡線上付款',
      '下單時間': '2026/04/10 14:32',
      '顧客備註': '管理室代收，請下午送達',
    },
    {
      '訂單編號': 'ORD-2026-8802',
      '訂購人姓名': '陳雅婷',
      '收件人手機': '0928-876-543',
      '寄送縣市': '新北市',
      '行政區域': '板橋區',
      '街道門牌': '文化路一段 188 巷 8 號',
      '商品品名': '人體工學護脊辦公椅',
      '數量': '1',
      '單價': '6800',
      '付款方式': '貨到付款',
      '下單時間': '2026/04/10 15:45',
      '顧客備註': '送達前請先電聯',
    },
    {
      '訂單編號': 'ORD-2026-8803',
      '訂購人姓名': '張立國',
      '收件人手機': '0933-221-998',
      '寄送縣市': '台中市',
      '行政區域': '西屯區',
      '街道門牌': '台灣大道三段 301 號',
      '商品品名': '4K 高畫質曲面電競螢幕 32吋',
      '數量': '1',
      '單價': '12500',
      '付款方式': 'LINE Pay',
      '下單時間': '2026/04/11 09:12',
      '顧客備註': '易碎品小心輕放',
    },
    {
      '訂單編號': 'ORD-2026-8804',
      '訂購人姓名': '黃靜雯',
      '收件人手機': '0955-432-109',
      '寄送縣市': '高雄市',
      '行政區域': '左營區',
      '街道門牌': '博愛二路 777 號 5 樓',
      '商品品名': '磁吸無線快充行動電源',
      '數量': '3',
      '單價': '990',
      '付款方式': '信用卡線上付款',
      '下單時間': '2026/04/11 11:20',
      '顧客備註': '無',
    },
  ];

  const sampleBHeaders = [
    '物流單號',
    '收件客戶',
    '聯絡電話',
    '完整收件地址',
    '商品明細與件數',
    '代收貨款金額',
    '配送溫層',
    '出貨日期',
    '處理狀態',
  ];

  const sampleBRows = [
    {
      '物流單號': 'TC-8801',
      '收件客戶': '林冠宇',
      '聯絡電話': '0912-345-678',
      '完整收件地址': '台北市大安區敦化南路二段 216 號 12 樓',
      '商品明細與件數': '極致降噪無線耳機 Pro x2',
      '代收貨款金額': '3400',
      '配送溫層': '常溫',
      '出貨日期': '2026-04-10',
      '處理狀態': '待分揀',
    },
    {
      '物流單號': 'TC-8802',
      '收件客戶': '陳雅婷',
      '聯絡電話': '0928-876-543',
      '完整收件地址': '新北市板橋區文化路一段 188 巷 8 號',
      '商品明細與件數': '人體工學護脊辦公椅 x1',
      '代收貨款金額': '6800',
      '配送溫層': '常溫',
      '出貨日期': '2026-04-10',
      '處理狀態': '待分揀',
    },
  ];

  return {
    sampleA: {
      fileName: '來源_蝦皮線上訂單_A.xlsx',
      fileSize: 14200,
      sheetNames: ['Orders'],
      activeSheet: 'Orders',
      headers: Object.keys(sampleARows[0]),
      rows: sampleARows,
      totalRows: sampleARows.length,
      headerRowIndex: 0,
    },
    sampleB: {
      fileName: '目標_黑貓物流託運標準檔_B.xlsx',
      fileSize: 12800,
      sheetNames: ['DeliveryBatch'],
      activeSheet: 'DeliveryBatch',
      headers: sampleBHeaders,
      rows: sampleBRows,
      totalRows: sampleBRows.length,
      headerRowIndex: 0,
      detectedStyles: {
        headerRowNumber: 1,
        headerBgColor: '#1F4E79',
        headerFontColor: '#FFFFFF',
        fontName: '微軟正黑體',
        sampleRowCount: 2,
        columnWidthsCount: 9,
        hasZebraStriping: true,
      },
    },
  };
}
