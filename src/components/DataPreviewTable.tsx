import React, { useState } from 'react';
import {
  Search,
  Download,
  FileSpreadsheet,
  CheckCircle2,
  Clock,
  ArrowRight,
  Maximize2,
  Columns,
  Table as TableIcon,
} from 'lucide-react';
import { ExcelSheetData, ConversionStats } from '../types/excel';

interface DataPreviewTableProps {
  sheetA: ExcelSheetData | null;
  sheetB?: ExcelSheetData | null;
  targetHeaders: string[];
  targetRows: Record<string, any>[];
  stats: ConversionStats | null;
  onExportXLSX: () => void;
  onExportCSV: () => void;
}

export const DataPreviewTable: React.FC<DataPreviewTableProps> = ({
  sheetA,
  sheetB,
  targetHeaders,
  targetRows,
  stats,
  onExportXLSX,
  onExportCSV,
}) => {
  const [viewMode, setViewMode] = useState<'split' | 'targetOnly' | 'sourceOnly'>('split');
  const [searchQuery, setSearchQuery] = useState('');
  const [page, setPage] = useState(1);
  const rowsPerPage = 10;

  // Filter rows by search
  const filteredTargetRows = targetRows.filter((row) => {
    if (!searchQuery) return true;
    const query = searchQuery.toLowerCase();
    return Object.values(row).some((val) => String(val).toLowerCase().includes(query));
  });

  const totalPages = Math.ceil(filteredTargetRows.length / rowsPerPage) || 1;
  const paginatedTargetRows = filteredTargetRows.slice(
    (page - 1) * rowsPerPage,
    page * rowsPerPage
  );

  const paginatedSourceRows = (sheetA?.rows || []).slice(
    (page - 1) * rowsPerPage,
    page * rowsPerPage
  );

  return (
    <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-xs space-y-0">
      {/* Top Bar / Stats Strip */}
      <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center gap-4 flex-wrap text-xs text-slate-600">
          <div className="flex items-center gap-1.5 font-medium text-slate-900">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500" />
            <span>轉譯完成</span>
          </div>
          <span>·</span>
          <span>
            共 <strong className="font-mono tabular-nums text-slate-900">{targetRows.length}</strong> 筆資料
          </span>
          <span>·</span>
          <span>
            目標欄位 <strong className="font-mono tabular-nums text-slate-900">{targetHeaders.length}</strong> 個
          </span>
          {stats && (
            <>
              <span>·</span>
              <span className="inline-flex items-center gap-1 text-slate-500">
                <Clock className="w-3 h-3" />
                耗時 <span className="font-mono tabular-nums">{stats.executionTimeMs}ms</span>
              </span>
            </>
          )}
        </div>

        {/* View mode toggle & Actions */}
        <div className="flex items-center gap-2 flex-wrap">
          {/* Search */}
          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              placeholder="搜尋預覽資料..."
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setPage(1);
              }}
              className="text-xs pl-8 pr-3 py-1.5 border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-emerald-600 w-44"
            />
          </div>

          {/* Mode Switcher */}
          <div className="inline-flex rounded-lg border border-slate-200 bg-white p-0.5 text-xs font-medium text-slate-600">
            <button
              onClick={() => setViewMode('split')}
              className={`px-2.5 py-1 rounded-md transition-colors ${
                viewMode === 'split' ? 'bg-slate-100 text-slate-900 font-semibold' : 'hover:text-slate-900'
              }`}
              title="雙表左右並排對照"
            >
              雙表並排
            </button>
            <button
              onClick={() => setViewMode('targetOnly')}
              className={`px-2.5 py-1 rounded-md transition-colors ${
                viewMode === 'targetOnly'
                  ? 'bg-emerald-50 text-emerald-900 font-semibold'
                  : 'hover:text-slate-900'
              }`}
              title="僅顯示目標 Excel B"
            >
              目標 Excel B
            </button>
            <button
              onClick={() => setViewMode('sourceOnly')}
              className={`px-2.5 py-1 rounded-md transition-colors ${
                viewMode === 'sourceOnly' ? 'bg-slate-100 text-slate-900 font-semibold' : 'hover:text-slate-900'
              }`}
              title="僅顯示來源 Excel A"
            >
              來源 Excel A
            </button>
          </div>

          {/* Quick Export */}
          <button
            onClick={onExportXLSX}
            disabled={targetRows.length === 0}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 rounded-lg shadow-xs transition-colors shrink-0"
          >
            <Download className="w-3.5 h-3.5" />
            下載 .xlsx
          </button>
        </div>
      </div>

      {/* Main Table Viewports */}
      <div className="overflow-x-auto">
        {viewMode === 'split' ? (
          <div className="grid grid-cols-1 lg:grid-cols-2 divide-y lg:divide-y-0 lg:divide-x divide-slate-200 min-h-[380px]">
            {/* Left: Source Table Preview */}
            <div className="overflow-x-auto p-4">
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-2">
                  <span className="w-5 h-5 rounded bg-slate-800 text-white text-xs font-bold flex items-center justify-center">
                    A
                  </span>
                  <span className="text-xs font-semibold text-slate-800">
                    來源 Excel A 原始資料
                  </span>
                  <span className="text-[11px] text-slate-400">({sheetA?.fileName || '未上傳'})</span>
                </div>
              </div>

              {!sheetA || sheetA.rows.length === 0 ? (
                <div className="py-16 text-center text-xs text-slate-400">尚未上傳來源資料</div>
              ) : (
                <div className="border border-slate-200 rounded-lg overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-semibold">
                        <th className="py-2 px-3 w-10 text-slate-400 font-mono">#</th>
                        {sheetA.headers.map((h, i) => (
                          <th key={i} className="py-2 px-3 whitespace-nowrap">
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {paginatedSourceRows.map((row, rIdx) => (
                        <tr key={rIdx} className="hover:bg-slate-50 transition-colors">
                          <td className="py-2 px-3 text-slate-400 font-mono text-[11px]">
                            {(page - 1) * rowsPerPage + rIdx + 1}
                          </td>
                          {sheetA.headers.map((h, cIdx) => (
                            <td key={cIdx} className="py-2 px-3 whitespace-nowrap text-slate-700 max-w-xs truncate">
                              {row[h] !== undefined && row[h] !== null ? String(row[h]) : ''}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* Right: Converted Target Table Preview */}
            <div className="overflow-x-auto p-4 bg-emerald-50/10">
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-2">
                  <span className="w-5 h-5 rounded bg-indigo-600 text-white text-xs font-bold flex items-center justify-center">
                    B
                  </span>
                  <span className="text-xs font-semibold text-slate-900">
                    目標 Excel B 轉譯結果 (實時生成)
                  </span>
                </div>
                <span className="text-[11px] text-orange-950 bg-orange-100 border border-orange-300 px-2.5 py-0.5 rounded font-semibold flex items-center gap-1">
                  🎯 已套用規範：第 1 行粗體表頭 ＋ 全黑色上下左右框線 ＋ 白橘交錯底色 (白 ➔ 橘 ➔ 白)
                </span>
              </div>

              {targetRows.length === 0 ? (
                <div className="py-16 text-center text-xs text-slate-400">
                  尚未產出目標資料，請確認轉換規則
                </div>
              ) : (
                <div className="border-2 border-black rounded-lg overflow-x-auto bg-white shadow-xs">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-orange-500 text-white font-bold">
                        <th className="py-2.5 px-3 w-10 font-mono text-[11px] text-center border border-black bg-orange-600 text-white font-bold">
                          #
                        </th>
                        {targetHeaders.map((h, i) => (
                          <th
                            key={i}
                            className="py-2.5 px-3 whitespace-nowrap text-center border border-black font-bold text-white bg-orange-500 tracking-wider"
                          >
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedTargetRows.map((row, rIdx) => {
                        const isOrange = rIdx % 2 === 1; // 0=White, 1=Orange, 2=White
                        return (
                          <tr
                            key={rIdx}
                            className={isOrange ? 'bg-[#FFEDD5]' : 'bg-white'}
                          >
                            <td className="py-2 px-3 text-slate-600 font-mono text-[11px] text-center border border-black font-medium">
                              {(page - 1) * rowsPerPage + rIdx + 1}
                            </td>
                            {targetHeaders.map((h, cIdx) => {
                              const rawVal = row[h];
                              const strVal = rawVal !== undefined && rawVal !== null ? String(rawVal) : '';
                              const isNum =
                                strVal !== '' &&
                                /^-?\d+(\.\d+)?$/.test(strVal) &&
                                !(strVal.startsWith('0') && strVal.length > 1 && !strVal.startsWith('0.'));

                              const isMoney = /金額|貨款|單價|價格|總計|小計|Price|Amount/i.test(h);

                              return (
                                <td
                                  key={cIdx}
                                  className={`py-2 px-3 whitespace-nowrap text-black border border-black font-medium max-w-xs truncate ${
                                    isNum ? 'text-right font-mono' : 'text-left'
                                  }`}
                                >
                                  {isMoney && isNum
                                    ? `$${Number(strVal).toLocaleString()}`
                                    : strVal}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        ) : viewMode === 'targetOnly' ? (
          /* Target Only Full Width */
          <div className="p-4 overflow-x-auto">
            <div className="border-2 border-black rounded-lg overflow-x-auto bg-white shadow-xs">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-orange-500 text-white font-bold">
                    <th className="py-2.5 px-3 w-12 font-mono text-center border border-black bg-orange-600 text-white font-bold">
                      #
                    </th>
                    {targetHeaders.map((h, i) => (
                      <th key={i} className="py-2.5 px-3 whitespace-nowrap text-center border border-black font-bold text-white bg-orange-500 tracking-wider">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {paginatedTargetRows.map((row, rIdx) => {
                    const isOrange = rIdx % 2 === 1;
                    return (
                      <tr key={rIdx} className={isOrange ? 'bg-[#FFEDD5]' : 'bg-white'}>
                        <td className="py-2 px-3 text-slate-600 font-mono text-[11px] text-center border border-black font-medium">
                          {(page - 1) * rowsPerPage + rIdx + 1}
                        </td>
                        {targetHeaders.map((h, cIdx) => (
                          <td key={cIdx} className="py-2 px-3 whitespace-nowrap text-black font-medium border border-black">
                            {row[h] !== undefined && row[h] !== null ? String(row[h]) : ''}
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          /* Source Only Full Width */
          <div className="p-4 overflow-x-auto">
            <div className="border border-slate-200 rounded-lg overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-100 border-b border-slate-200 text-slate-800 font-semibold">
                    <th className="py-2 px-3 w-12 font-mono text-slate-400">#</th>
                    {(sheetA?.headers || []).map((h, i) => (
                      <th key={i} className="py-2 px-3 whitespace-nowrap">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {paginatedSourceRows.map((row, rIdx) => (
                    <tr key={rIdx} className="hover:bg-slate-50 transition-colors">
                      <td className="py-2 px-3 text-slate-400 font-mono text-[11px]">
                        {(page - 1) * rowsPerPage + rIdx + 1}
                      </td>
                      {(sheetA?.headers || []).map((h, cIdx) => (
                        <td key={cIdx} className="py-2 px-3 whitespace-nowrap text-slate-700">
                          {row[h] !== undefined && row[h] !== null ? String(row[h]) : ''}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* Pagination Footer */}
      <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-600">
        <div className="font-mono tabular-nums">
          顯示第 {(page - 1) * rowsPerPage + 1} 至{' '}
          {Math.min(page * rowsPerPage, filteredTargetRows.length)} 筆（共{' '}
          {filteredTargetRows.length} 筆）
        </div>
        <div className="flex items-center gap-1.5">
          <button
            disabled={page <= 1}
            onClick={() => setPage(page - 1)}
            className="px-2.5 py-1 rounded border border-slate-200 bg-white hover:bg-slate-100 disabled:opacity-40 transition-colors"
          >
            上一頁
          </button>
          <span className="font-mono px-2">
            {page} / {totalPages}
          </span>
          <button
            disabled={page >= totalPages}
            onClick={() => setPage(page + 1)}
            className="px-2.5 py-1 rounded border border-slate-200 bg-white hover:bg-slate-100 disabled:opacity-40 transition-colors"
          >
            下一頁
          </button>
        </div>
      </div>
    </div>
  );
};
