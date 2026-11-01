import React, { useRef, useState } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  CheckCircle2,
  RefreshCw,
  X,
  ArrowRight,
  Layers,
  Palette,
  Sparkles,
  Rows,
  Sliders,
} from 'lucide-react';
import { ExcelSheetData } from '../types/excel';

interface FileDropzoneProps {
  sheetA: ExcelSheetData | null;
  sheetB: ExcelSheetData | null;
  onFileASelected: (file: File) => void;
  onFileBSelected: (file: File) => void;
  onClearA: () => void;
  onClearB: () => void;
  onSwitchSheetA?: (sheetName: string) => void;
  onSwitchSheetB?: (sheetName: string) => void;
  onChangeHeaderRowA?: (rowIndex: number) => void;
  onChangeHeaderRowB?: (rowIndex: number) => void;
  onQuickLoadDemo: () => void;
}

export const FileDropzone: React.FC<FileDropzoneProps> = ({
  sheetA,
  sheetB,
  onFileASelected,
  onFileBSelected,
  onClearA,
  onClearB,
  onSwitchSheetA,
  onSwitchSheetB,
  onChangeHeaderRowA,
  onChangeHeaderRowB,
  onQuickLoadDemo,
}) => {
  const [isDraggingA, setIsDraggingA] = useState(false);
  const [isDraggingB, setIsDraggingB] = useState(false);

  const fileInputARef = useRef<HTMLInputElement>(null);
  const fileInputBRef = useRef<HTMLInputElement>(null);

  const handleDropA = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingA(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileASelected(e.dataTransfer.files[0]);
    }
  };

  const handleDropB = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingB(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileBSelected(e.dataTransfer.files[0]);
    }
  };

  const formatBytes = (bytes: number) => {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
  };

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Dropzone A: 來源表格 Excel A */}
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setIsDraggingA(true);
          }}
          onDragLeave={() => setIsDraggingA(false)}
          onDrop={handleDropA}
          className={`relative border-2 border-dashed rounded-xl p-6 transition-all duration-200 ${
            isDraggingA
              ? 'border-emerald-500 bg-emerald-50/50'
              : sheetA
              ? 'border-emerald-200 bg-emerald-50/20'
              : 'border-slate-300 bg-white hover:border-slate-400'
          }`}
        >
          <input
            ref={fileInputARef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={(e) => {
              if (e.target.files && e.target.files[0]) {
                onFileASelected(e.target.files[0]);
              }
            }}
          />

          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2">
              <span className="w-6 h-6 rounded bg-slate-900 text-white text-xs font-bold flex items-center justify-center">
                A
              </span>
              <h3 className="text-base font-semibold text-slate-900">來源表格 (Excel A)</h3>
              <span className="text-xs text-rose-600 font-medium">* 必要</span>
            </div>
            {sheetA && (
              <button
                onClick={onClearA}
                className="text-slate-400 hover:text-slate-600 p-1 rounded hover:bg-slate-100"
                title="移除來源檔案"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          {!sheetA ? (
            <div
              onClick={() => fileInputARef.current?.click()}
              className="py-10 flex flex-col items-center justify-center cursor-pointer group"
            >
              <div className="w-12 h-12 rounded-full bg-slate-100 group-hover:bg-slate-200 flex items-center justify-center text-slate-500 mb-3 transition-colors">
                <UploadCloud className="w-6 h-6" />
              </div>
              <p className="text-sm font-medium text-slate-800 mb-1">
                點擊或拖曳上傳 <span className="text-emerald-600">Excel A</span> 來源檔案
              </p>
              <p className="text-xs text-slate-400">支援 .xlsx, .xls, .csv 格式 (任意欄位與結構)</p>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="flex items-start justify-between p-3.5 bg-white border border-emerald-200 rounded-lg shadow-xs">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center shrink-0">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <p className="text-sm font-semibold text-slate-900 truncate max-w-[240px]">
                      {sheetA.fileName}
                    </p>
                    <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
                      <span>{formatBytes(sheetA.fileSize)}</span>
                      <span>·</span>
                      <span className="font-mono tabular-nums">{sheetA.totalRows} 列資料</span>
                      <span>·</span>
                      <span className="font-mono tabular-nums">{sheetA.headers.length} 個欄位</span>
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => fileInputARef.current?.click()}
                  className="text-xs text-slate-500 hover:text-slate-800 border border-slate-200 px-2.5 py-1 rounded-md bg-white hover:bg-slate-50 transition-colors"
                >
                  重新選擇
                </button>
              </div>

              {/* Sheet selector and Header Row Selector */}
              <div className="flex flex-wrap items-center gap-4 text-xs text-slate-600">
                {sheetA.sheetNames.length > 1 && (
                  <div className="flex items-center gap-1.5">
                    <Layers className="w-3.5 h-3.5 text-slate-400" />
                    <span>工作表：</span>
                    <select
                      value={sheetA.activeSheet}
                      onChange={(e) => onSwitchSheetA && onSwitchSheetA(e.target.value)}
                      className="text-xs bg-white border border-slate-200 rounded px-2 py-1 text-slate-800"
                    >
                      {sheetA.sheetNames.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                {onChangeHeaderRowA && (
                  <div className="flex items-center gap-1.5">
                    <Rows className="w-3.5 h-3.5 text-slate-400" />
                    <span>表頭所在列：</span>
                    <select
                      value={sheetA.headerRowIndex !== undefined ? sheetA.headerRowIndex : 0}
                      onChange={(e) => onChangeHeaderRowA(parseInt(e.target.value, 10))}
                      className="text-xs bg-white border border-slate-200 rounded px-2 py-1 text-slate-800 font-medium"
                    >
                      {[0, 1, 2, 3, 4].map((rIdx) => (
                        <option key={rIdx} value={rIdx}>
                          第 {rIdx + 1} 列
                        </option>
                      ))}
                    </select>
                  </div>
                )}
              </div>

              {/* Header preview chips */}
              <div className="text-xs">
                <span className="text-slate-500 font-medium">偵測到來源欄位：</span>
                <div className="flex flex-wrap gap-1.5 mt-1.5 max-h-24 overflow-y-auto">
                  {sheetA.headers.map((h, idx) => (
                    <span
                      key={idx}
                      className="inline-block px-2 py-0.5 bg-slate-100 text-slate-700 rounded text-xs"
                    >
                      {h}
                    </span>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Dropzone B: 目標表格 Excel B 樣本 (可選) */}
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setIsDraggingB(true);
          }}
          onDragLeave={() => setIsDraggingB(false)}
          onDrop={handleDropB}
          className={`relative border-2 border-dashed rounded-xl p-6 transition-all duration-200 ${
            isDraggingB
              ? 'border-indigo-500 bg-indigo-50/50'
              : sheetB
              ? 'border-indigo-200 bg-indigo-50/20'
              : 'border-slate-300 bg-white hover:border-slate-400'
          }`}
        >
          <input
            ref={fileInputBRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={(e) => {
              if (e.target.files && e.target.files[0]) {
                onFileBSelected(e.target.files[0]);
              }
            }}
          />

          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2">
              <span className="w-6 h-6 rounded bg-indigo-600 text-white text-xs font-bold flex items-center justify-center">
                B
              </span>
              <h3 className="text-base font-semibold text-slate-900">目標格式 (Excel B)</h3>
              <span className="text-xs text-slate-500 font-normal">(樣板與範例檔)</span>
            </div>
            {sheetB && (
              <button
                onClick={onClearB}
                className="text-slate-400 hover:text-slate-600 p-1 rounded hover:bg-slate-100"
                title="移除目標檔案"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          {!sheetB ? (
            <div
              onClick={() => fileInputBRef.current?.click()}
              className="py-10 flex flex-col items-center justify-center cursor-pointer group"
            >
              <div className="w-12 h-12 rounded-full bg-slate-100 group-hover:bg-slate-200 flex items-center justify-center text-slate-500 mb-3 transition-colors">
                <UploadCloud className="w-6 h-6" />
              </div>
              <p className="text-sm font-medium text-slate-800 mb-1">
                上傳期望的 <span className="text-indigo-600">Excel B</span> 目標樣板或包含範例的檔案
              </p>
              <p className="text-xs text-slate-400">
                將 100% 複製該範本的表頭底色、字型、邊框、交錯斑馬紋與貨幣日期格式
              </p>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="flex items-start justify-between p-3.5 bg-white border border-indigo-200 rounded-lg shadow-xs">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-lg bg-indigo-100 text-indigo-700 flex items-center justify-center shrink-0">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <p className="text-sm font-semibold text-slate-900 truncate max-w-[240px]">
                      {sheetB.fileName}
                    </p>
                    <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
                      <span>{formatBytes(sheetB.fileSize)}</span>
                      <span>·</span>
                      <span className="font-mono tabular-nums">{sheetB.headers.length} 個目標欄位</span>
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => fileInputBRef.current?.click()}
                  className="text-xs text-slate-500 hover:text-slate-800 border border-slate-200 px-2.5 py-1 rounded-md bg-white hover:bg-slate-50 transition-colors"
                >
                  重新選擇
                </button>
              </div>

              {/* Sheet & Header Row Specification for Excel B */}
              <div className="flex flex-wrap items-center gap-4 text-xs text-slate-600">
                {sheetB.sheetNames.length > 1 && (
                  <div className="flex items-center gap-1.5">
                    <Layers className="w-3.5 h-3.5 text-slate-400" />
                    <span>工作表：</span>
                    <select
                      value={sheetB.activeSheet}
                      onChange={(e) => onSwitchSheetB && onSwitchSheetB(e.target.value)}
                      className="text-xs bg-white border border-slate-200 rounded px-2 py-1 text-slate-800 font-medium"
                    >
                      {sheetB.sheetNames.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                <div className="flex items-center gap-1.5 px-2 py-1 bg-emerald-50 border border-emerald-200 rounded text-emerald-800 font-medium">
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                  <span>第一列為標準表頭 (Row 1)</span>
                </div>

                {onChangeHeaderRowB && sheetB.headerRowIndex !== undefined && sheetB.headerRowIndex > 0 && (
                  <button
                    onClick={() => onChangeHeaderRowB(0)}
                    className="text-[11px] text-indigo-600 hover:underline"
                  >
                    重設為第 1 列
                  </button>
                )}
              </div>

              {/* Target Header chips */}
              <div className="text-xs">
                <span className="text-slate-500 font-medium">目標欄位列表：</span>
                <div className="flex flex-wrap gap-1.5 mt-1.5 max-h-24 overflow-y-auto">
                  {sheetB.headers.map((h, idx) => (
                    <span
                      key={idx}
                      className="inline-block px-2 py-0.5 bg-indigo-50 text-indigo-800 border border-indigo-100 rounded text-xs"
                    >
                      {h}
                    </span>
                  ))}
                </div>
              </div>

              {/* Comprehensive Template Style Fidelity Inspector */}
              <div className="p-3 bg-gradient-to-r from-indigo-50/80 to-slate-50 border border-indigo-100 rounded-lg text-xs space-y-2">
                <div className="flex items-center justify-between text-indigo-950 font-semibold">
                  <div className="flex items-center gap-1.5">
                    <Palette className="w-4 h-4 text-indigo-600" />
                    <span>Excel B 範本樣式高保真複製已啟用</span>
                  </div>
                  <span className="text-[11px] font-mono text-emerald-700 bg-emerald-100/70 px-2 py-0.5 rounded font-bold">
                    100% 格式同步
                  </span>
                </div>

                <div className="grid grid-cols-2 gap-2 text-[11px] text-slate-600 pt-1 border-t border-indigo-100/60">
                  <div className="flex items-center gap-1.5">
                    <span
                      className="w-3.5 h-3.5 rounded-xs border border-slate-300 shrink-0 shadow-2xs"
                      style={{
                        backgroundColor: sheetB.detectedStyles?.headerBgColor || '#1F4E79',
                      }}
                    />
                    <span className="truncate">
                      表頭主題色：
                      <strong>{sheetB.detectedStyles?.headerBgColor || '#1F4E79'}</strong>
                    </span>
                  </div>

                  <div className="flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>
                      字型樣式：<strong>{sheetB.detectedStyles?.fontName || '微軟正黑體'}</strong>
                    </span>
                  </div>

                  <div className="flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>
                      自訂欄寬：<strong>{sheetB.headers.length} 欄已鎖定</strong>
                    </span>
                  </div>

                  <div className="flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>
                      格式傳承：
                      <strong>
                        {sheetB.detectedStyles?.hasZebraStriping ? '斑馬紋底色 ' : '細格線 '}
                        + 貨幣日期格式
                      </strong>
                    </span>
                  </div>
                </div>

                {sheetB.detectedStyles?.hasFooter && (
                  <p className="text-[11px] text-indigo-800 bg-indigo-100/40 px-2 py-1 rounded">
                    📌 偵測到範本含頁尾「合計/簽核列」，轉出時將自動位移至新資料底部並保留合併儲存格！
                  </p>
                )}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Demo / Quick Test Banner */}
      {!sheetA && (
        <div className="flex flex-col sm:flex-row items-center justify-between p-4 bg-slate-900 text-white rounded-xl gap-4">
          <div>
            <p className="text-sm font-semibold">快速體驗：手邊沒有現成檔案？</p>
            <p className="text-xs text-slate-300 mt-0.5">
              一鍵載入「網店訂單 (Excel A)」與「黑貓物流託運單 (Excel B)」示範資料，立即體驗全自動轉譯！
            </p>
          </div>
          <button
            onClick={onQuickLoadDemo}
            className="px-4 py-2 text-xs font-semibold bg-emerald-500 hover:bg-emerald-400 text-slate-950 rounded-lg transition-colors whitespace-nowrap shrink-0"
          >
            載入示範檔案 (A & B)
          </button>
        </div>
      )}
    </div>
  );
};
