import React, { useState } from 'react';
import { Sparkles, Wand2, Check, CheckCircle2, Zap } from 'lucide-react';
import { ExcelSheetData, MappingRule } from '../types/excel';
import { generateSmartClientMapping } from '../utils/smartMatcher';

interface AiAlignBannerProps {
  sheetA: ExcelSheetData | null;
  sheetB: ExcelSheetData | null;
  currentRules: MappingRule[];
  onApplyRules: (rules: MappingRule[], summary?: string) => void;
  disabled?: boolean;
}

export const AiAlignBanner: React.FC<AiAlignBannerProps> = ({
  sheetA,
  sheetB,
  onApplyRules,
  disabled,
}) => {
  const [matchSummary, setMatchSummary] = useState<string | null>(null);

  const handleRunSmartAlignment = () => {
    if (!sheetA) return;

    const targetHeaders = sheetB && sheetB.headers.length > 0 ? sheetB.headers : sheetA.headers;
    const { rules, summary } = generateSmartClientMapping(sheetA.headers, targetHeaders);

    setMatchSummary(summary);
    onApplyRules(rules, summary);
  };

  if (!sheetA) return null;

  return (
    <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-xl p-5 border border-indigo-900/60 shadow-xs">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-start gap-3.5">
          <div className="w-10 h-10 rounded-lg bg-indigo-500/20 border border-indigo-500/40 text-indigo-400 flex items-center justify-center shrink-0">
            <Zap className="w-5 h-5 text-emerald-400" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h4 className="text-base font-semibold text-white">智慧自動欄位對齊與規則推導</h4>
              <span className="text-xs bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 px-2 py-0.5 rounded font-medium">
                純前端靜態運算 · 無需後端
              </span>
            </div>
            <p className="text-xs text-slate-300 mt-1 max-w-2xl leading-relaxed">
              {sheetB
                ? `本機自動對齊「${sheetA.fileName}」與「${sheetB.fileName}」：支援地址三段合併、品名件數組合、同義詞庫對應與金額四則運算，完全於瀏覽器內執行。`
                : `自動分析 Excel A 的欄位命名與型態，在瀏覽器端直接建立精準的對照規則。`}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            disabled={disabled}
            onClick={handleRunSmartAlignment}
            className="inline-flex items-center gap-2 px-4 py-2.5 text-xs font-semibold bg-emerald-500 hover:bg-emerald-400 text-slate-950 rounded-lg transition-all disabled:opacity-50 shadow-sm cursor-pointer"
          >
            <Wand2 className="w-4 h-4" />
            一鍵智慧自動對齊
          </button>
        </div>
      </div>

      {/* Match Summary notification */}
      {matchSummary && (
        <div className="mt-4 p-3 bg-emerald-950/60 border border-emerald-500/40 rounded-lg text-xs text-emerald-200 flex items-start gap-2.5">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
          <div>
            <span className="font-semibold text-emerald-300">配對完成：</span> {matchSummary}
          </div>
        </div>
      )}
    </div>
  );
};
