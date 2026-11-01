import React, { useState } from 'react';
import { Sparkles, Loader2, Wand2, HelpCircle, Check, AlertCircle } from 'lucide-react';
import { ExcelSheetData, MappingRule } from '../types/excel';

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
  currentRules,
  onApplyRules,
  disabled,
}) => {
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [userInstructions, setUserInstructions] = useState('');
  const [showPromptInput, setShowPromptInput] = useState(false);
  const [aiSummary, setAiSummary] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const handleRunAiAlignment = async () => {
    if (!sheetA) return;

    setIsAnalyzing(true);
    setErrorMessage(null);

    try {
      const response = await fetch('/api/ai/analyze-mapping', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sourceHeaders: sheetA.headers,
          sampleSourceRows: sheetA.rows.slice(0, 3),
          targetHeaders: sheetB ? sheetB.headers : [],
          sampleTargetRows: sheetB ? sheetB.rows.slice(0, 3) : [],
          userInstructions: userInstructions.trim() || undefined,
        }),
      });

      if (!response.ok) {
        const errJson = await response.json();
        throw new Error(errJson.error || errJson.details || '伺服器分析失敗');
      }

      const result = await response.json();
      if (result.success && result.data && result.data.mappingRules) {
        const aiRules: MappingRule[] = result.data.mappingRules.map((r: any, idx: number) => ({
          id: `ai-rule-${Date.now()}-${idx}`,
          targetColumn: r.targetColumn,
          ruleType: r.ruleType || 'direct',
          sourceColumns: Array.isArray(r.sourceColumns) ? r.sourceColumns : [],
          defaultValue: r.defaultValue || '',
          separator: r.separator !== undefined ? r.separator : ' ',
          splitIndex: r.splitIndex || 0,
          formula: r.formula || '',
          datePattern: r.datePattern || 'YYYY-MM-DD',
          explanation: r.explanation || '',
          trimWhitespace: true,
          enabled: true,
        }));

        setAiSummary(result.data.summary);
        onApplyRules(aiRules, result.data.summary);
      } else {
        throw new Error('未取得有效的對應規則');
      }
    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || 'AI 分析對應發生錯誤，請重試');
    } finally {
      setIsAnalyzing(false);
    }
  };

  if (!sheetA) return null;

  return (
    <div className="bg-gradient-to-r from-emerald-950 via-slate-900 to-indigo-950 text-white rounded-xl p-5 border border-slate-800 shadow-sm">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-start gap-3.5">
          <div className="w-10 h-10 rounded-lg bg-emerald-500/20 border border-emerald-500/40 text-emerald-400 flex items-center justify-center shrink-0">
            <Sparkles className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h4 className="text-base font-semibold text-white">Gemini 智慧對齊與語意轉譯</h4>
              <span className="text-xs bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 px-2 py-0.5 rounded">
                AI 欄位推導
              </span>
            </div>
            <p className="text-xs text-slate-300 mt-1 max-w-2xl">
              {sheetB
                ? `自動理解「${sheetA.fileName}」與「${sheetB.fileName}」的語意差異，自動推論欄位合併、常數補齊、算式與格式轉譯。`
                : `目標 Excel B 未指定：AI 將自動依據 Excel A 的資料內容，推導出一套標準化的高品質結構。`}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            onClick={() => setShowPromptInput(!showPromptInput)}
            className="text-xs text-slate-300 hover:text-white px-3 py-2 rounded-lg border border-slate-700 hover:border-slate-600 transition-colors"
          >
            {showPromptInput ? '收起指示' : '自訂轉譯指示...'}
          </button>
          <button
            disabled={isAnalyzing || disabled}
            onClick={handleRunAiAlignment}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold bg-emerald-500 hover:bg-emerald-400 text-slate-950 rounded-lg transition-all disabled:opacity-50 shadow-sm"
          >
            {isAnalyzing ? (
              <>
                <Loader2 className="w-4 h-4 animate-spin" />
                正在深度分析資料結構...
              </>
            ) : (
              <>
                <Wand2 className="w-4 h-4" />
                執行 AI 智慧對齊
              </>
            )}
          </button>
        </div>
      </div>

      {/* Optional custom instructions */}
      {showPromptInput && (
        <div className="mt-4 pt-4 border-t border-slate-800">
          <label className="block text-xs font-medium text-slate-300 mb-1.5">
            提供給 AI 的客製轉譯要求（選填）：
          </label>
          <input
            type="text"
            value={userInstructions}
            onChange={(e) => setUserInstructions(e.target.value)}
            placeholder="例如：將『縣市』與『地址』合併；計算『數量*單價』填入總金額；將民國日期轉為西元 YYYY-MM-DD"
            className="w-full text-xs px-3 py-2 bg-slate-900 border border-slate-700 rounded-lg text-white placeholder-slate-500 focus:outline-none focus:border-emerald-500"
          />
        </div>
      )}

      {/* AI Summary report */}
      {aiSummary && (
        <div className="mt-4 p-3 bg-emerald-950/60 border border-emerald-500/30 rounded-lg text-xs text-emerald-200 flex items-start gap-2.5">
          <Check className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
          <div>
            <span className="font-semibold text-emerald-300">AI 分析完成：</span> {aiSummary}
          </div>
        </div>
      )}

      {/* Error report */}
      {errorMessage && (
        <div className="mt-4 p-3 bg-rose-950/60 border border-rose-500/30 rounded-lg text-xs text-rose-200 flex items-start gap-2.5">
          <AlertCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
          <div>
            <span className="font-semibold text-rose-300">發生錯誤：</span> {errorMessage}
          </div>
        </div>
      )}
    </div>
  );
};
