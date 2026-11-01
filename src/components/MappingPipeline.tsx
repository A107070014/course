import React, { useState } from 'react';
import {
  ArrowRight,
  Plus,
  Trash2,
  Edit2,
  Sparkles,
  BookmarkPlus,
  CheckCircle2,
  Shuffle,
  Layers,
  Info,
} from 'lucide-react';
import { MappingRule, RuleType } from '../types/excel';

interface MappingPipelineProps {
  rules: MappingRule[];
  sourceHeaders: string[];
  targetHeaders: string[];
  onUpdateRule: (updatedRule: MappingRule) => void;
  onDeleteRule: (ruleId: string) => void;
  onAddRule: () => void;
  onOpenEditModal: (rule: MappingRule) => void;
  onSaveAsTemplate: (name: string, description: string) => void;
  onAutoMatchByNames: () => void;
}

export const MappingPipeline: React.FC<MappingPipelineProps> = ({
  rules,
  sourceHeaders,
  targetHeaders,
  onUpdateRule,
  onDeleteRule,
  onAddRule,
  onOpenEditModal,
  onSaveAsTemplate,
  onAutoMatchByNames,
}) => {
  const [showSaveTemplateModal, setShowSaveTemplateModal] = useState(false);
  const [templateName, setTemplateName] = useState('');
  const [templateDesc, setTemplateDesc] = useState('');

  const getRuleBadge = (type: RuleType) => {
    switch (type) {
      case 'direct':
        return <span className="text-[11px] font-mono text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">直接複製</span>;
      case 'combine':
        return <span className="text-[11px] font-mono text-blue-700 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">欄位合併</span>;
      case 'constant':
        return <span className="text-[11px] font-mono text-purple-700 bg-purple-50 px-2 py-0.5 rounded border border-purple-200">固定常數</span>;
      case 'math':
        return <span className="text-[11px] font-mono text-amber-700 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">算術公式</span>;
      case 'date_format':
        return <span className="text-[11px] font-mono text-cyan-700 bg-cyan-50 px-2 py-0.5 rounded border border-cyan-200">日期格式</span>;
      case 'split':
        return <span className="text-[11px] font-mono text-orange-700 bg-orange-50 px-2 py-0.5 rounded border border-orange-200">文字拆分</span>;
      case 'lookup':
        return <span className="text-[11px] font-mono text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200">代碼對照</span>;
      case 'case_trim':
        return <span className="text-[11px] font-mono text-slate-700 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">去除空格</span>;
      default:
        return <span className="text-[11px] font-mono text-slate-600 bg-slate-50 px-2 py-0.5 rounded">對應</span>;
    }
  };

  const handleSaveTemplateSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!templateName.trim()) return;
    onSaveAsTemplate(templateName.trim(), templateDesc.trim());
    setShowSaveTemplateModal(false);
    setTemplateName('');
    setTemplateDesc('');
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-6 space-y-6">
      {/* Action Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-200">
        <div>
          <h3 className="text-base font-semibold text-slate-900">
            轉換管道與欄位對齊清單 ({rules.filter((r) => r.enabled).length}/{rules.length} 啟用中)
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            定義從 Excel A 產出 Excel B 每一欄的轉譯管道，可任意拖曳微調或新增目標欄位
          </p>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            onClick={onAutoMatchByNames}
            className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors inline-flex items-center gap-1.5"
            title="依同名或相似名稱自動連線"
          >
            <Shuffle className="w-3.5 h-3.5" />
            快速相似名稱配對
          </button>
          <button
            onClick={onAddRule}
            className="px-3 py-1.5 text-xs font-semibold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors inline-flex items-center gap-1.5"
          >
            <Plus className="w-3.5 h-3.5" />
            新增目標欄位
          </button>
          <button
            onClick={() => setShowSaveTemplateModal(true)}
            className="px-3 py-1.5 text-xs font-semibold text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors inline-flex items-center gap-1.5"
          >
            <BookmarkPlus className="w-3.5 h-3.5 text-slate-600" />
            儲存為新範本
          </button>
        </div>
      </div>

      {/* Rules List */}
      {rules.length === 0 ? (
        <div className="text-center py-12 bg-slate-50 rounded-xl border border-dashed border-slate-200">
          <Layers className="w-10 h-10 text-slate-400 mx-auto mb-2.5" />
          <h4 className="text-sm font-semibold text-slate-800">尚未建立任何欄位對應規則</h4>
          <p className="text-xs text-slate-500 mt-1 max-w-sm mx-auto">
            可點擊上方「AI 智慧對齊」自動推導，或點擊「新增目標欄位」手動設定。
          </p>
          <button
            onClick={onAddRule}
            className="mt-4 px-4 py-2 text-xs font-semibold bg-emerald-600 text-white rounded-lg hover:bg-emerald-700"
          >
            手動新增第一個目標欄位
          </button>
        </div>
      ) : (
        <div className="space-y-2.5">
          {rules.map((rule, idx) => (
            <div
              key={rule.id}
              className={`flex flex-col md:flex-row md:items-center justify-between p-3.5 rounded-lg border transition-all ${
                rule.enabled
                  ? 'bg-white border-slate-200 hover:border-slate-300 shadow-xs'
                  : 'bg-slate-50/70 border-slate-200 opacity-60'
              }`}
            >
              {/* Left: Source -> Rule -> Target */}
              <div className="flex-1 flex flex-wrap items-center gap-2 sm:gap-4">
                {/* Index & Toggle */}
                <div className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    checked={rule.enabled}
                    onChange={(e) => onUpdateRule({ ...rule, enabled: e.target.checked })}
                    className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                    title={rule.enabled ? '已啟用 (點擊可停用此欄)' : '已停用'}
                  />
                  <span className="text-xs font-mono text-slate-400 w-5">
                    {String(idx + 1).padStart(2, '0')}.
                  </span>
                </div>

                {/* Target Column (Excel B) */}
                <div className="flex items-center gap-2 min-w-[150px]">
                  <span className="w-4 h-4 rounded-full bg-indigo-100 text-indigo-700 text-[10px] font-bold flex items-center justify-center">
                    B
                  </span>
                  <span className="text-sm font-semibold text-slate-900">{rule.targetColumn}</span>
                </div>

                {/* Arrow */}
                <div className="flex items-center text-slate-400 px-1">
                  <ArrowRight className="w-4 h-4" />
                </div>

                {/* Rule Type Badge */}
                <div>{getRuleBadge(rule.ruleType)}</div>

                {/* Source Column(s) (Excel A) or Constant */}
                <div className="flex items-center gap-1.5 flex-wrap">
                  {rule.ruleType === 'constant' ? (
                    <span className="text-xs font-mono px-2 py-0.5 bg-slate-100 text-slate-800 rounded">
                      常數值："{rule.defaultValue || ''}"
                    </span>
                  ) : rule.sourceColumns.length > 0 ? (
                    rule.sourceColumns.map((colName) => (
                      <span
                        key={colName}
                        className="inline-flex items-center gap-1 text-xs px-2 py-0.5 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded font-mono"
                      >
                        <span className="text-[10px] text-emerald-600 font-bold">A:</span>
                        {colName}
                      </span>
                    ))
                  ) : (
                    <span className="text-xs text-amber-600 italic">未指定來源欄位</span>
                  )}

                  {rule.formula && (
                    <span className="text-[11px] font-mono text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded">
                      算式: {rule.formula}
                    </span>
                  )}
                  {rule.separator !== undefined && rule.ruleType === 'combine' && (
                    <span className="text-[11px] font-mono text-slate-400">
                      (分隔符: "{rule.separator}")
                    </span>
                  )}
                  {rule.datePattern && rule.ruleType === 'date_format' && (
                    <span className="text-[11px] font-mono text-slate-500">
                      [{rule.datePattern}]
                    </span>
                  )}
                </div>

                {/* Explanation info if any */}
                {rule.explanation && (
                  <span className="text-xs text-slate-400 hidden xl:inline-block max-w-[200px] truncate">
                    {rule.explanation}
                  </span>
                )}
              </div>

              {/* Right: Actions */}
              <div className="flex items-center gap-1.5 mt-2 md:mt-0 shrink-0 self-end md:self-center">
                <button
                  onClick={() => onOpenEditModal(rule)}
                  className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-md transition-colors"
                  title="詳細編輯此欄位規則"
                >
                  <Edit2 className="w-4 h-4" />
                </button>
                <button
                  onClick={() => onDeleteRule(rule.id)}
                  className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-md transition-colors"
                  title="刪除此目標欄位"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Save Template Modal */}
      {showSaveTemplateModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 p-4">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6 animate-in fade-in zoom-in-95">
            <h4 className="text-base font-semibold text-slate-900 mb-2">儲存為轉換範本</h4>
            <p className="text-xs text-slate-500 mb-4">
              儲存目前設定好的 {rules.length} 個欄位轉譯規則，日後直接上傳 Excel A 即可一鍵自動轉出！
            </p>

            <form onSubmit={handleSaveTemplateSubmit} className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  範本名稱
                </label>
                <input
                  type="text"
                  required
                  placeholder="例如：蝦皮訂單轉黑貓物流單"
                  value={templateName}
                  onChange={(e) => setTemplateName(e.target.value)}
                  className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600 font-medium"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  範本備註說明 (選填)
                </label>
                <textarea
                  rows={2}
                  placeholder="簡短描述這個範本的適用情境..."
                  value={templateDesc}
                  onChange={(e) => setTemplateDesc(e.target.value)}
                  className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600 text-slate-700"
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setShowSaveTemplateModal(false)}
                  className="px-3 py-2 text-xs text-slate-600 hover:bg-slate-100 rounded-lg"
                >
                  取消
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow-sm"
                >
                  確認儲存範本
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
