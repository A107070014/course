import React, { useState } from 'react';
import { X, Check, Calculator, Calendar, ArrowRightLeft, Sparkles, Layers, Sliders } from 'lucide-react';
import { MappingRule, RuleType } from '../types/excel';

interface RuleEditorModalProps {
  rule: MappingRule | null;
  sourceHeaders: string[];
  sampleSourceRow?: Record<string, any>;
  onSave: (updatedRule: MappingRule) => void;
  onClose: () => void;
}

export const RuleEditorModal: React.FC<RuleEditorModalProps> = ({
  rule,
  sourceHeaders,
  sampleSourceRow,
  onSave,
  onClose,
}) => {
  if (!rule) return null;

  const [form, setForm] = useState<MappingRule>({ ...rule });
  const [lookupKeyInput, setLookupKeyInput] = useState('');
  const [lookupValInput, setLookupValInput] = useState('');

  const ruleOptions: { type: RuleType; label: string; desc: string }[] = [
    { type: 'direct', label: '直接複製 (Direct)', desc: '直接取來源單一欄位的值' },
    { type: 'combine', label: '欄位合併 (Combine)', desc: '將多個來源欄位依序以特定符號串接' },
    { type: 'constant', label: '固定常數 (Constant)', desc: '所有列填入固定預設值或代碼' },
    { type: 'split', label: '文字拆分 (Split)', desc: '依符號拆解文字並提取第 N 部分' },
    { type: 'math', label: '數學運算 (Math)', desc: '針對數值進行四則運算 (如 單價*數量)' },
    { type: 'date_format', label: '日期轉換 (Date)', desc: '標準化為 YYYY-MM-DD 或民國年轉西元' },
    { type: 'lookup', label: '對照字典 (Lookup)', desc: '將特定舊值替換為新代碼 (如 男->M)' },
    { type: 'case_trim', label: '清除空格 (Trim)', desc: '修除前後空白或大寫' },
  ];

  const handleToggleSourceColumn = (col: string) => {
    if (form.ruleType === 'direct' || form.ruleType === 'split' || form.ruleType === 'date_format') {
      setForm({ ...form, sourceColumns: [col] });
    } else {
      const exists = form.sourceColumns.includes(col);
      const updated = exists
        ? form.sourceColumns.filter((c) => c !== col)
        : [...form.sourceColumns, col];
      setForm({ ...form, sourceColumns: updated });
    }
  };

  const handleAddLookup = () => {
    if (!lookupKeyInput.trim()) return;
    const current = { ...(form.lookupMap || {}) };
    current[lookupKeyInput.trim()] = lookupValInput.trim();
    setForm({ ...form, lookupMap: current });
    setLookupKeyInput('');
    setLookupValInput('');
  };

  const handleRemoveLookup = (key: string) => {
    const current = { ...(form.lookupMap || {}) };
    delete current[key];
    setForm({ ...form, lookupMap: current });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4">
      <div className="bg-white rounded-xl shadow-xl border border-slate-200 w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-slate-50">
          <div>
            <h3 className="text-base font-semibold text-slate-900">
              編輯轉換規則：
              <span className="text-emerald-700 font-mono ml-1">{form.targetColumn}</span>
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">自訂目標表格欄位的來源對應與計算邏輯</p>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-slate-600 p-1 rounded-md hover:bg-slate-200 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Content Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          {/* Target Column Name */}
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1.5">
              目標欄位名稱 (Excel B)
            </label>
            <input
              type="text"
              value={form.targetColumn}
              onChange={(e) => setForm({ ...form, targetColumn: e.target.value })}
              className="w-full text-sm px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600 font-medium"
            />
          </div>

          {/* Rule Type Selector */}
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-2">
              轉換邏輯類型 (Rule Type)
            </label>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
              {ruleOptions.map((opt) => (
                <button
                  key={opt.type}
                  type="button"
                  onClick={() => {
                    setForm({
                      ...form,
                      ruleType: opt.type,
                      // if switching to single-col, keep only first col
                      sourceColumns:
                        ['direct', 'split', 'date_format'].includes(opt.type) &&
                        form.sourceColumns.length > 1
                          ? [form.sourceColumns[0]]
                          : form.sourceColumns,
                    });
                  }}
                  className={`p-2.5 rounded-lg border text-left transition-all ${
                    form.ruleType === opt.type
                      ? 'border-emerald-600 bg-emerald-50/60 ring-1 ring-emerald-600'
                      : 'border-slate-200 bg-white hover:border-slate-300'
                  }`}
                >
                  <p
                    className={`text-xs font-semibold ${
                      form.ruleType === opt.type ? 'text-emerald-900' : 'text-slate-800'
                    }`}
                  >
                    {opt.label.split(' ')[0]}
                  </p>
                  <p className="text-[11px] text-slate-500 mt-0.5 line-clamp-1">{opt.desc}</p>
                </button>
              ))}
            </div>
          </div>

          {/* Source Columns Selection (except constant) */}
          {form.ruleType !== 'constant' && (
            <div>
              <div className="flex items-center justify-between mb-2">
                <label className="text-xs font-semibold text-slate-700">
                  選擇來源欄位 (Excel A)
                  {form.ruleType === 'combine' || form.ruleType === 'math' ? (
                    <span className="text-slate-500 font-normal ml-1">
                      (可多選，點擊可依序加入)
                    </span>
                  ) : (
                    <span className="text-slate-500 font-normal ml-1">(單選)</span>
                  )}
                </label>
                {form.sourceColumns.length > 0 && (
                  <button
                    onClick={() => setForm({ ...form, sourceColumns: [] })}
                    className="text-xs text-slate-500 hover:text-slate-700 underline"
                  >
                    清除重選
                  </button>
                )}
              </div>

              <div className="flex flex-wrap gap-1.5 max-h-36 overflow-y-auto p-2 bg-slate-50 border border-slate-200 rounded-lg">
                {sourceHeaders.length === 0 ? (
                  <p className="text-xs text-slate-400 py-1">尚未載入來源欄位</p>
                ) : (
                  sourceHeaders.map((col) => {
                    const isSelected = form.sourceColumns.includes(col);
                    const orderIdx = form.sourceColumns.indexOf(col);
                    return (
                      <button
                        key={col}
                        type="button"
                        onClick={() => handleToggleSourceColumn(col)}
                        className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded text-xs transition-colors ${
                          isSelected
                            ? 'bg-emerald-600 text-white font-medium shadow-xs'
                            : 'bg-white border border-slate-300 text-slate-700 hover:bg-slate-100'
                        }`}
                      >
                        {isSelected && orderIdx >= 0 && (
                          <span className="w-3.5 h-3.5 rounded-full bg-emerald-800 text-[10px] flex items-center justify-center font-mono">
                            {orderIdx + 1}
                          </span>
                        )}
                        <span>{col}</span>
                      </button>
                    );
                  })
                )}
              </div>
            </div>
          )}

          {/* Dynamic Configuration per Rule Type */}
          {form.ruleType === 'combine' && (
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  串接分隔符號 (Separator)
                </label>
                <input
                  type="text"
                  value={form.separator !== undefined ? form.separator : ''}
                  onChange={(e) => setForm({ ...form, separator: e.target.value })}
                  placeholder="如 空白、破折號 -、或留空直接連接"
                  className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600 font-mono"
                />
                <p className="text-[11px] text-slate-400 mt-1">
                  例如合併縣市與地址請填留空；合併姓名可填半形空格
                </p>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  缺漏時預設值 (Default Value)
                </label>
                <input
                  type="text"
                  value={form.defaultValue || ''}
                  onChange={(e) => setForm({ ...form, defaultValue: e.target.value })}
                  placeholder="來源皆為空時填入"
                  className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                />
              </div>
            </div>
          )}

          {form.ruleType === 'constant' && (
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                固定常數值 (Constant Value)
              </label>
              <input
                type="text"
                value={form.defaultValue || ''}
                onChange={(e) => setForm({ ...form, defaultValue: e.target.value })}
                placeholder="例如：常溫、TW、待處理、0"
                className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
              />
              <p className="text-[11px] text-slate-400 mt-1">所有轉換列均會填入此特定值</p>
            </div>
          )}

          {form.ruleType === 'split' && (
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  拆分符號 (Delimiter)
                </label>
                <input
                  type="text"
                  value={form.separator || ' '}
                  onChange={(e) => setForm({ ...form, separator: e.target.value })}
                  className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg font-mono focus:outline-none focus:border-emerald-600"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  選取段落索引 (0 為第一段)
                </label>
                <input
                  type="number"
                  min="0"
                  max="10"
                  value={form.splitIndex || 0}
                  onChange={(e) =>
                    setForm({ ...form, splitIndex: parseInt(e.target.value, 10) || 0 })
                  }
                  className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg font-mono focus:outline-none focus:border-emerald-600"
                />
              </div>
            </div>
          )}

          {form.ruleType === 'math' && (
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                運算公式 (Formula)
              </label>
              <input
                type="text"
                value={form.formula || ''}
                onChange={(e) => setForm({ ...form, formula: e.target.value })}
                placeholder="例如：A * B 或 A * 1.05 或 A / 32.5 (A 代表第一欄，B 代表第二欄)"
                className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg font-mono focus:outline-none focus:border-emerald-600"
              />
              <p className="text-[11px] text-slate-400 mt-1">
                若不填寫公式，預設會將選取的所有來源欄位相乘
              </p>
            </div>
          )}

          {form.ruleType === 'date_format' && (
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                目標日期格式 (Date Pattern)
              </label>
              <select
                value={form.datePattern || 'YYYY-MM-DD'}
                onChange={(e) => setForm({ ...form, datePattern: e.target.value })}
                className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg bg-white focus:outline-none focus:border-emerald-600 font-mono"
              >
                <option value="YYYY-MM-DD">YYYY-MM-DD (例如 2026-04-10)</option>
                <option value="YYYY/MM/DD">YYYY/MM/DD (例如 2026/04/10)</option>
                <option value="YYYY年MM月DD日">YYYY年MM月DD日</option>
                <option value="YYYYMMDD">YYYYMMDD (純數字無符號)</option>
              </select>
              <p className="text-[11px] text-slate-400 mt-1">
                自動支援辨識標準西元日期、ISO 時間與民國年格式
              </p>
            </div>
          )}

          {form.ruleType === 'lookup' && (
            <div className="space-y-3">
              <label className="block text-xs font-semibold text-slate-700">
                數值對照表 (Value Mapping)
              </label>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  placeholder="來源舊值 (如 信用卡線上付款)"
                  value={lookupKeyInput}
                  onChange={(e) => setLookupKeyInput(e.target.value)}
                  className="flex-1 text-xs px-3 py-1.5 border border-slate-300 rounded-md"
                />
                <span className="text-slate-400 text-xs">→</span>
                <input
                  type="text"
                  placeholder="替換目標值 (如 0 或 CREDIT)"
                  value={lookupValInput}
                  onChange={(e) => setLookupValInput(e.target.value)}
                  className="flex-1 text-xs px-3 py-1.5 border border-slate-300 rounded-md"
                />
                <button
                  type="button"
                  onClick={handleAddLookup}
                  className="px-3 py-1.5 text-xs bg-slate-800 text-white rounded-md hover:bg-slate-700 transition-colors"
                >
                  新增
                </button>
              </div>

              {form.lookupMap && Object.keys(form.lookupMap).length > 0 && (
                <div className="border border-slate-200 rounded-md divide-y divide-slate-100 max-h-32 overflow-y-auto">
                  {Object.entries(form.lookupMap).map(([k, v]) => (
                    <div key={k} className="flex items-center justify-between px-3 py-1.5 text-xs">
                      <span>
                        <span className="font-mono text-slate-700">{k}</span>
                        <span className="text-slate-400 mx-1.5">→</span>
                        <span className="font-mono text-emerald-700 font-semibold">{v}</span>
                      </span>
                      <button
                        onClick={() => handleRemoveLookup(k)}
                        className="text-slate-400 hover:text-rose-500"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Clean Options */}
          <div className="flex items-center gap-6 pt-2 border-t border-slate-100">
            <label className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
              <input
                type="checkbox"
                checked={form.trimWhitespace !== false}
                onChange={(e) => setForm({ ...form, trimWhitespace: e.target.checked })}
                className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
              />
              自動去除前後空白 (Trim)
            </label>

            <div className="flex items-center gap-2 text-xs text-slate-700">
              <span>大小寫：</span>
              <select
                value={form.textCase || 'none'}
                onChange={(e) => setForm({ ...form, textCase: e.target.value as any })}
                className="text-xs bg-white border border-slate-200 rounded px-2 py-0.5"
              >
                <option value="none">保持原樣</option>
                <option value="upper">全部大寫 (UPPER)</option>
                <option value="lower">全部小寫 (lower)</option>
              </select>
            </div>
          </div>

          {/* Explanation */}
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">
              轉換備註 / 說明
            </label>
            <input
              type="text"
              value={form.explanation || ''}
              onChange={(e) => setForm({ ...form, explanation: e.target.value })}
              placeholder="例如：訂購人姓名轉收件客戶"
              className="w-full text-xs px-3 py-2 border border-slate-300 rounded-lg text-slate-600 focus:outline-none"
            />
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between px-6 py-4 border-t border-slate-200 bg-slate-50">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-xs font-medium text-slate-700 hover:bg-slate-200 rounded-lg transition-colors"
          >
            取消
          </button>
          <button
            type="button"
            onClick={() => {
              onSave(form);
              onClose();
            }}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition-colors shadow-sm"
          >
            <Check className="w-4 h-4" />
            確認套用規則
          </button>
        </div>
      </div>
    </div>
  );
};
