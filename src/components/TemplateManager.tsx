import React, { useRef } from 'react';
import {
  Bookmark,
  CheckCircle2,
  Trash2,
  Download,
  Upload,
  ArrowRight,
  Layers,
  Sparkles,
} from 'lucide-react';
import { ConversionTemplate } from '../types/excel';
import { PRESET_TEMPLATES } from '../utils/presetTemplates';

interface TemplateManagerProps {
  customTemplates: ConversionTemplate[];
  onApplyTemplate: (template: ConversionTemplate) => void;
  onDeleteCustomTemplate: (templateId: string) => void;
  onImportTemplates: (templates: ConversionTemplate[]) => void;
  onExportAllTemplates: () => void;
}

export const TemplateManager: React.FC<TemplateManagerProps> = ({
  customTemplates,
  onApplyTemplate,
  onDeleteCustomTemplate,
  onImportTemplates,
  onExportAllTemplates,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleImportFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const json = JSON.parse(event.target?.result as string);
        const imported = Array.isArray(json) ? json : [json];
        onImportTemplates(imported);
      } catch (err) {
        alert('範本檔案格式錯誤，必須為合法的 JSON 格式。');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  return (
    <div className="space-y-8">
      {/* Top Banner & Actions */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 p-5 bg-white border border-slate-200 rounded-xl shadow-xs">
        <div>
          <h3 className="text-base font-semibold text-slate-900">轉換範本庫 (Conversion Templates)</h3>
          <p className="text-xs text-slate-500 mt-0.5">
            儲存常用的 Excel A → Excel B 對應藍圖，下次直接拖放 Excel A 即可秒速完成轉譯！
          </p>
        </div>

        <div className="flex items-center gap-2">
          <input
            ref={fileInputRef}
            type="file"
            accept=".json"
            className="hidden"
            onChange={handleImportFile}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors inline-flex items-center gap-1.5"
          >
            <Upload className="w-3.5 h-3.5" />
            匯入範本 (JSON)
          </button>
          <button
            onClick={onExportAllTemplates}
            className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors inline-flex items-center gap-1.5"
          >
            <Download className="w-3.5 h-3.5" />
            備份所有範本
          </button>
        </div>
      </div>

      {/* User Custom Saved Templates */}
      <div>
        <div className="flex items-center gap-2 mb-3">
          <Bookmark className="w-4 h-4 text-emerald-600" />
          <h4 className="text-sm font-semibold text-slate-900">我的自訂範本 (Custom Templates)</h4>
          <span className="text-xs text-slate-500">({customTemplates.length} 個已儲存)</span>
        </div>

        {customTemplates.length === 0 ? (
          <div className="p-8 text-center bg-white rounded-xl border border-dashed border-slate-200 text-xs text-slate-500">
            尚未儲存任何自訂範本。在「對應規則」頁面完成欄位設定後，點擊「儲存為新範本」即可記錄於此。
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {customTemplates.map((tpl) => (
              <div
                key={tpl.id}
                className="p-5 bg-white rounded-xl border border-slate-200 hover:border-slate-300 shadow-xs flex flex-col justify-between transition-all"
              >
                <div>
                  <div className="flex items-start justify-between gap-2">
                    <h5 className="text-sm font-semibold text-slate-900">{tpl.name}</h5>
                    <button
                      onClick={() => onDeleteCustomTemplate(tpl.id)}
                      className="text-slate-400 hover:text-rose-600 p-1 rounded"
                      title="刪除此範本"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                  {tpl.description && (
                    <p className="text-xs text-slate-500 mt-1 line-clamp-2">{tpl.description}</p>
                  )}
                  <div className="mt-3 text-xs text-slate-600 space-y-1">
                    <div className="flex items-center gap-1.5">
                      <Layers className="w-3.5 h-3.5 text-slate-400" />
                      <span>
                        目標欄位：<strong>{tpl.targetHeaders.length}</strong> 個
                      </span>
                    </div>
                    <div className="text-[11px] text-slate-400">
                      儲存時間：{tpl.createdAt || '近期'}
                    </div>
                  </div>
                </div>

                <div className="mt-5 pt-3 border-t border-slate-100 flex items-center justify-between">
                  <span className="text-[11px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-100">
                    自訂藍圖
                  </span>
                  <button
                    onClick={() => onApplyTemplate(tpl)}
                    className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors"
                  >
                    套用此範本
                    <ArrowRight className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Built-in Preset Templates */}
      <div>
        <div className="flex items-center gap-2 mb-3">
          <Sparkles className="w-4 h-4 text-indigo-600" />
          <h4 className="text-sm font-semibold text-slate-900">系統內建商用情境範本</h4>
          <span className="text-xs text-slate-500">(即點即用)</span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {PRESET_TEMPLATES.map((tpl) => (
            <div
              key={tpl.id}
              className="p-5 bg-white rounded-xl border border-slate-200 hover:border-slate-300 shadow-xs flex flex-col justify-between transition-all"
            >
              <div>
                <h5 className="text-sm font-semibold text-slate-900">{tpl.name}</h5>
                <p className="text-xs text-slate-500 mt-1.5 leading-relaxed">{tpl.description}</p>
                <div className="mt-3 text-xs text-slate-600 space-y-1">
                  <div className="flex items-center gap-1.5">
                    <Layers className="w-3.5 h-3.5 text-slate-400" />
                    <span>
                      包含 <strong>{tpl.targetHeaders.length}</strong> 個目標欄位與自動轉譯規則
                    </span>
                  </div>
                </div>
              </div>

              <div className="mt-5 pt-3 border-t border-slate-100 flex items-center justify-between">
                <span className="text-[11px] text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-100">
                  官方預設
                </span>
                <button
                  onClick={() => onApplyTemplate(tpl)}
                  className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
                >
                  套用此範本
                  <ArrowRight className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};
