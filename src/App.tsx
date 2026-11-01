import React, { useState, useEffect, useMemo } from 'react';
import {
  FileSpreadsheet,
  Sparkles,
  Download,
  ArrowRight,
  Layers,
  Settings,
  CheckCircle2,
  RefreshCw,
  FolderKanban,
  SlidersHorizontal,
  Table as TableIcon,
  HelpCircle,
  Zap,
} from 'lucide-react';
import { ExcelSheetData, MappingRule, ConversionTemplate } from './types/excel';
import {
  parseExcelFile,
  exportToXLSXWithTemplate,
  exportToCSV,
  getDemoSampleDatasets,
  generateStyledSampleBBuffer,
} from './utils/excelParser';
import { executeTransformation } from './utils/ruleEngine';
import { PRESET_TEMPLATES } from './utils/presetTemplates';
import { Header } from './components/Header';
import { FileDropzone } from './components/FileDropzone';
import { AiAlignBanner } from './components/AiAlignBanner';
import { MappingPipeline } from './components/MappingPipeline';
import { DataPreviewTable } from './components/DataPreviewTable';
import { TemplateManager } from './components/TemplateManager';
import { RuleEditorModal } from './components/RuleEditorModal';

const STORAGE_KEY_TEMPLATES = 'excel_morph_custom_templates_v1';

export default function App() {
  const [sheetA, setSheetA] = useState<ExcelSheetData | null>(null);
  const [sheetB, setSheetB] = useState<ExcelSheetData | null>(null);
  const [rules, setRules] = useState<MappingRule[]>([]);
  const [activeTab, setActiveTab] = useState<'convert' | 'pipeline' | 'preview' | 'templates'>('convert');
  const [editingRule, setEditingRule] = useState<MappingRule | null>(null);
  const [customTemplates, setCustomTemplates] = useState<ConversionTemplate[]>([]);
  const [activeTemplateId, setActiveTemplateId] = useState<string | null>(null);
  const [toast, setToast] = useState<string | null>(null);
  const [selectedPresetId, setSelectedPresetId] = useState<string>('preset-ecommerce-to-logistics');

  // Load custom templates from localStorage
  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_TEMPLATES);
      if (saved) {
        setCustomTemplates(JSON.parse(saved));
      }
    } catch (e) {
      console.error('Failed to load custom templates from localStorage', e);
    }
  }, []);

  const showToast = (msg: string) => {
    setToast(msg);
    setTimeout(() => {
      setToast((curr) => (curr === msg ? null : curr));
    }, 3500);
  };

  // Save custom templates
  const saveCustomTemplates = (templates: ConversionTemplate[]) => {
    setCustomTemplates(templates);
    try {
      localStorage.setItem(STORAGE_KEY_TEMPLATES, JSON.stringify(templates));
    } catch (e) {
      console.error('Failed to save templates to localStorage', e);
    }
  };

  // Execute transformation in real-time
  const { targetHeaders, targetRows, stats } = useMemo(() => {
    if (!sheetA || sheetA.rows.length === 0 || rules.length === 0) {
      return { targetHeaders: [], targetRows: [], stats: null };
    }
    return executeTransformation(sheetA.rows, rules);
  }, [sheetA, rules]);

  // When Excel A is selected
  const handleFileASelected = async (file: File) => {
    try {
      const parsed = await parseExcelFile(file);
      setSheetA(parsed);
      showToast(`已成功讀取來源 Excel A: ${parsed.fileName} (${parsed.totalRows} 筆資料)`);

      // If rules are empty and sheetB is loaded, auto-init rules for sheetB
      if (rules.length === 0 && sheetB && sheetB.headers.length > 0) {
        const initRules = generateDefaultRulesForTarget(sheetB.headers, parsed.headers);
        setRules(initRules);
      } else if (rules.length === 0 && activeTemplateId) {
        // template already active
      }
    } catch (err: any) {
      alert(`讀取 Excel A 失敗: ${err.message}`);
    }
  };

  // When Excel B (Target Template) is selected
  const handleFileBSelected = async (file: File) => {
    try {
      const parsed = await parseExcelFile(file);
      setSheetB(parsed);
      showToast(`已成功讀取目標範本 Excel B: ${parsed.fileName} (${parsed.headers.length} 欄位)`);

      // Initialize rules targeting Sheet B's columns
      const sourceHeaders = sheetA ? sheetA.headers : [];
      const initRules = generateDefaultRulesForTarget(parsed.headers, sourceHeaders);
      setRules(initRules);
      setActiveTemplateId(null);
    } catch (err: any) {
      alert(`讀取 Excel B 失敗: ${err.message}`);
    }
  };

  // Change header row index for Excel A
  const handleChangeHeaderRowA = async (rowIndex: number) => {
    if (!sheetA || !sheetA.rawBuffer) return;
    try {
      const dummyFile = new File([sheetA.rawBuffer], sheetA.fileName);
      const reParsed = await parseExcelFile(dummyFile, sheetA.activeSheet, rowIndex);
      setSheetA(reParsed);
      showToast(`已將 Excel A 表頭切換至第 ${rowIndex + 1} 列 (${reParsed.headers.length} 欄位)`);
    } catch (e: any) {
      alert(`切換 Excel A 表頭失敗: ${e.message}`);
    }
  };

  // Change header row index for Excel B
  const handleChangeHeaderRowB = async (rowIndex: number) => {
    if (!sheetB || !sheetB.rawBuffer) return;
    try {
      const dummyFile = new File([sheetB.rawBuffer], sheetB.fileName);
      const reParsed = await parseExcelFile(dummyFile, sheetB.activeSheet, rowIndex);
      setSheetB(reParsed);
      const initRules = generateDefaultRulesForTarget(reParsed.headers, sheetA ? sheetA.headers : []);
      setRules(initRules);
      showToast(`已將 Excel B 表頭切換至第 ${rowIndex + 1} 列 (${reParsed.headers.length} 欄位)`);
    } catch (e: any) {
      alert(`切換 Excel B 表頭失敗: ${e.message}`);
    }
  };

  // Generate basic rules from target headers
  const generateDefaultRulesForTarget = (
    targets: string[],
    sources: string[]
  ): MappingRule[] => {
    return targets.map((tCol, idx) => {
      // Find exact or substring match
      const exactMatch = sources.find((s) => s.trim() === tCol.trim());
      const subMatch = !exactMatch
        ? sources.find(
            (s) =>
              s.includes(tCol) ||
              tCol.includes(s) ||
              s.replace(/[\s_()（）]/g, '') === tCol.replace(/[\s_()（）]/g, '')
          )
        : undefined;

      const matchedSource = exactMatch || subMatch;

      return {
        id: `rule-${Date.now()}-${idx}`,
        targetColumn: tCol,
        ruleType: matchedSource ? 'direct' : 'constant',
        sourceColumns: matchedSource ? [matchedSource] : [],
        defaultValue: matchedSource ? '' : '',
        separator: ' ',
        enabled: true,
      };
    });
  };

  // Quick name matching
  const handleAutoMatchByNames = () => {
    if (!sheetA) {
      alert('請先上傳來源 Excel A');
      return;
    }
    const updated = rules.map((r) => {
      const exact = sheetA.headers.find((s) => s.trim() === r.targetColumn.trim());
      const sub = !exact
        ? sheetA.headers.find(
            (s) =>
              s.includes(r.targetColumn) ||
              r.targetColumn.includes(s) ||
              s.replace(/[\s_()（）]/g, '') === r.targetColumn.replace(/[\s_()（）]/g, '')
          )
        : undefined;

      const match = exact || sub;
      if (match) {
        return {
          ...r,
          ruleType: 'direct' as const,
          sourceColumns: [match],
          enabled: true,
        };
      }
      return r;
    });

    setRules(updated);
    showToast('已依欄位名稱相似度完成快速配對！');
  };

  // Apply a template (preset or custom)
  const handleApplyTemplate = (tpl: ConversionTemplate) => {
    setRules(tpl.rules);
    setActiveTemplateId(tpl.id);
    showToast(`已成功套用範本：「${tpl.name}」`);
    setActiveTab('convert');
  };

  // Load demo dataset (Sample A + Sample B + Logistics Preset)
  const handleLoadDemo = async () => {
    const { sampleA, sampleB } = getDemoSampleDatasets();
    const styledBuffer = await generateStyledSampleBBuffer();
    sampleB.rawBuffer = styledBuffer;

    setSheetA(sampleA);
    setSheetB(sampleB);

    // Load matching preset
    const preset = PRESET_TEMPLATES[0]; // Ecommerce to Logistics
    setRules(preset.rules);
    setActiveTemplateId(preset.id);
    setSelectedPresetId(preset.id);

    showToast('已載入示範資料！Excel B 範本樣式（第 1 列為表頭、第 2 列起繼承幣別與斑馬紋）已就緒');
  };

  // Export handlers
  const handleExportXLSX = async () => {
    if (targetRows.length === 0) {
      alert('目前沒有已轉譯的目標資料可供匯出');
      return;
    }
    const filename = sheetB?.fileName
      ? `轉出_${sheetB.fileName}`
      : sheetA?.fileName
      ? `${sheetA.fileName.replace(/\.[^/.]+$/, '')}_轉譯結果.xlsx`
      : 'excel_b_converted.xlsx';

    await exportToXLSXWithTemplate(targetHeaders, targetRows, sheetB, filename);
    showToast(`已 100% 還原 Excel B 範例樣式成功匯出：${filename}`);
  };

  const handleExportCSV = () => {
    if (targetRows.length === 0) return;
    const filename = sheetB?.fileName
      ? `轉出_${sheetB.fileName.replace(/\.[^/.]+$/, '')}.csv`
      : 'excel_b_converted.csv';
    exportToCSV(targetHeaders, targetRows, filename);
    showToast(`已匯出 ${filename}`);
  };

  // Rule operations
  const handleUpdateRule = (updated: MappingRule) => {
    setRules((prev) => prev.map((r) => (r.id === updated.id ? updated : r)));
  };

  const handleDeleteRule = (id: string) => {
    setRules((prev) => prev.filter((r) => r.id !== id));
  };

  const handleAddRule = () => {
    const newRule: MappingRule = {
      id: `rule-${Date.now()}`,
      targetColumn: `新欄位_${rules.length + 1}`,
      ruleType: 'direct',
      sourceColumns: sheetA && sheetA.headers.length > 0 ? [sheetA.headers[0]] : [],
      enabled: true,
    };
    setRules((prev) => [...prev, newRule]);
    setEditingRule(newRule);
  };

  // Save as template
  const handleSaveAsTemplate = (name: string, description: string) => {
    const newTpl: ConversionTemplate = {
      id: `custom-tpl-${Date.now()}`,
      name,
      description,
      createdAt: new Date().toISOString().split('T')[0],
      targetHeaders: rules.map((r) => r.targetColumn),
      rules,
    };
    const updated = [newTpl, ...customTemplates];
    saveCustomTemplates(updated);
    setActiveTemplateId(newTpl.id);
    showToast(`已將目前規則儲存為新範本：「${name}」`);
  };

  const handleDeleteCustomTemplate = (id: string) => {
    const updated = customTemplates.filter((t) => t.id !== id);
    saveCustomTemplates(updated);
    showToast('已刪除指定範本');
  };

  const handleImportTemplates = (imported: ConversionTemplate[]) => {
    const merged = [...imported, ...customTemplates];
    // deduplicate by id
    const unique = Array.from(new Map(merged.map((item) => [item.id, item])).values());
    saveCustomTemplates(unique);
    showToast(`成功匯入 ${imported.length} 個範本！`);
  };

  const handleExportAllTemplates = () => {
    const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(customTemplates, null, 2));
    const a = document.createElement('a');
    a.href = dataStr;
    a.download = `excel_morph_templates_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col text-slate-900 antialiased">
      {/* Toast Notification */}
      {toast && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900 text-white text-xs px-4 py-3 rounded-lg shadow-lg border border-slate-700 flex items-center gap-2.5 animate-in fade-in slide-in-from-bottom-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toast}</span>
        </div>
      )}

      {/* Top Bar Header */}
      <Header
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        canExport={targetRows.length > 0}
        onExportXLSX={handleExportXLSX}
        onExportCSV={handleExportCSV}
        onLoadDemo={handleLoadDemo}
        hasData={!!sheetA}
      />

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
        {/* Hero Concept & Quick Scenario Selector */}
        <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="max-w-2xl">
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
              Excel A 丟進來，自動轉出 Excel B
            </h1>
            <p className="text-sm text-slate-500 mt-1.5 leading-relaxed">
              解決不同系統、供應商、平台間格式完全不相容的問題。支援 AI
              語意對齊、欄位拆分與合併、四則運算、代碼映射與自訂範本，一鍵無痛轉檔。
            </p>
          </div>

          {/* Quick Preset Selector for instant workflow */}
          <div className="bg-slate-50 border border-slate-200 p-3.5 rounded-lg shrink-0 flex flex-col gap-2">
            <span className="text-xs font-semibold text-slate-700 flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-indigo-600" />
              快速選擇常用轉譯情境範本：
            </span>
            <div className="flex items-center gap-2">
              <select
                value={selectedPresetId}
                onChange={(e) => {
                  setSelectedPresetId(e.target.value);
                  const found = PRESET_TEMPLATES.find((p) => p.id === e.target.value);
                  if (found) handleApplyTemplate(found);
                }}
                className="text-xs bg-white border border-slate-300 rounded px-3 py-1.5 font-medium text-slate-800 focus:outline-none focus:border-emerald-600 max-w-xs truncate"
              >
                {PRESET_TEMPLATES.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.name}
                  </option>
                ))}
              </select>
              <button
                onClick={() => {
                  const found = PRESET_TEMPLATES.find((p) => p.id === selectedPresetId);
                  if (found) handleApplyTemplate(found);
                }}
                className="px-3 py-1.5 text-xs font-semibold bg-emerald-600 text-white rounded hover:bg-emerald-700 transition-colors whitespace-nowrap"
              >
                套用
              </button>
            </div>
          </div>
        </div>

        {/* Tab 1: Fast Convert & Upload */}
        {activeTab === 'convert' && (
          <div className="space-y-8">
            {/* File Dropzone */}
            <FileDropzone
              sheetA={sheetA}
              sheetB={sheetB}
              onFileASelected={handleFileASelected}
              onFileBSelected={handleFileBSelected}
              onClearA={() => {
                setSheetA(null);
                setRules([]);
              }}
              onClearB={() => {
                setSheetB(null);
              }}
              onChangeHeaderRowA={handleChangeHeaderRowA}
              onChangeHeaderRowB={handleChangeHeaderRowB}
              onQuickLoadDemo={handleLoadDemo}
            />

            {/* AI Auto-Align Banner */}
            {sheetA && (
              <AiAlignBanner
                sheetA={sheetA}
                sheetB={sheetB}
                currentRules={rules}
                onApplyRules={(newRules, summary) => {
                  setRules(newRules);
                  showToast('已成功套用 AI 智慧欄位對齊！');
                }}
              />
            )}

            {/* Quick summary of rules if configured */}
            {sheetA && rules.length > 0 && (
              <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                    <h3 className="text-sm font-semibold text-slate-900">
                      轉譯管道已就緒：共 {rules.filter((r) => r.enabled).length} 個目標欄位
                    </h3>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => setActiveTab('pipeline')}
                      className="text-xs font-semibold text-slate-700 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 px-3 py-1.5 rounded-lg transition-colors inline-flex items-center gap-1"
                    >
                      <SlidersHorizontal className="w-3.5 h-3.5" />
                      自訂微調規則
                    </button>
                    <button
                      onClick={() => setActiveTab('preview')}
                      className="text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 px-3 py-1.5 rounded-lg transition-colors inline-flex items-center gap-1"
                    >
                      <TableIcon className="w-3.5 h-3.5" />
                      查看完整雙表預覽
                    </button>
                  </div>
                </div>

                {/* Live Data Preview */}
                <DataPreviewTable
                  sheetA={sheetA}
                  sheetB={sheetB}
                  targetHeaders={targetHeaders}
                  targetRows={targetRows}
                  stats={stats}
                  onExportXLSX={handleExportXLSX}
                  onExportCSV={handleExportCSV}
                />
              </div>
            )}
          </div>
        )}

        {/* Tab 2: Visual Mapping Pipeline */}
        {activeTab === 'pipeline' && (
          <div className="space-y-6">
            {sheetA && (
              <AiAlignBanner
                sheetA={sheetA}
                sheetB={sheetB}
                currentRules={rules}
                onApplyRules={(newRules) => {
                  setRules(newRules);
                  showToast('已套用 AI 對齊規則！');
                }}
              />
            )}

            <MappingPipeline
              rules={rules}
              sourceHeaders={sheetA?.headers || []}
              targetHeaders={targetHeaders}
              onUpdateRule={handleUpdateRule}
              onDeleteRule={handleDeleteRule}
              onAddRule={handleAddRule}
              onOpenEditModal={(rule) => setEditingRule(rule)}
              onSaveAsTemplate={handleSaveAsTemplate}
              onAutoMatchByNames={handleAutoMatchByNames}
            />

            {targetRows.length > 0 && (
              <div className="pt-2">
                <DataPreviewTable
                  sheetA={sheetA}
                  sheetB={sheetB}
                  targetHeaders={targetHeaders}
                  targetRows={targetRows}
                  stats={stats}
                  onExportXLSX={handleExportXLSX}
                  onExportCSV={handleExportCSV}
                />
              </div>
            )}
          </div>
        )}

        {/* Tab 3: Detailed Data Preview Comparison */}
        {activeTab === 'preview' && (
          <div className="space-y-6">
            <DataPreviewTable
              sheetA={sheetA}
              sheetB={sheetB}
              targetHeaders={targetHeaders}
              targetRows={targetRows}
              stats={stats}
              onExportXLSX={handleExportXLSX}
              onExportCSV={handleExportCSV}
            />
          </div>
        )}

        {/* Tab 4: Template Manager */}
        {activeTab === 'templates' && (
          <div className="space-y-6">
            <TemplateManager
              customTemplates={customTemplates}
              onApplyTemplate={handleApplyTemplate}
              onDeleteCustomTemplate={handleDeleteCustomTemplate}
              onImportTemplates={handleImportTemplates}
              onExportAllTemplates={handleExportAllTemplates}
            />
          </div>
        )}
      </main>

      {/* Rule Editor Modal */}
      {editingRule && (
        <RuleEditorModal
          rule={editingRule}
          sourceHeaders={sheetA?.headers || []}
          sampleSourceRow={sheetA?.rows[0]}
          onSave={handleUpdateRule}
          onClose={() => setEditingRule(null)}
        />
      )}

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 mt-16 py-6 text-center text-xs text-slate-500">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-4">
          <p>© 2026 Excel Morph. 支援 XLSX, XLS 與 CSV 異質格式智慧自動轉譯。</p>
          <div className="flex items-center gap-4 text-xs">
            <span>本機記憶體即時運算</span>
            <span>·</span>
            <span>UTF-8 with BOM 相容模式</span>
            <span>·</span>
            <span>Powered by Gemini 3.8 Flash</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
