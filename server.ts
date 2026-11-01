import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import dotenv from 'dotenv';
import { GoogleGenAI, Type } from '@google/genai';

dotenv.config();

const app = express();
const port = 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Server-side Gemini initialization
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      'User-Agent': 'aistudio-build',
    },
  },
});

// API: AI-Powered Smart Mapping Analysis
app.post('/api/ai/analyze-mapping', async (req, res) => {
  try {
    const {
      sourceHeaders,
      sampleSourceRows,
      targetHeaders,
      sampleTargetRows,
      userInstructions,
    } = req.body;

    if (!sourceHeaders || !Array.isArray(sourceHeaders)) {
      return res.status(400).json({ error: 'Missing sourceHeaders' });
    }

    const prompt = `你是一個專業的資料處理與 Excel 表格工程專家。
使用者的目標是將來源表格「Excel A」自動轉換為目標表格「Excel B」。
兩者的欄位名稱與資料結構不同，甚至可能需要欄位合併、拆分、常數填寫、計算或格式轉換。

【來源表格 Excel A】
欄位列表：${JSON.stringify(sourceHeaders)}
前 3 筆樣本資料：
${JSON.stringify(sampleSourceRows || [], null, 2)}

【目標表格 Excel B】
${
  targetHeaders && targetHeaders.length > 0
    ? `目標欄位列表：${JSON.stringify(targetHeaders)}
目標樣本資料：
${JSON.stringify(sampleTargetRows || [], null, 2)}`
    : `目標表格未提供現成檔案，請根據來源資料自動推導出乾淨、標準化的商業目標欄位架構。`
}

${userInstructions ? `【使用者額外指示】：${userInstructions}` : ''}

請為每一個目標欄位（Excel B）決定最精準的對應規則。
規則類型說明（ruleType）：
- "direct": 直接複製來源單一欄位值 (sourceColumns 為單一元素陣列)
- "combine": 合併多個來源欄位 (sourceColumns 為多個欄位，separator 為分隔符號，例如 " " 或 "")
- "constant": 固定常數或預設值 (defaultValue，如 "已確認", "TW", 0)
- "split": 拆分來源欄位 (sourceColumns 為單一欄位，separator 為拆分符號，splitIndex 為取第幾個部分 0-indexed)
- "math": 數值運算 (如 sourceColumns 為 ["單價", "數量"], formula 為 "A * B" 或 "A * 1.05")
- "date_format": 日期格式化 (如 "YYYY-MM-DD", "YYYY/MM/DD", "民國轉西元")
- "lookup": 值映射字典 (如 mappingDict: {"男":"M", "女":"F"})
- "case_trim": 去除頭尾空格或大小寫轉換

請返回 JSON 格式，包含 mappingRules 陣列與 summary 簡短總結。`;

    const response = await ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: prompt,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            summary: {
              type: Type.STRING,
              description: '繁體中文的轉換邏輯簡短說明與分析重點',
            },
            mappingRules: {
              type: Type.ARRAY,
              description: '針對每一個目標欄位的轉換規則',
              items: {
                type: Type.OBJECT,
                properties: {
                  targetColumn: {
                    type: Type.STRING,
                    description: '目標欄位名稱 (Excel B)',
                  },
                  ruleType: {
                    type: Type.STRING,
                    description:
                      '規則類型: direct, combine, constant, split, math, date_format, lookup, case_trim',
                  },
                  sourceColumns: {
                    type: Type.ARRAY,
                    items: { type: Type.STRING },
                    description: '對應的來源欄位名稱 (Excel A)',
                  },
                  defaultValue: {
                    type: Type.STRING,
                    description: '預設值或常數值',
                  },
                  separator: {
                    type: Type.STRING,
                    description: '合併或拆分用的分隔符號',
                  },
                  splitIndex: {
                    type: Type.INTEGER,
                    description: '拆分時選取的索引值 (0 開始)',
                  },
                  formula: {
                    type: Type.STRING,
                    description: '運算公式或算式說明',
                  },
                  datePattern: {
                    type: Type.STRING,
                    description: '目標日期格式，如 YYYY-MM-DD',
                  },
                  explanation: {
                    type: Type.STRING,
                    description: '此欄位對齊理由與注意事項',
                  },
                },
                required: ['targetColumn', 'ruleType', 'sourceColumns'],
              },
            },
          },
          required: ['summary', 'mappingRules'],
        },
      },
    });

    const parsed = JSON.parse(response.text || '{}');
    res.json({ success: true, data: parsed });
  } catch (error: any) {
    console.error('AI analyze mapping error:', error);
    res.status(500).json({
      error: 'AI 分析對應失敗，請稍後重試',
      details: error?.message || String(error),
    });
  }
});

// API: Batch AI Smart Row Assist (for complex semantic transformations)
app.post('/api/ai/batch-transform', async (req, res) => {
  try {
    const { rows, instruction, outputColumnName } = req.body;
    if (!rows || !Array.isArray(rows) || !instruction) {
      return res.status(400).json({ error: 'Missing rows or instruction' });
    }

    const limitedRows = rows.slice(0, 50); // limit per request for safety
    const prompt = `依據使用者指示處理以下每一筆資料行，並產出目標欄位「${outputColumnName || 'Result'}」的結果。
指示：${instruction}
資料：
${JSON.stringify(limitedRows, null, 2)}

請返回 JSON 陣列，長度必須與輸入資料完全相同，每一項為字串型態之結果。`;

    const response = await ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: prompt,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            results: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
            },
          },
          required: ['results'],
        },
      },
    });

    const parsed = JSON.parse(response.text || '{"results": []}');
    res.json({ success: true, results: parsed.results });
  } catch (error: any) {
    console.error('Batch transform error:', error);
    res.status(500).json({ error: '批次 AI 處理失敗', details: error?.message });
  }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// Setup Vite middleware in dev or static files in production
async function startServer() {
  const isProd = process.env.NODE_ENV === 'production';

  if (!isProd) {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.resolve(distPath, 'index.html'));
    });
  }

  app.listen(port, '0.0.0.0', () => {
    console.log(`Excel Morph Server running on http://0.0.0.0:${port}`);
  });
}

startServer();
