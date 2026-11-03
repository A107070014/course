import { MappingRule } from '../types/excel';

function normalize(s: string): string {
  return (s || '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[\s_()（）[\]【】]/g, '')
    .trim()
    .toLowerCase();
}

/**
 * 100% 純前端靜態語意欄位配對引擎 (No Backend, No AI required)
 * 專為 GitHub Pages 靜態網站設計，支援同義詞庫、多欄位自動合併、四則運算推導與精確比對。
 */
export function generateSmartClientMapping(
  sourceHeaders: string[],
  targetHeaders: string[]
): {
  rules: MappingRule[];
  matchedCount: number;
  summary: string;
} {
  const rules: MappingRule[] = [];
  let matchedCount = 0;

  // Synonym dictionaries for standard commercial fields
  const synonymGroups: Record<string, string[]> = {
    orderId: ['訂單編號', '訂單號', '單號', '單據編號', '訂單序號', 'orderid', 'orderno', '訂單編號/id'],
    customerName: ['姓名', '客戶姓名', '客戶名稱', '收件客戶', '收件人', '收件人姓名', '聯絡人', '訂購人', '訂購人姓名', '買家', '顧客', 'name', 'customer'],
    phone: ['電話', '手機', '聯絡電話', '收件人手機', '收件人電話', '行動電話', '手機號碼', '電話號碼', '市話', 'phone', 'tel', 'mobile'],
    address: ['地址', '完整收件地址', '收件地址', '配送完整地址', '寄送地址', '收貨地址', '門牌地址', '送貨地址', 'address'],
    product: ['商品', '品名', '商品名稱', '商品品名', '商品明細', '商品明細與件數', '品項', '購買項目', 'product', 'item'],
    qty: ['數量', '件數', '總件數', '訂購數量', '購買數量', 'qty', 'quantity', 'count'],
    price: ['單價', '商品單價', '價格', '售價', 'unitprice', 'price'],
    amount: ['金額', '總金額', '代收貨款', '代收貨款金額', '應付金額', '應收金額', '付款金額', '總計', '總計金額', 'amount', 'total'],
    date: ['日期', '出貨日期', '下單時間', '訂單日期', '下單日期', '建立日期', '建立時間', '交易日期', 'date', 'orderdate'],
    status: ['狀態', '處理狀態', '訂單狀態', '配送狀態', '出貨狀態', '運送狀態', 'status'],
    notes: ['備註', '顧客備註', '買家備註', '訂單備註', '備註說明', '注意事項', '說明', 'notes', 'remarks'],
    payment: ['付款方式', '付款狀態', '支付方式', '付款類別', 'payment', 'paymethod'],
    temperature: ['溫層', '配送溫層', '溫層分類', '常溫/低溫', 'temperature'],
  };

  targetHeaders.forEach((targetCol, idx) => {
    const normTarget = normalize(targetCol);

    // 1. Direct exact match
    const exactSource = sourceHeaders.find((s) => s.trim() === targetCol.trim());
    if (exactSource) {
      rules.push({
        id: `rule-smart-${Date.now()}-${idx}`,
        targetColumn: targetCol,
        ruleType: 'direct',
        sourceColumns: [exactSource],
        defaultValue: '',
        enabled: true,
      });
      matchedCount++;
      return;
    }

    // 2. Normalized match (ignoring spaces, parentheses, underscores)
    const normSource = sourceHeaders.find((s) => normalize(s) === normTarget);
    if (normSource) {
      rules.push({
        id: `rule-smart-${Date.now()}-${idx}`,
        targetColumn: targetCol,
        ruleType: 'direct',
        sourceColumns: [normSource],
        defaultValue: '',
        enabled: true,
      });
      matchedCount++;
      return;
    }

    // 3. Multi-column combine: Address parts (寄送縣市 + 行政區域 + 街道門牌 -> 完整地址)
    const isAddressTarget = synonymGroups.address.some((term) => normTarget.includes(normalize(term)));
    if (isAddressTarget) {
      const cityCol = sourceHeaders.find((s) => normalize(s).includes('縣市') || normalize(s).includes('城市'));
      const distCol = sourceHeaders.find((s) => normalize(s).includes('行政區') || normalize(s).includes('鄉鎮市區') || normalize(s).includes('區域'));
      const streetCol = sourceHeaders.find((s) => normalize(s).includes('街道') || normalize(s).includes('門牌') || normalize(s).includes('路街') || normalize(s).includes('地址'));

      const addressParts = [cityCol, distCol, streetCol].filter(Boolean) as string[];
      if (addressParts.length >= 2) {
        rules.push({
          id: `rule-smart-${Date.now()}-${idx}`,
          targetColumn: targetCol,
          ruleType: 'combine',
          sourceColumns: addressParts,
          separator: '',
          defaultValue: '',
          enabled: true,
        });
        matchedCount++;
        return;
      }
    }

    // 4. Multi-column combine: Product + Qty (商品品名 + 數量 -> 商品明細與件數)
    if (normTarget.includes('明細') && (normTarget.includes('件數') || normTarget.includes('數量'))) {
      const prodCol = sourceHeaders.find((s) => normalize(s).includes('商品') || normalize(s).includes('品名'));
      const qtyCol = sourceHeaders.find((s) => normalize(s).includes('數量') || normalize(s).includes('件數'));
      if (prodCol && qtyCol) {
        rules.push({
          id: `rule-smart-${Date.now()}-${idx}`,
          targetColumn: targetCol,
          ruleType: 'combine',
          sourceColumns: [prodCol, qtyCol],
          separator: ' x',
          defaultValue: '',
          enabled: true,
        });
        matchedCount++;
        return;
      }
    }

    // 5. Synonym dictionary lookup
    let matchedSynonymSource: string | undefined = undefined;
    for (const [, terms] of Object.entries(synonymGroups)) {
      const targetMatchesGroup = terms.some((term) => normTarget.includes(normalize(term)) || normalize(term).includes(normTarget));
      if (targetMatchesGroup) {
        matchedSynonymSource = sourceHeaders.find((s) => {
          const normS = normalize(s);
          return terms.some((term) => normS.includes(normalize(term)) || normalize(term).includes(normS));
        });
        if (matchedSynonymSource) break;
      }
    }

    if (matchedSynonymSource) {
      rules.push({
        id: `rule-smart-${Date.now()}-${idx}`,
        targetColumn: targetCol,
        ruleType: 'direct',
        sourceColumns: [matchedSynonymSource],
        defaultValue: '',
        enabled: true,
      });
      matchedCount++;
      return;
    }

    // 6. Substring match
    const subMatch = sourceHeaders.find(
      (s) => normalize(s).includes(normTarget) || normTarget.includes(normalize(s))
    );
    if (subMatch) {
      rules.push({
        id: `rule-smart-${Date.now()}-${idx}`,
        targetColumn: targetCol,
        ruleType: 'direct',
        sourceColumns: [subMatch],
        defaultValue: '',
        enabled: true,
      });
      matchedCount++;
      return;
    }

    // 7. Math derivation: Amount = Qty * Price
    const isAmountTarget = synonymGroups.amount.some((term) => normTarget.includes(normalize(term)));
    if (isAmountTarget) {
      const qtyCol = sourceHeaders.find((s) => normalize(s).includes('數量') || normalize(s).includes('件數'));
      const priceCol = sourceHeaders.find((s) => normalize(s).includes('單價') || normalize(s).includes('價格'));
      if (qtyCol && priceCol) {
        rules.push({
          id: `rule-smart-${Date.now()}-${idx}`,
          targetColumn: targetCol,
          ruleType: 'math',
          sourceColumns: [priceCol, qtyCol],
          formula: 'A * B',
          defaultValue: '0',
          enabled: true,
        });
        matchedCount++;
        return;
      }
    }

    // 8. Default fallback: constant
    const defaultVal =
      normTarget.includes('狀態') ? '待處理' :
      normTarget.includes('溫層') ? '常溫' : '';

    rules.push({
      id: `rule-smart-${Date.now()}-${idx}`,
      targetColumn: targetCol,
      ruleType: 'constant',
      sourceColumns: [],
      defaultValue: defaultVal,
      enabled: true,
    });
  });

  const summary = `純前端本機配對完成：已成功配對 ${matchedCount} / ${targetHeaders.length} 個欄位（支援同義詞庫、多欄位地址自動合併與運算規則，100% 靜態運行無後端 API）。`;

  return { rules, matchedCount, summary };
}
