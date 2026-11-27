/*
 * Excel → 照合用データ への変換
 * 列は見出し名で探すので、列の並びが変わっても読める。
 * SheetJS (XLSX) のワークブックを受け取る。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Readers = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const clean = (s) => String(s == null ? '' : s).normalize('NFKC').replace(/\s+/g, '');

  const PURCHASE_COLS = {
    date: ['日付', '仕入日', '伝票日付'],
    slip: ['伝票№', '伝票No', '伝票NO', '伝票番号'],
    maker: ['メーカコード', 'メーカーコード', 'メーカ', 'メーカー'],
    code: ['品番', '商品コード', '型番'],
    name: ['品名', '商品名'],
    qty: ['数量'],
    price: ['伝票単価', '単価', '仕入単価'],
    amount: ['金額', '仕入金額'],
    po: ['発注№', '発注No', '発注NO', '発注番号', '工事番号'],
  };
  const INVOICE_COLS = {
    date: ['月日', '日付', '納品日'],
    slip: ['伝票番号', '伝票№', '伝票No', '伝票NO'],
    text: ['商品名・内容', '商品名', '品名', '品名・規格', '内容'],
    qty: ['数量'],
    price: ['単価'],
    amount: ['金額'],
    tax: ['消費税', '税額'],
    po: ['備考', '摘要', '発注番号', '注文番号'],
  };

  function findHeader(rows, spec) {
    const required = Object.keys(spec).filter((k) => ['qty', 'amount'].includes(k));
    for (let r = 0; r < Math.min(rows.length, 30); r++) {
      const cells = (rows[r] || []).map(clean);
      const map = {};
      for (const [key, names] of Object.entries(spec)) {
        const wanted = names.map(clean);
        // 完全一致を優先、なければ前方一致
        let idx = cells.findIndex((c) => wanted.includes(c));
        if (idx < 0) idx = cells.findIndex((c) => c && wanted.some((w) => c.startsWith(w)));
        if (idx >= 0) map[key] = idx;
      }
      const hits = Object.keys(map).length;
      if (required.every((k) => k in map) && hits >= 4) return { row: r, map, hits };
    }
    return null;
  }

  function readSheet(XLSX, wb, spec, sheetHint) {
    let best = null;
    for (const name of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
      const h = findHeader(rows, spec);
      if (!h) continue;
      const bonus = sheetHint && name.includes(sheetHint) ? 5 : 0;
      if (!best || h.hits + bonus > best.h.hits + best.bonus) best = { name, rows, h, bonus };
    }
    if (!best) return null;
    const { rows, h } = best;
    const out = [];
    for (let r = h.row + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      if (!row.some((c) => c != null && c !== '')) continue;
      const rec = { id: r + 1 }; // Excel 上の行番号
      for (const [k, idx] of Object.entries(h.map)) rec[k] = row[idx];
      out.push(rec);
    }
    return { sheet: best.name, headerRow: h.row + 1, rows: out };
  }

  const toNum = (v) => {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v;
    const n = Number(String(v).normalize('NFKC').replace(/[,，\s円¥]/g, ''));
    return Number.isFinite(n) ? n : null;
  };
  // 品番などの文字列列。日付書式が付いた数値（例：5240）が Date で来ることがあるので数値に戻す
  const str = (v) => {
    if (v == null) return '';
    if (v instanceof Date) return String(Math.round(v.getTime() / 86400000 + 25569));
    return String(v).replace(/[　\s]+$/g, '').trim();
  };

  function readPurchases(XLSX, wb) {
    const s = readSheet(XLSX, wb, PURCHASE_COLS);
    if (!s) throw new Error('仕入データの見出し行（品番・数量・金額 など）が見つかりません');
    s.rows = s.rows
      .map((r) => ({
        id: r.id, date: r.date, slip: r.slip, maker: str(r.maker), code: str(r.code), name: str(r.name),
        qty: toNum(r.qty), price: toNum(r.price), amount: toNum(r.amount), po: str(r.po),
      }))
      .filter((r) => r.code || r.name || r.amount);
    return s;
  }

  function readInvoices(XLSX, wb) {
    const s = readSheet(XLSX, wb, INVOICE_COLS, '明細');
    if (!s) throw new Error('請求明細の見出し行（商品名・数量・単価・金額 など）が見つかりません');
    s.rows = s.rows
      .map((r) => ({
        id: r.id, date: r.date, slip: r.slip, text: str(r.text),
        qty: toNum(r.qty), price: toNum(r.price), amount: toNum(r.amount), tax: toNum(r.tax), po: str(r.po),
      }))
      .filter((r) => r.text || r.amount);
    // 請求書鑑（当月お買上額）があれば拾う
    s.header = {};
    for (const name of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
      for (const row of rows) {
        for (let i = 0; i < (row || []).length - 1; i++) {
          const k = clean(row[i]);
          if (/^(当月お?買上額|今回お?買上額|当月売上額)$/.test(k)) s.header.purchaseAmount = toNum(row[i + 1]);
          if (/^(消費税等?額|消費税)$/.test(k) && s.header.tax == null) s.header.tax = toNum(row[i + 1]);
          if (/^(当月ご?請求額|今回ご?請求額)$/.test(k)) s.header.billed = toNum(row[i + 1]);
        }
      }
    }
    return s;
  }

  return { readPurchases, readInvoices };
});
