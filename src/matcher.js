/*
 * 請求書 × 仕入データ 照合ロジック
 *
 * ブラウザ（index.html）と Node（テスト・CLI）の両方から使う。
 * 入力は「行の配列」に正規化済みのデータ。Excel の読み込みは readers 側で行う。
 */
(function (root, factory) {
  const M = factory();
  M.__factory = factory; // Web Worker でも同じロジックを動かすため（index.html 参照）
  if (typeof module === 'object' && module.exports) module.exports = M;
  else root.Matcher = M;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------- 文字列正規化

  // 同じ文字列の正規化を何度も計算しないよう結果を覚えておく（大量データ対策）
  const MEMOS = [];
  function memo(fn) {
    const cache = new Map();
    MEMOS.push(cache);
    return (s) => {
      const k = s == null ? '' : String(s);
      let v = cache.get(k);
      if (v === undefined) { v = fn(k); cache.set(k, v); }
      return v;
    };
  }
  const clearMemos = () => MEMOS.forEach((c) => c.clear());

  // 全角/半角・大小文字をそろえ、区切り記号を除く
  const norm = memo(normRaw);
  function normRaw(s) {
    return String(s == null ? '' : s)
      .normalize('NFKC')
      .toUpperCase()
      .replace(/[\s\-‐‑–—−ｰ_\/・.,，、。()（）\[\]［］「」*＊＃#'"]/g, '');
  }

  // OCR や手入力で起きやすい揺れを吸収した「ゆるい」正規化
  //   I/L/| → 1, O → 0, 濁点・半濁点を除去（バイブ/パイプ）, 小書きカナ → 大きいカナ, 長音を除去
  const SMALL = { 'ァ': 'ア', 'ィ': 'イ', 'ゥ': 'ウ', 'ェ': 'エ', 'ォ': 'オ', 'ッ': 'ツ', 'ャ': 'ヤ', 'ュ': 'ユ', 'ョ': 'ヨ', 'ヮ': 'ワ' };
  // 同じ意味の語をそろえる（送料・運賃など）
  const SYNONYMS = [
    [/ソウリヨウ|ウンチン|運賃|配送料|運送料/g, '送料'],
  ];
  const loose = memo(looseRaw);
  function looseRaw(s) {
    let t = norm(s)
      .normalize('NFD').replace(/[゙゚]/g, '').normalize('NFC')
      .replace(/[ァィゥェォッャュョヮ]/g, (c) => SMALL[c])
      .replace(/ー/g, '')
      .replace(/[IL|]/g, '1')
      .replace(/O/g, '0');
    for (const [re, to] of SYNONYMS) t = t.replace(re, to);
    return t;
  }

  const bigrams = memo(bigramsRaw);
  function bigramsRaw(s) {
    const out = [];
    if (s.length === 1) out.push(s);
    for (let i = 0; i < s.length - 1; i++) out.push(s.slice(i, i + 2));
    return out;
  }

  // a の bigram がどれだけ b に含まれているか（0〜1）。a が短い品番、b が長い品名文字列を想定
  function containment(a, b) {
    if (!a || !b) return 0;
    if (b.includes(a)) return 1;
    const bg = bigrams(a);
    const pool = bigramCounts(b);
    const used = new Map();
    let hit = 0;
    for (const g of bg) {
      const n = pool.get(g) || 0, u = used.get(g) || 0;
      if (u < n) { hit++; used.set(g, u + 1); }
    }
    return hit / bg.length;
  }
  const bigramCounts = memo((b) => {
    const pool = new Map();
    for (const g of bigrams(b)) pool.set(g, (pool.get(g) || 0) + 1);
    return pool;
  });
  // 品名を空白で区切った語（ゆるい正規化済み）
  const tokensOf = memo((text) => text.normalize('NFKC').toUpperCase().split(/[\s　]+/).map(loose));

  // a の中で b に連続して現れる最長部分の長さ / a の長さ
  function longestRun(a, b) {
    if (!a || !b) return 0;
    let best = 0;
    for (let i = 0; i < a.length; i++) {
      for (let j = i + best + 1; j <= a.length; j++) {
        if (b.includes(a.slice(i, j))) best = j - i; else break;
      }
    }
    return best / a.length;
  }

  // 品番の類似度：完全包含なら 1、そうでなければ部分一致の度合い
  function codeSimilarity(code, text) {
    const c1 = norm(code), t1 = norm(text);
    if (!c1) return 0;
    // 1〜2文字の品番は、品名の中で独立した語として出てくる場合だけ一致とみなす
    if (loose(code).length <= 2) {
      return tokensOf(text).includes(loose(code)) ? 0.9 : 0;
    }
    if (t1.includes(c1)) return 1;
    const c2 = loose(code), t2 = loose(text);
    if (t2.includes(c2)) return 0.95;
    const s = 0.5 * containment(c2, t2) + 0.5 * longestRun(c2, t2);
    // 短い品番（3文字以下）は偶然一致しやすいので割り引く
    return c2.length <= 3 ? s * 0.6 : s;
  }

  function nameSimilarity(name, text) {
    const n = loose(name);
    if (n.length < 2) return 0;
    return containment(n, loose(text));
  }

  // ---------------------------------------------------------------- 発注番号

  // 備考欄などから発注番号らしきもの（英字1〜2 + 数字4〜8）を取り出す
  const extractPo = memo(extractPoRaw);
  function extractPoRaw(s) {
    const t = String(s == null ? '' : s).normalize('NFKC').toUpperCase();
    const m = t.match(/[A-Z]{1,2}\d{4,8}[A-Z]?/);
    if (m) return m[0];
    const d = t.match(/\d{4,8}/);
    return d ? d[0] : '';
  }
  const poDigits = memo((a) => a.replace(/\D/g, ''));
  function poSimilarity(a, b) {
    a = extractPo(a); b = extractPo(b);
    if (!a || !b) return 0;
    if (a === b) return 1;
    const da = poDigits(a), db = poDigits(b);
    if (da && db && (da === db || da.endsWith(db) || db.endsWith(da))) return 0.7;
    return 0;
  }

  // ---------------------------------------------------------------- 数値・日付

  function num(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v;
    const n = Number(String(v).normalize('NFKC').replace(/[,，\s円¥]/g, ''));
    return Number.isFinite(n) ? n : null;
  }
  function eq(a, b, tol) {
    if (a == null || b == null) return false;
    return Math.abs(a - b) <= (tol == null ? 0.005 : tol);
  }

  // 日付を「年初からの通算日」に変換（Date / Excel シリアル値 / "10/01" / "2026/10/01" に対応）
  function dayIndex(v, fallbackYear) {
    if (v == null || v === '') return null;
    let d;
    if (v instanceof Date) d = v;
    else if (typeof v === 'number') d = new Date(Math.round((v - 25569) * 86400000));
    else {
      const t = String(v).normalize('NFKC').trim();
      let m = t.match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
      if (m) d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
      else if ((m = t.match(/^(\d{1,2})[\/\-.月](\d{1,2})/))) d = new Date(Date.UTC(fallbackYear || 2000, +m[1] - 1, +m[2]));
      else return null;
    }
    return Math.round(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) / 86400000);
  }

  // ---------------------------------------------------------------- 照合本体

  const GENERIC = /^(送料|値引|値引キ|ネビキ|手数料|代引|割増|調整)$/;

  const DEFAULTS = {
    acceptScore: 45,   // これ未満のペアは「対応なし」とみなす
    minIdentity: 0.45, // 品番・品名の一致度がこれ未満なら、数値が合っていても採用しない
    priceTol: 0.005,   // 単価の許容誤差（円）
    amountTol: 1,      // 金額の許容誤差（円）
  };

  // 1 組（仕入 p × 請求 v）の採点
  function scorePair(p, v, opt) {
    const code = codeSimilarity(p.code, v.text);
    const name = nameSimilarity(p.name, v.text);
    // 品名だけ（メーカー名だけ等）で一致しているペアは弱く扱う
    let identity = norm(p.code) ? Math.max(code, 0.6 * code + 0.4 * name) : 0.8 * name;
    const po = poSimilarity(p.po, v.po);
    // 「送料」のような一般的な語は、発注番号か伝票番号が合わない限り同じ明細とはみなさない
    if (GENERIC.test(loose(p.code)) && po < 0.7 && String(p.slip) !== String(v.slip)) identity *= 0.4;
    const qty = eq(p.qty, v.qty, 0.0001);
    const price = eq(p.price, v.price, opt.priceTol);
    const amount = eq(p.amount, v.amount, opt.amountTol);
    let date = 0;
    if (p.day != null && v.day != null) {
      const dd = Math.abs(p.day - v.day);
      date = dd === 0 ? 1 : dd <= 3 ? 0.6 : dd <= 10 ? 0.3 : 0;
    }
    let slip = 0;
    if (p.slip != null && v.slip != null && p.slip !== '' && v.slip !== '') {
      const a = Number(p.slip), b = Number(v.slip);
      if (String(p.slip) === String(v.slip)) slip = 1;
      else if (Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) <= 2) slip = 0.5;
    }
    // 符号（返品）が逆なら別物
    const signOk = (p.qty == null || v.qty == null) || Math.sign(p.qty) === Math.sign(v.qty) || p.qty === 0 || v.qty === 0;

    let score =
      45 * identity +
      15 * po +
      8 * (qty ? 1 : 0) +
      8 * (price ? 1 : 0) +
      8 * (amount ? 1 : 0) +
      8 * date +
      8 * slip;
    if (!signOk) score -= 40;
    return { score, identity, code, name, po, qty, price, amount, date, slip };
  }

  function judge(p, v, opt) {
    const qtyOk = eq(p.qty, v.qty, 0.0001);
    const priceOk = eq(p.price, v.price, opt.priceTol);
    const amountOk = eq(p.amount, v.amount, opt.amountTol);
    const issues = [];
    // 数量・単価は違うが金額は一致 → 単位の違い（巻⇔m、箱⇔個 など）の可能性が高い
    if (amountOk && !qtyOk && !priceOk) {
      return { status: '単位違い（金額一致）', ok: true, warn: true, qtyDiff: (v.qty ?? 0) - (p.qty ?? 0), priceDiff: (v.price ?? 0) - (p.price ?? 0), amountDiff: (v.amount ?? 0) - (p.amount ?? 0) };
    }
    if (!qtyOk) issues.push('数量');
    if (!priceOk) issues.push('単価');
    if (!amountOk && qtyOk && priceOk) issues.push('金額');
    return {
      status: issues.length ? issues.join('・') + '差異' : '一致',
      ok: issues.length === 0,
      qtyDiff: (v.qty ?? 0) - (p.qty ?? 0),
      priceDiff: (v.price ?? 0) - (p.price ?? 0),
      amountDiff: (v.amount ?? 0) - (p.amount ?? 0),
    };
  }

  function confidenceLabel(s) {
    if (s.numericOnly) return '中';
    if (s.identity >= 0.9 && (s.qty || s.price)) return '高';
    if (s.identity >= 0.65) return '中';
    return '低';
  }

  // 一致率（%）：品物の一致度 50% ＋ 発注番号/伝票番号 15% ＋ 数量・単価・金額 35%
  // 日付は期ズレがあるので含めない
  function matchPercent(s) {
    if (!s) return null;
    const numFrac = ((s.qty ? 1 : 0) + (s.price ? 1 : 0) + (s.amount ? 1 : 0)) / 3;
    return Math.round(100 * (0.5 * s.identity + 0.15 * Math.max(s.po, s.slip) + 0.35 * numFrac));
  }

  // 判定マーク
  //   ◎ 自動消込：品番がほぼそのまま見つかり、数量・単価・金額も一致（確認不要）
  //   ○ ほぼ確実：数値は一致するが品番の書き方が違う／単位違い／0円（さっと目視）
  //   △ 要確認　：相手は見つかったが数値が違う、または有力な候補がある
  //   × 該当なし：近い仕入行が見つからない（手で探す）
  const MARKS = ['◎', '○', '△', '×'];
  function rateRow(r) {
    if (r.kind === 'pair') {
      if (!r.ok) return { mark: '△', reason: r.status };
      if (r.warn && /値引/.test(r.note)) return { mark: '○', reason: '値引き行と合算して一致' };
      if (r.warn) return { mark: '○', reason: '単位違い（金額は一致）' };
      if (r.confidence === '高') return { mark: '◎', reason: '品番・数量・単価・金額が一致' };
      if (r.confidence === '手動') return { mark: '○', reason: '手動で消込' };
      return { mark: '○', reason: '数値は一致／品番の書き方が違う' };
    }
    if (r.status.startsWith('取消') || r.status === '調整行') return { mark: '◎', reason: r.status };
    // 0円明細（セット部材・無償品）は支払いに影響しないので確認不要
    if (r.status === '0円明細') return { mark: '◎', reason: '0円明細（セット部材・無償品）' };
    const cands = r.cands || [];
    const free = cands.find((c) => !c.usedBy && c.pct >= 50);
    if (free) return { mark: '△', reason: '候補あり（自動では決めきれず）' };
    const used = cands.find((c) => c.usedBy && c.pct >= 50);
    if (used) return { mark: '×', reason: '同じ品物の仕入は他の請求行で使用済み（仕入計上漏れ／二重請求？）' };
    return { mark: '×', reason: '近い仕入行なし' };
  }

  // 仕入行(複数可) × 請求行(複数可) から結果行を作る。手動で消込したときにも使う
  function buildPairRow(ps, vs, s, options) {
    const opt = Object.assign({}, DEFAULTS, options || {});
    const sum = (arr, k) => arr.reduce((a, r) => a + (r[k] || 0), 0);
    const p = { qty: sum(ps, 'qty'), amount: sum(ps, 'amount'), price: ps[0].price };
    const v = { qty: sum(vs, 'qty'), amount: sum(vs, 'amount'), price: vs[0].price };
    const j = judge(p, v, opt);
    const split = ps.length > 1 || vs.length > 1;
    return {
      kind: 'pair',
      status: j.status,
      ok: j.ok,
      warn: !!j.warn,
      confidence: s ? confidenceLabel(s) : '手動',
      note: [split ? `分割計上（仕入${ps.length}行／請求${vs.length}行）` : '', s && s.numericOnly ? '品番表記が異なる（伝票・金額で照合）' : ''].filter(Boolean).join('／'),
      p: ps, v: vs,
      qtyDiff: j.qtyDiff, priceDiff: j.priceDiff, amountDiff: j.amountDiff,
      score: s ? Math.round(s.score) : null,
      pct: s ? matchPercent(s) : null,
      identity: s ? s.identity : null,
    };
  }

  // 値引き行を含むペア：金額の合計で比べる
  function buildRebateRow(pr, opt) {
    const row = buildPairRow(pr.ps, pr.vs, pr.s, opt);
    const net = pr.vs.reduce((a, v) => a + (v.amount || 0), 0);
    const pAmt = pr.ps.reduce((a, p) => a + (p.amount || 0), 0);
    if (eq(net, pAmt, opt.amountTol)) {
      Object.assign(row, { status: '一致', ok: true, warn: true, amountDiff: net - pAmt, qtyDiff: 0, priceDiff: 0 });
      row.note = '値引き・単価訂正の行と合算すると一致';
    }
    return row;
  }

  /**
   * @param {Array} purchases  [{id, date, slip, maker, code, name, qty, price, amount, po}]
   * @param {Array} invoices   [{id, date, slip, text, qty, price, amount, tax, po}]
   */
  function match(purchases, invoices, options) {
    const opt = Object.assign({}, DEFAULTS, options || {});
    const year = opt.year || new Date().getFullYear();
    clearMemos();
    const progress = typeof opt.onProgress === 'function' ? opt.onProgress : () => {};
    const P = purchases.map((r, i) => Object.assign({ _i: i }, r, { day: dayIndex(r.date, year) }));
    const V = invoices.map((r, i) => Object.assign({ _i: i }, r, { day: dayIndex(r.date, year) }));

    // 調整行（消費税調整・数量0）は照合の対象外
    const isAdjust = (v) => /消費[税稅]調整|端数調整/.test(String(v.text || '')) || (!v.qty && !v.amount);

    // ---- 候補の絞り込み（大量データ対策）
    // 全組み合わせを採点すると「請求行数 × 仕入行数」回の計算になり、全拠点分では数分かかる。
    // そこで、次のどれかを共有する組み合わせだけを採点する（共有しない組は一致率がほぼ 0 なので結果は変わらない）。
    //   ・品番/品名の 3 文字の並び　・発注番号（数字部分）　・伝票番号　・金額　・1〜2 文字の品番
    const GRAM_CAP = Math.max(300, Math.ceil(P.length * 0.03)); // ありふれた並び（多くの行に出るもの）は手がかりにしない
    const index = new Map();
    const add = (k, p) => { let a = index.get(k); if (!a) index.set(k, (a = [])); a.push(p); };
    const grams3 = (t) => { const g = new Set(); for (let i = 0; i + 3 <= t.length; i++) g.add(t.slice(i, i + 3)); return g; };
    for (const p of P) {
      const c = loose(p.code);
      if (c.length > 0 && c.length <= 2) add('s:' + c, p);
      for (const g of grams3(c)) add('g:' + g, p);
      for (const g of grams3(loose(p.name))) add('g:' + g, p);
      const po = extractPo(p.po);
      if (po) add('o:' + poDigits(po), p);
      if (p.slip != null && p.slip !== '') add('l:' + p.slip, p);
      if (p.amount) add('a:' + Math.round(p.amount), p);
    }
    const candidatesFor = (v) => {
      const set = new Set();
      const take = (k, cap) => { const a = index.get(k); if (a && (!cap || a.length <= cap)) for (const p of a) set.add(p); };
      for (const g of grams3(loose(v.text))) take('g:' + g, GRAM_CAP);
      for (const t of tokensOf(String(v.text || ''))) if (t.length > 0 && t.length <= 2) take('s:' + t);
      const po = extractPo(v.po);
      if (po) take('o:' + poDigits(po));
      if (v.slip != null && v.slip !== '') take('l:' + v.slip);
      if (v.amount) take('a:' + Math.round(v.amount), GRAM_CAP);
      return set;
    };

    // 絞り込んだ組み合わせを採点して、点の高い順に 1 対 1 で確定（貪欲法）
    const cand = [];
    const topByV = new Map(); // 請求行ごとの「近い仕入行」上位（候補表示用）
    const topByP = new Map(); // 仕入行ごとの「近い請求行」（逆方向の候補・分割計上の判定用）
    V.forEach((v, vi) => {
      if (vi % 200 === 0) progress(vi / V.length);
      if (isAdjust(v)) return;
      for (const p of candidatesFor(v)) {
        const s = scorePair(p, v, opt);
        if (s.score >= 25) {
          let t = topByV.get(v._i);
          if (!t) topByV.set(v._i, (t = []));
          t.push({ p, s });
          if (t.length > 12) { t.sort((a, b) => b.s.score - a.s.score); t.length = 6; }
          let u = topByP.get(p._i);
          if (!u) topByP.set(p._i, (u = []));
          u.push({ v, s });
          if (u.length > 24) { u.sort((a, b) => b.s.score - a.s.score); u.length = 12; }
        }
        // 品番の表記が全く違っても（社内コード等）、伝票番号か発注番号が合い、金額と数量/単価が一致すれば同じ明細とみなす
        const numericMatch = p.amount && s.amount && (((s.qty || s.price) && (s.slip === 1 || s.po >= 0.7)) || (s.slip === 1 && s.po >= 0.7));
        if ((s.score >= opt.acceptScore && s.identity >= opt.minIdentity) || numericMatch) {
          s.numericOnly = s.identity < opt.minIdentity;
          cand.push({ p, v, s });
        }
      }
    });
    progress(1);
    cand.sort((a, b) => b.s.score - a.s.score);
    const usedP = new Set(), usedV = new Set();
    const pairs = [];
    for (const c of cand) {
      if (usedP.has(c.p._i) || usedV.has(c.v._i)) continue;
      usedP.add(c.p._i); usedV.add(c.v._i);
      pairs.push({ ps: [c.p], vs: [c.v], s: c.s });
    }

    const byCodePo = new Map();
    for (const p of P) { const k = norm(p.code) + '|' + extractPo(p.po); if (!byCodePo.has(k)) byCodePo.set(k, []); byCodePo.get(k).push(p); }

    // 2 段目a：1対1で数量が合わなかったペアに、同じ品番・発注番号の残り仕入行を足すと数量が合うなら取り込む
    for (const pr of pairs) {
      const v = pr.vs[0], p0 = pr.ps[0];
      if (eq(p0.qty, v.qty, 0.0001)) continue;
      const rest = (byCodePo.get(norm(p0.code) + '|' + extractPo(p0.po)) || []).filter((p) => !usedP.has(p._i));
      if (!rest.length) continue;
      // 小さい組合せから順に試す（行数は少ないので全探索）
      let found = null;
      for (let mask = 1; mask < 1 << Math.min(rest.length, 10) && !found; mask++) {
        const pick = rest.filter((_, i) => mask & (1 << i));
        const q = p0.qty + pick.reduce((a, p) => a + (p.qty || 0), 0);
        if (eq(q, v.qty, 0.0001)) found = pick;
      }
      if (found) { found.forEach((p) => usedP.add(p._i)); pr.ps.push(...found); }
    }

    // 2 段目b：分割計上（仕入 2 行 → 請求 1 行、またはその逆）を拾う
    const freeP = P.filter((p) => !usedP.has(p._i));
    const freeV = V.filter((v) => !usedV.has(v._i) && !isAdjust(v));
    const groupKey = (p) => norm(p.code) + '|' + extractPo(p.po);
    const groups = new Map();
    for (const p of freeP) {
      const k = groupKey(p);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(p);
    }
    for (const g of groups.values()) {
      if (g.length < 2) continue;
      const sumQty = g.reduce((a, p) => a + (p.qty || 0), 0);
      const sumAmt = g.reduce((a, p) => a + (p.amount || 0), 0);
      let best = null;
      const near = new Set((topByP.get(g[0]._i) || []).map((x) => x.v));
      for (const v of freeV) {
        if (usedV.has(v._i) || !near.has(v)) continue;
        const s = scorePair(Object.assign({}, g[0], { qty: sumQty, amount: sumAmt }), v, opt);
        if (s.identity >= 0.65 && (s.qty || s.amount) && (!best || s.score > best.s.score)) best = { v, s };
      }
      if (best) {
        g.forEach((p) => usedP.add(p._i));
        usedV.add(best.v._i);
        pairs.push({ ps: g, vs: [best.v], s: best.s });
      }
    }
    // 請求側の分割（請求 2 行 → 仕入 1 行）
    for (const p of P.filter((x) => !usedP.has(x._i))) {
      const vs = (topByP.get(p._i) || []).map((x) => x.v).filter((v) => !usedV.has(v._i) && !isAdjust(v) && codeSimilarity(p.code, v.text) >= 0.9)
        .sort((a, b) => a._i - b._i);
      if (vs.length < 2) continue;
      const q = vs.reduce((a, v) => a + (v.qty || 0), 0);
      const amt = vs.reduce((a, v) => a + (v.amount || 0), 0);
      if (eq(q, p.qty, 0.0001) || eq(amt, p.amount, opt.amountTol)) {
        vs.forEach((v) => usedV.add(v._i));
        usedP.add(p._i);
        pairs.push({ ps: [p], vs, s: scorePair(p, Object.assign({}, vs[0], { qty: q, amount: amt }), opt) });
      }
    }

    const negV = V.filter((v) => !isAdjust(v) && (v.qty || 0) < 0 && (v.amount || 0) < 0);
    // 3 段目：値引き・単価訂正の行（マイナス数量）を、同じ品番・発注番号のペアに合算する
    //   例）請求 2個×31,920 ＋ 値引 −2個×1,680 ＝ 2個×30,240（仕入）
    for (const pr of pairs) {
      if (pr.ps.length !== 1) continue;
      const p = pr.ps[0];
      const base = pr.vs.reduce((a, v) => a + (v.amount || 0), 0);
      if (eq(base, p.amount, opt.amountTol)) continue;
      const adj = negV.filter((v) => !usedV.has(v._i)
        && codeSimilarity(p.code, v.text) >= 0.9 && poSimilarity(p.po, v.po) >= 0.7);
      const hit = adj.find((v) => eq(base + v.amount, p.amount, opt.amountTol));
      if (hit) {
        usedV.add(hit._i);
        pr.vs.push(hit);
        pr.rebate = true;
      }
    }

    // ---------------------------------------------------------------- 結果行を組み立てる
    const rows = [];
    for (const pr of pairs) {
      const row = pr.rebate ? buildRebateRow(pr, opt) : buildPairRow(pr.ps, pr.vs, pr.s, opt);
      rows.push(row);
    }
    // 請求のみの行のうち、同じ品名で金額が打ち消し合うもの（請求→取消）は相殺として扱う
    const cancelled = new Set();
    const restV = V.filter((v) => !usedV.has(v._i) && !isAdjust(v));
    const negByText = new Map();
    for (const x of restV) if (x.amount < 0) { const k = loose(x.text) + '|' + Math.round(-x.amount); if (!negByText.has(k)) negByText.set(k, []); negByText.get(k).push(x); }
    for (const a of restV) {
      if (cancelled.has(a._i) || !(a.amount > 0)) continue;
      const b = (negByText.get(loose(a.text) + '|' + Math.round(a.amount)) || []).find((x) => !cancelled.has(x._i));
      if (b) { cancelled.add(a._i); cancelled.add(b._i); }
    }
    for (const v of V) {
      if (usedV.has(v._i)) continue;
      if (cancelled.has(v._i)) {
        rows.push({ kind: 'invoiceOnly', status: '取消（相殺済）', ok: true, confidence: '', note: '請求書内で同額の取消行と相殺', p: [], v: [v], qtyDiff: null, priceDiff: null, amountDiff: v.amount || 0, score: null });
        continue;
      }
      if (!isAdjust(v) && !v.amount) {
        rows.push({ kind: 'invoiceOnly', status: '0円明細', ok: true, confidence: '', note: '金額0円（無償・セット品など）', p: [], v: [v], qtyDiff: null, priceDiff: null, amountDiff: 0, score: null });
        continue;
      }
      rows.push({
        kind: 'invoiceOnly',
        status: isAdjust(v) ? '調整行' : '請求のみ',
        ok: isAdjust(v),
        confidence: '', note: isAdjust(v) ? '' : '仕入データに該当なし：未計上・二重請求・取消の可能性',
        p: [], v: [v], qtyDiff: null, priceDiff: null, amountDiff: v.amount || 0, score: null,
      });
    }
    const cancelledP = new Set();
    const restP = P.filter((p) => !usedP.has(p._i));
    const negByCode = new Map();
    for (const x of restP) if (x.amount < 0) { const k = norm(x.code) + '|' + Math.round(-x.amount); if (!negByCode.has(k)) negByCode.set(k, []); negByCode.get(k).push(x); }
    for (const a of restP) {
      if (cancelledP.has(a._i) || !(a.amount > 0)) continue;
      const b = (negByCode.get(norm(a.code) + '|' + Math.round(a.amount)) || []).find((x) => !cancelledP.has(x._i));
      if (b) { cancelledP.add(a._i); cancelledP.add(b._i); }
    }
    for (const p of P) {
      if (usedP.has(p._i)) continue;
      if (cancelledP.has(p._i)) {
        rows.push({ kind: 'purchaseOnly', status: '取消（相殺済）', ok: true, confidence: '', note: '仕入データ内で同額の赤伝と相殺', p: [p], v: [], qtyDiff: null, priceDiff: null, amountDiff: -(p.amount || 0), score: null });
        continue;
      }
      rows.push({
        kind: 'purchaseOnly', status: '仕入のみ', ok: false, confidence: '',
        note: '請求書に該当なし：翌月請求・請求漏れ・社内伝票の可能性',
        p: [p], v: [], qtyDiff: null, priceDiff: null, amountDiff: -(p.amount || 0), score: null,
      });
    }

    // 請求行ごとに候補（近い仕入行の上位3件）と判定マークを付ける
    const ownerOfP = new Map();
    for (const r of rows) if (r.kind === 'pair') for (const p of r.p) ownerOfP.set(p._i, r);
    for (const r of rows) {
      if (!r.v.length) continue;
      const mine = new Set(r.p.map((p) => p._i));
      const zero = r.v.every((v) => !v.amount);
      const list = (zero ? [] : topByV.get(r.v[0]._i) || [])
        .filter((c) => !mine.has(c.p._i))
        .sort((x, y) => y.s.score - x.s.score)
        .map((c) => ({ p: c.p, pct: matchPercent(c.s), usedBy: ownerOfP.has(c.p._i) ? ownerOfP.get(c.p._i).v.map((v) => v.id) : null }))
        .filter((c) => c.pct >= 40) // 一致率 40% 未満は候補として出さない（ノイズになるため）
        .slice(0, 3);
      r.cands = list;
      Object.assign(r, rateRow(r));
      if (r.kind !== 'pair') r.pct = r.mark === '△' || r.mark === '×' ? (list[0] ? list[0].pct : 0) : null;
    }
    for (const r of rows) if (!r.v.length) { r.mark = ''; r.reason = r.status; }

    // 同じ発注番号の中で単価の振り分けが違うだけ（合計は一致）なら ○ に格上げ
    //   例）本体 98,420＋配管 3,800（請求） ⇔ 本体 102,220＋配管 0（仕入）
    const byPo = new Map();
    for (const r of rows) {
      if (r.kind !== 'pair' || r.mark !== '△') continue;
      const k = extractPo(r.p[0].po) || extractPo(r.v[0].po);
      if (!k) continue;
      if (!byPo.has(k)) byPo.set(k, []);
      byPo.get(k).push(r);
    }
    for (const g of byPo.values()) {
      if (g.length < 2) continue;
      const vs = g.reduce((a, r) => a + r.v.reduce((b, v) => b + (v.amount || 0), 0), 0);
      const ps = g.reduce((a, r) => a + r.p.reduce((b, p) => b + (p.amount || 0), 0), 0);
      if (eq(vs, ps, opt.amountTol)) {
        for (const r of g) { r.mark = '○'; r.reason = '同じ発注内で合計は一致（単価の振り分け違い）'; }
      }
    }

    // 仕入行ごとに「近い請求行」上位3件を付ける（仕入データから見る画面の候補）
    for (const p of P) {
      p._cands = (topByP.get(p._i) || []).slice().sort((a, b) => b.s.score - a.s.score)
        .map((x) => ({ v: x.v, pct: matchPercent(x.s) })).filter((x) => x.pct >= 40 && x.v.amount).slice(0, 3);
    }

    // 並び順：請求書の順（請求のない行は仕入の日付順で最後）
    rows.sort((a, b) => {
      const ka = a.v.length ? [0, a.v[0]._i] : [1, a.p[0].day ?? 0, a.p[0]._i];
      const kb = b.v.length ? [0, b.v[0]._i] : [1, b.p[0].day ?? 0, b.p[0]._i];
      for (let i = 0; i < 3; i++) if ((ka[i] ?? 0) !== (kb[i] ?? 0)) return (ka[i] ?? 0) - (kb[i] ?? 0);
      return 0;
    });

    const total = (arr, k) => arr.reduce((a, r) => a + (num(r[k]) || 0), 0);
    const summary = {
      purchaseCount: P.length,
      invoiceCount: V.length,
      purchaseTotal: total(P, 'amount'),
      invoiceTotal: total(V, 'amount'),
      matched: rows.filter((r) => r.kind === 'pair' && r.ok).length,
      mismatched: rows.filter((r) => r.kind === 'pair' && !r.ok).length,
      warned: rows.filter((r) => r.warn).length,
      invoiceOnly: rows.filter((r) => r.status === '請求のみ').length,
      purchaseOnly: rows.filter((r) => r.status === '仕入のみ').length,
      marks: Object.fromEntries(MARKS.map((m) => [m, rows.filter((r) => r.mark === m).length])),
    };
    return { rows, summary };
  }

  return { match, buildPairRow, rateRow, matchPercent, MARKS, norm, loose, codeSimilarity, nameSimilarity, extractPo, num, dayIndex, scorePair };
});
