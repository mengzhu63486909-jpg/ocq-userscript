// ==UserScript==
// @name         OCQ取数
// @namespace    ocq-qushu
// @version      0.5.0
// @description  OCQ 明细表一键取数：顾问排行+小组排名、自定义报表，导出 Excel/图片。数据只存在本机。
// @match        https://*.imocq.com/*
// @updateURL    https://raw.githubusercontent.com/mengzhu63486909-jpg/ocq-userscript/main/OCQ%E5%8F%96%E6%95%B0.user.js
// @downloadURL  https://raw.githubusercontent.com/mengzhu63486909-jpg/ocq-userscript/main/OCQ%E5%8F%96%E6%95%B0.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_setClipboard
// @grant        GM_openInTab
// @grant        GM_addValueChangeListener
// @grant        window.close
// @grant        unsafeWindow
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window.self && !window.__ocqAllowFrame) return;
  if (window.__ocqUS) return;

  const W = (typeof unsafeWindow !== 'undefined' && unsafeWindow) ? unsafeWindow : window;
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // ---------------- 存储 ----------------
  const gmGet = (k, d) => {
    try { const v = GM_getValue(k); if (v == null || v === '') return d; return typeof v === 'string' ? JSON.parse(v) : v; }
    catch (_) { return d; }
  };
  const gmSet = (k, v) => GM_setValue(k, JSON.stringify(v));
  const getSources = () => gmGet('ocq_sources', []);
  const setSources = s => gmSet('ocq_sources', s);
  const getLog = () => gmGet('ocq_log', []);
  const setLog = l => gmSet('ocq_log', l.slice(-60));
  const addLog = line => setLog(getLog().concat(line));
  const getSteps = () => gmGet('ocq_steps', []);
  const setSteps = s => gmSet('ocq_steps', s);
  const getRoster = () => gmGet('ocq_roster', []);
  const setRoster = r => gmSet('ocq_roster', r);
  const getSettings = () => Object.assign({ threshold: 80, redMode: 'progress', groupSubtotal: false, onlyRoster: false, title: '顾问业绩排行', tpl: '' }, gmGet('ocq_settings', {}));
  const setSettings = s => gmSet('ocq_settings', s);
  const getPrefs = () => gmGet('ocq_prefs', {});
  const setPrefs = p => gmSet('ocq_prefs', p);

  // ---------------- 标准字段 ----------------
  const STD = ['成交额(元)', '成单', '看位数', '总客资数量', '无效客资', '有效客资', '变动金额(元)'];
  const FIXED_COLS = ['排名', '顾问', '所属小组', '成交额(元)', '成单', '看位数', '总客资数量', '无效客资', '有效客资', '有效客资成单率', '变动金额(元)', '总成交(元)', '月目标(元)', '月完成率'];
  const COUNT_FIELDS = ['成单', '看位数', '总客资数量', '无效客资', '有效客资'];

  // ---------------- 基础工具 ----------------
  const norm = s => String(s == null ? '' : s).replace(/\s+/g, '').trim();
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ('&' + { '&': 'amp', '<': 'lt', '>': 'gt', '"': 'quot' }[c] + ';'));
  function num(s) {
    let t = String(s == null ? '' : s).trim();
    if (!t) return null;
    let neg = false;
    if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
    t = t.replace(/[−–—]/g, '-');
    if (/^-/.test(t)) neg = true;
    t = t.replace(/[,，\s元¥￥%％]/g, '').replace(/^-/, '');
    const m = t.match(/^\d+(\.\d+)?$/) || t.match(/^\d+(\.\d+)?万$/);
    if (!m) return null;
    let v = parseFloat(t);
    if (/万$/.test(t)) v *= 10000;
    return neg ? -v : v;
  }
  function isAddressLike(s) {
    const t = String(s == null ? '' : s).trim();
    if (!t) return true;
    if (/合计|总计|小计|汇总/.test(t)) return true;
    if (/^.+-\d{4,}/.test(t)) return true;
    if (/\d{5,}/.test(t) && t.length >= 8) return true;
    if (/[樓楼邨村街路巷道號号房苑園园厦廈]/.test(t) && /\d/.test(t)) return true;
    return t.length > 16;
  }
  function looksLikePersonName(s) {
    const t = String(s == null ? '' : s).trim();
    if (!t || isAddressLike(t) || /^\d+$/.test(t)) return false;
    if (/^[\u4e00-\u9fff·]{2,8}$/.test(t)) return true;
    return t.length <= 8 && !/\d{3,}/.test(t);
  }
  function visible(el) {
    if (!el || !el.getBoundingClientRect) return false;
    const r = el.getBoundingClientRect();
    if (r.width < 2 && r.height < 2) return false;
    const st = getComputedStyle(el);
    return st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
  }
  function isDisabled(el) {
    if (!el) return true;
    if (el.disabled || el.getAttribute('disabled') != null) return true;
    if (/\bdisabled\b|\bis-disabled\b|\bbtn-disabled\b/i.test(el.className || '')) return true;
    if (el.getAttribute('aria-disabled') === 'true') return true;
    return !!el.closest('[disabled], .disabled, .is-disabled, .btn-disabled');
  }
  function cellText(el) {
    if (!el) return '';
    const cell = el.querySelector && el.querySelector('.cell');
    return ((cell || el).innerText || (cell || el).textContent || '').replace(/\u00a0/g, ' ').trim();
  }

  // ---------------- 表格解析 ----------------
  function parseTable(root) {
    if (!root) return null;
    const box = root.classList && root.classList.contains('el-table') ? root
      : ((root.closest && root.closest('.el-table, .ant-table, .vxe-table')) || root);
    if (box.classList && box.classList.contains('el-table')) {
      const hw = [...box.children].find(c => c.classList && c.classList.contains('el-table__header-wrapper')) || box.querySelector('.el-table__header-wrapper');
      const bw = [...box.children].find(c => c.classList && c.classList.contains('el-table__body-wrapper'))
        || [...box.querySelectorAll('.el-table__body-wrapper')].find(w => !w.closest('.el-table__fixed, .el-table__fixed-left, .el-table__fixed-right'));
      if (!hw || !bw) return null;
      const ths = [...hw.querySelectorAll('thead th')].filter(th => !th.classList.contains('gutter'));
      const headers = ths.map((th, i) => {
        const t = norm(cellText(th));
        if (t) return t;
        if (th.classList.contains('el-table-column--selection') || th.querySelector('.el-checkbox')) return `__sel${i}`;
        return `__col${i}`;
      });
      const rows = [...bw.querySelectorAll('tbody tr')].map(tr => {
        const tds = [...tr.querySelectorAll(':scope > td')];
        return headers.map((_, i) => tds[i] ? cellText(tds[i]) : '');
      }).filter(r => r.some(c => c));
      return { headers, rows };
    }
    const table = box.tagName === 'TABLE' ? box : box.querySelector('table');
    if (!table) return null;
    const ths = [...table.querySelectorAll('thead th')];
    if (ths.length) {
      const headers = ths.map((th, i) => norm(cellText(th)) || `__col${i}`);
      const rows = [...table.querySelectorAll('tbody tr')].map(tr => {
        const tds = [...tr.children];
        return headers.map((_, i) => tds[i] ? cellText(tds[i]) : '');
      }).filter(r => r.some(c => c));
      return { headers, rows };
    }
    const all = [...table.querySelectorAll('tr')].map(tr => [...tr.children].map(cellText));
    if (all.length < 2 || !table.querySelector('th')) return null;
    return { headers: all[0].map((h, i) => norm(h) || `__col${i}`), rows: all.slice(1) };
  }
  function findBestTable() {
    const cands = [...document.querySelectorAll('.el-table, .ant-table, .vxe-table')]
      .concat([...document.querySelectorAll('table')].filter(t => !t.closest('.el-table, .ant-table, .vxe-table')));
    let best = null, score = -1;
    for (const t of cands) {
      if (/el-table__fixed/.test(t.className || '')) continue;
      if (!visible(t)) continue;
      const p = parseTable(t);
      if (!p || !p.rows.length) continue;
      const real = p.headers.filter(h => !/^__/.test(h)).length;
      if (real < 2) continue;
      let s = p.rows.length * 10 + real;
      if (p.headers.some(h => /沟通负责人|顾问|业务员|负责人/.test(h))) s += 200;
      if (p.headers.some(h => /金额/.test(h))) s += 100;
      if (s > score) { score = s; best = { el: t, parsed: p }; }
    }
    return best;
  }

  // ---------------- 翻页 ----------------
  function findByExactText(texts, clickableOnly) {
    const want = (Array.isArray(texts) ? texts : [texts]).map(String);
    const hits = [];
    for (const el of document.querySelectorAll('a, button, span, li, div, em, i, label')) {
      const t = (el.childElementCount === 0 ? el.textContent : el.innerText || '').replace(/\s+/g, ' ').trim();
      if (!want.includes(t) || !visible(el)) continue;
      if (clickableOnly && isDisabled(el)) continue;
      hits.push(el);
    }
    const score = el => (/^(BUTTON|A|LI)$/.test(el.tagName) ? 0 : 1) + (el.closest('.el-pagination, .layui-laypage, .pagination, [class*=page], [class*=Page]') ? -1 : 0);
    hits.sort((a, b) => score(a) - score(b));
    return hits[0] || null;
  }
  function parseTotal() {
    const chunks = [];
    for (const el of document.querySelectorAll('body *')) {
      if (el.children.length > 8) continue;
      const t = (el.innerText || '').trim();
      if (t.length < 5 || t.length > 400 || !/记录数/.test(t)) continue;
      if (visible(el)) chunks.push(t);
    }
    chunks.push((document.body.innerText || '').slice(0, 15000));
    for (const t of chunks) {
      let m = t.match(/记录数\s*[:：]\s*([0-9,]+)/);
      if (m) return parseInt(m[1].replace(/,/g, ''), 10);
      m = t.match(/共\s*([0-9,]+)\s*条/);
      if (m) return parseInt(m[1].replace(/,/g, ''), 10);
    }
    return null;
  }
  function footerText() {
    const bits = [];
    for (const el of document.querySelectorAll('body *')) {
      if (el.children.length > 6) continue;
      const t = (el.innerText || '').trim();
      if (/记录数/.test(t) && t.length < 600 && visible(el)) bits.push(t);
    }
    bits.push((document.body.innerText || '').slice(0, 15000));
    return bits.join('\n');
  }
  function isLoading() {
    return [...document.querySelectorAll('.el-loading-mask, .ant-spin-spinning, .el-loading-spinner, .layui-layer-loading')].some(visible);
  }
  async function waitSettle(maxMs) {
    const t0 = Date.now();
    while (Date.now() - t0 < (maxMs || 12000)) {
      if (!isLoading() && findBestTable()) return true;
      await sleep(300);
    }
    return !!findBestTable();
  }
  function fingerprint() {
    const t = findBestTable();
    if (!t || !t.parsed.rows.length) return '';
    return t.parsed.rows.slice(0, 3).map(r => r.join('|')).join(';;') + '#' + t.parsed.rows.length;
  }
  async function waitChange(prev, maxMs) {
    const t0 = Date.now();
    while (Date.now() - t0 < (maxMs || 10000)) {
      await sleep(300);
      if (isLoading()) continue;
      const now = fingerprint();
      if (now && now !== prev) return true;
    }
    return false;
  }
  function findPageSizeSelect() {
    const sels = [...document.querySelectorAll('select')].filter(visible);
    for (const sel of sels) {
      let p = sel.parentElement;
      for (let i = 0; i < 5 && p; i++, p = p.parentElement) if (/每页显示|条\/页|每页/.test(p.innerText || '')) return sel;
    }
    for (const sel of sels) {
      const ns = [...sel.options].map(o => parseInt(String(o.value || o.text).replace(/\D/g, ''), 10)).filter(Boolean);
      if (ns.includes(20) && ns.some(n => n >= 50)) return sel;
    }
    return null;
  }
  async function setPageSize(total) {
    const prefer = [];
    if (total) for (const n of [50, 100, 200, 500]) if (n >= total) prefer.push(n);
    prefer.push(500, 200, 100, 50);
    const before = fingerprint();
    const sel = findPageSizeSelect();
    if (sel) {
      const opts = [...sel.options].map(o => ({ o, n: parseInt(String(o.value || o.text).replace(/\D/g, ''), 10) })).filter(x => x.n);
      let pick = null;
      for (const w of prefer) { pick = opts.find(x => x.n === w); if (pick) break; }
      if (!pick && opts.length) pick = opts.reduce((a, b) => (a.n >= b.n ? a : b));
      if (pick && String(sel.value) !== String(pick.o.value)) {
        sel.value = pick.o.value;
        try { if (W.jQuery) W.jQuery(sel).val(pick.o.value).trigger('change'); } catch (_) {}
        sel.dispatchEvent(new Event('input', { bubbles: true }));
        sel.dispatchEvent(new Event('change', { bubbles: true }));
        await waitChange(before, 6000);
        await waitSettle();
        return pick.n;
      }
      return pick ? pick.n : null;
    }
    // Element-UI / 自定义下拉
    const trig = document.querySelector('.el-pagination__sizes .el-input, .el-pagination__sizes .el-select');
    if (trig) {
      trig.click();
      await sleep(350);
      const items = [...document.querySelectorAll('.el-select-dropdown__item')].filter(visible);
      for (const w of prefer) {
        const hit = items.find(li => parseInt(li.innerText.replace(/\D/g, ''), 10) === w);
        if (hit) { hit.click(); await waitChange(before, 6000); await waitSettle(); return w; }
      }
      document.body.click();
    }
    return null;
  }
  function currentPageNum() {
    const act = document.querySelector('.layui-laypage-curr, .el-pager li.active, .el-pager li.is-active, .pagination .active, [class*=page] .active, [class*=page] .on, [class*=page] .current');
    if (act) { const n = parseInt((act.innerText || '').replace(/\D/g, ''), 10); if (n) return n; }
    return null;
  }
  function clickNextPage(cur) {
    const next = findByExactText(['下一页', '下页', '›', '»'], true);
    if (next) { next.click(); return true; }
    if (cur) {
      const btn = findByExactText([String(cur + 1)], true);
      if (btn) {
        const ctx = (btn.parentElement && btn.parentElement.parentElement && btn.parentElement.parentElement.innerText) || '';
        if (btn.closest('[class*=page], [class*=Page], .el-pagination') || /上一页|下一页|跳转|每页/.test(ctx)) { btn.click(); return true; }
      }
    }
    const jump = findByExactText(['跳转', 'GO', 'Go'], true);
    if (jump && cur) {
      let box = jump.parentElement, input = null;
      for (let i = 0; i < 3 && box && !input; i++, box = box.parentElement) input = box.querySelector('input:not([type=checkbox])');
      if (input) {
        input.value = String(cur + 1);
        try { if (W.jQuery) W.jQuery(input).val(String(cur + 1)).trigger('input').trigger('change'); } catch (_) {}
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        jump.click();
        return true;
      }
    }
    return false;
  }
  async function scrapeAll(onProgress, nameIdx) {
    await waitSettle();
    let total = parseTotal();
    onProgress && onProgress(0, 0, total, '调整每页条数…');
    await setPageSize(total);
    total = parseTotal() || total;
    const first = findBestTable();
    if (!first) return { ok: false, error: '没找到数据表' };
    const headers = first.parsed.headers.slice();
    const rows = [], seen = new Set();
    let pages = 0, stagnant = 0, cur = currentPageNum() || 1;
    const dataCount = () => rows.filter(r => { const n = (r[nameIdx] || '').trim(); return n && !/合计|总计|小计|汇总/.test(n); }).length;
    while (pages < 100) {
      await waitSettle(8000);
      const t = findBestTable();
      if (!t) break;
      const fp = fingerprint();
      let added = 0;
      for (const r of t.parsed.rows) {
        const row = headers.map((_, i) => r[i] != null ? r[i] : '');
        const key = row.join('\t');
        if (seen.has(key)) continue;
        seen.add(key); rows.push(row); added++;
      }
      pages++;
      onProgress && onProgress(pages, dataCount(), total);
      if (total != null && dataCount() >= total) break;
      if (!added) { if (++stagnant >= 2) break; } else stagnant = 0;
      cur = currentPageNum() || cur;
      if (!clickNextPage(cur)) break;
      const changed = await waitChange(fp, 10000);
      cur = currentPageNum() || cur + 1;
      if (!changed && ++stagnant >= 2) break;
    }
    const dc = dataCount();
    return { ok: true, headers, rows, pages, total, collected: dc, incomplete: total != null && dc < total };
  }

  // ---------------- 页面/列分析 ----------------
  function pageKey() {
    let a = '';
    try { a = new URL(location.href).searchParams.get('analysisName') || ''; } catch (_) {}
    if (a) return a;
    const tab = document.querySelector('.el-tabs__item.is-active, .ant-tabs-tab-active, .layui-this');
    const t = tab && norm(tab.innerText);
    return t || norm(document.title) || location.pathname;
  }
  function pageCountField(key) {
    if (/无效/.test(key)) return '无效客资';
    if (/有效客资|有效线索/.test(key)) return '有效客资';
    if (/看位|邀约/.test(key)) return '看位数';
    if (/客资|线索/.test(key)) return '总客资数量';
    return null;
  }
  function analyze(parsed) {
    const { headers, rows } = parsed;
    const occ = {};
    const cols = headers.map((h, i) => {
      occ[h] = (occ[h] || 0) + 1;
      const vals = rows.map(r => (r[i] || '').trim()).filter(v => v && !/合计|总计/.test(v));
      const nums = vals.filter(v => num(v) != null).length;
      return {
        idx: i, header: h, key: h + '#' + occ[h],
        display: /^__sel/.test(h) ? '（勾选框）' : /^__col/.test(h) ? `（无表头第${i + 1}列）` : h,
        filled: vals.length, numeric: vals.length > 0 && nums / vals.length >= 0.8,
        names: vals.filter(looksLikePersonName).length, placeholder: /^__/.test(h)
      };
    });
    let nameIdx = -1;
    for (const n of ['沟通负责人', '顾问', '业务员', '负责人']) {
      const c = cols.filter(c => c.header === n);
      if (c.length) { nameIdx = c.reduce((a, b) => (b.names >= a.names ? b : a)).idx; break; }
    }
    if (nameIdx < 0) {
      const c = cols.filter(c => !c.placeholder && /负责人|顾问|业务员|销售/.test(c.header) && !/地址|项目|客户|电话|渠道/.test(c.header));
      if (c.length) nameIdx = c.reduce((a, b) => (b.names >= a.names ? b : a)).idx;
    }
    if (nameIdx < 0) {
      const c = cols.filter(c => !c.placeholder && !c.numeric && c.names > 0);
      if (c.length) nameIdx = c.reduce((a, b) => (b.names > a.names ? b : a)).idx;
    }
    const key = pageKey() + ' ' + document.title;
    const isBD = /变动/.test(key);
    const best = name => {
      const c = cols.filter(c => c.header === name && c.numeric);
      return c.length ? c.reduce((a, b) => (b.filled >= a.filled ? b : a)) : null;
    };
    const defaults = {};
    for (const c of cols) defaults[c.key] = { on: false, mode: c.numeric ? 'sum' : 'count', target: guessTarget(c), custom: c.display };
    const cj = best('成交金额'), bd = best('变动金额');
    const nc = nameIdx >= 0 ? cols[nameIdx] : null;
    if (cj && !(isBD && bd)) {
      defaults[cj.key] = { on: true, mode: 'sum', target: '成交额(元)', custom: '' };
      if (nc) defaults[nc.key] = { on: true, mode: 'count', target: '成单', custom: '' };
    }
    if (bd) defaults[bd.key] = { on: true, mode: 'sum', target: '变动金额(元)', custom: '' };
    if (!cj && !bd && nc) {
      const f = pageCountField(key);
      if (f) defaults[nc.key] = { on: true, mode: 'count', target: f, custom: '' };
      else defaults[nc.key] = { on: true, mode: 'count', target: '成单', custom: '' };
    }
    return { cols, nameIdx, defaults };
  }
  function guessTarget(c) {
    if (c.header === '成交金额') return '成交额(元)';
    if (c.header === '变动金额') return '变动金额(元)';
    const h = c.header.replace(/\(元\)|（元）/, '');
    const hit = STD.find(s => s.replace('(元)', '') === h);
    return hit || '__custom';
  }
  const labelOf = ch => ch.target === '__custom' ? (ch.custom || '').trim() : ch.target;

  // ---------------- 汇总 / 来源 ----------------
  function aggregateCol(rows, nameIdx, colIdx, mode) {
    const out = {};
    let used = 0;
    for (const r of rows) {
      const name = (r[nameIdx] || '').trim();
      if (!name || /合计|总计|小计|汇总/.test(name) || !looksLikePersonName(name)) continue;
      const raw = (r[colIdx] || '').trim();
      if (mode === 'count') {
        if (colIdx !== nameIdx && !raw) continue;
        out[name] = (out[name] || 0) + 1; used++;
      } else {
        const v = num(raw);
        if (v == null) { out[name] = out[name] || 0; continue; }
        out[name] = (out[name] || 0) + v; used++;
      }
    }
    return { data: out, used };
  }
  function parseFooterMetric(label, text) {
    const lab = String(label).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const m = String(text || '').match(new RegExp(lab + '\\s*[:：]\\s*([\\-(−–—]?[0-9,]+(?:\\.[0-9]+)?\\)?)'));
    return m ? num(m[1]) : null;
  }
  const fmt2 = v => Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  function footerCheck(data, header, text) {
    if (/^__/.test(header)) return null;
    const footer = parseFooterMetric(header, text);
    if (footer == null) return null;
    const scraped = Object.values(data).reduce((a, b) => a + b, 0);
    const ok = Math.abs(scraped - footer) < 0.015;
    return { ok, scraped, footer, text: `${header} 合计核对：抓到 ${fmt2(scraped)}，页面合计 ${fmt2(footer)}` };
  }
  // ---------------- 名单粘贴 ----------------
  function parsePaste(text) {
    const split = l => (l.includes('\t') ? l.split('\t') : l.split(',')).map(s => s.trim());
    const lines = String(text || '').split(/\r?\n/).map(l => l.replace(/\s+$/, '')).filter(l => l.trim());
    if (!lines.length) return { rows: [], error: '粘贴区是空的' };
    const first = split(lines[0]);
    const isHeader = /顾问|姓名|月目标|所属小组|排名/.test(first.join(''));
    const find = preds => first.findIndex(h => preds.some(p => norm(h) === p || norm(h).includes(p)));
    let ni = 0, gi = 1, ti = 2, data = lines;
    if (isHeader) {
      ni = find(['顾问', '姓名']); gi = find(['所属小组', '小组', '团队', '组别']); ti = find(['月目标']);
      if (ni < 0) return { rows: [], error: '表头里找不到「顾问」列' };
      if (ti < 0) return { rows: [], error: '表头里找不到「月目标」列' };
      data = lines.slice(1);
    } else if (lines.some(l => split(l).length > 3)) {
      return { rows: [], error: '多列数据但没有表头：请连表头一起粘贴，或只粘 顾问/小组/月目标 三列' };
    }
    const rows = [];
    for (const l of data) {
      const c = split(l);
      const name = (c[ni] || '').trim();
      if (!name || /合计|总计|小计|汇总|平均/.test(name) || name === '顾问') continue;
      const tv = num(c[ti]);
      rows.push({ name, group: gi >= 0 ? (c[gi] || '').trim() : '', target: tv == null ? '' : String(tv) });
    }
    return { rows, header: isHeader ? { name: first[ni], group: gi >= 0 ? first[gi] : '（无）', target: first[ti] } : null };
  }

  const sourceId = (pk, colKey, mode) => `${pk}|${colKey}|${mode}`;
  const sumOf = d => Object.values(d || {}).reduce((a, b) => a + b, 0);

  /** 把一次抓取结果写成来源（同 id 替换，不累加） */
  function upsertSources(newOnes) {
    const all = getSources();
    const msgs = [];
    for (const s of newOnes) {
      const i = all.findIndex(x => x.id === s.id);
      if (i >= 0) {
        const old = all[i];
        const same = old.rows === s.rows && old.total === s.total && Math.abs(sumOf(old.data) - sumOf(s.data)) < 0.005;
        msgs.push(same ? `「${s.label}」数据无变化，已覆盖` : `「${s.label}」已更新（原 ${old.rows} 条 → ${s.rows} 条）`);
        all[i] = s;
      } else all.push(s);
    }
    setSources(all);
    return msgs;
  }

  /** 用表 + 映射算出来源 */
  function buildSources(res, nameIdx, picked, pk, url, stepId) {
    const ft = footerText();
    const out = [], lines = [];
    for (const p of picked) {
      const { data, used } = aggregateCol(res.rows, nameIdx, p.idx, p.mode);
      const src = {
        id: sourceId(pk, p.key, p.mode), pageKey: pk, label: p.label, colHeader: p.header, colKey: p.key,
        mode: p.mode, data, rows: used, total: res.total, time: Date.now(), url, stepId: stepId || null
      };
      let line = `✅ ${p.label} ← ${p.display || p.header}（${p.mode === 'sum' ? '求和' : '计数'} ${used} 条）`;
      if (p.mode === 'sum') {
        const fc = footerCheck(data, p.header, ft);
        if (fc) { line += `\n   ${fc.ok ? '✅' : '⚠'} ${fc.text}`; src.footer = fc; }
      }
      out.push(src); lines.push(line);
    }
    return { sources: out, lines };
  }

  function namesLookWrong(rows, nameIdx) {
    const names = new Set(rows.map(r => (r[nameIdx] || '').trim()).filter(n => n && !/合计|总计/.test(n)));
    const bad = [...names].filter(n => !looksLikePersonName(n)).length;
    return !names.size || bad / names.size >= 0.5;
  }

  // ---------------- 原始明细（给自定义报表用） ----------------
  const RAW_MAX_ROWS = 5000;
  const getRaw = () => gmGet('ocq_raw', {});
  const setRaw = r => gmSet('ocq_raw', r);
  function cleanRaw(headers, rows, nameIdx) {
    // 同名表头取填得最多的一列；去掉勾选框/无表头列、全空列、合计行
    const pick = {};
    headers.forEach((h, i) => {
      if (/^__/.test(h) || !norm(h)) return;
      const filled = rows.filter(r => String(r[i] || '').trim()).length;
      if (!filled) return;
      if (!pick[h] || filled > pick[h].filled) pick[h] = { i, filled };
    });
    if (nameIdx >= 0 && pick[headers[nameIdx]]) pick[headers[nameIdx]].i = nameIdx;
    const hs = Object.keys(pick).sort((a, b) => pick[a].i - pick[b].i);
    const idx = hs.map(h => pick[h].i);
    const isTotal = r => r.slice(0, 8).some(v => /^(合计|总计|小计|汇总)$/.test(norm(v)));
    const out = [];
    for (const r of rows) {
      if (isTotal(r)) continue;
      const row = idx.map(i => String(r[i] == null ? '' : r[i]).trim());
      if (row.every(v => !v)) continue;
      out.push(row);
      if (out.length >= RAW_MAX_ROWS) break;
    }
    return { headers: hs, rows: out };
  }
  function storeRaw(pk, name, res, nameIdx, range) {
    const c = cleanRaw(res.headers, res.rows, nameIdx);
    const raw = getRaw();
    raw[pk] = { pageKey: pk, name: name || pk, nameHeader: nameIdx >= 0 ? res.headers[nameIdx] : '', headers: c.headers, rows: c.rows, time: Date.now(), range: range || '' };
    const keys = Object.keys(raw).sort((a, b) => raw[a].time - raw[b].time);
    while (keys.length > 1 && JSON.stringify(raw).length > 4e6) delete raw[keys.shift()];
    setRaw(raw);
    return c.rows.length;
  }

  // ---------------- 自定义报表 ----------------
  const getTpls = () => gmGet('ocq_tpls', []);
  const setTpls = t => gmSet('ocq_tpls', t);
  const AGG_NAMES = { rows: '条数', sum: '求和', count: '计数', avg: '平均', max: '最大', min: '最小', distinct: '去重计数' };
  const FMT_NAMES = { money: '金额', int: '整数', pct: '百分比' };
  const defFmt = agg => (agg === 'rows' || agg === 'count' || agg === 'distinct' ? 'int' : 'money');
  /** 安全公式：只认 列显示名、数字、+ - * / ( )，不用 eval */
  function compileFormula(src, names) {
    const sorted = names.filter(Boolean).slice().sort((a, b) => b.length - a.length);
    const s = String(src || '').replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')').replace(/[×xX]/g, '*').replace(/÷/g, '/');
    const toks = [];
    let i = 0;
    while (i < s.length) {
      const nm = sorted.find(n => s.startsWith(n, i));
      if (nm) { toks.push({ t: 'id', v: nm }); i += nm.length; continue; }
      if ('+-*/()'.includes(s[i])) { toks.push({ t: s[i] }); i++; continue; }
      const m = s.slice(i).match(/^\d+(\.\d+)?%?/);
      if (m) { toks.push({ t: 'n', v: parseFloat(m[0]) / (m[0].endsWith('%') ? 100 : 1) }); i += m[0].length; continue; }
      throw new Error(`「${s.slice(i, i + 6)}」认不出来，只能用列的显示名、数字和 + - * / ( )`);
    }
    if (!toks.length) throw new Error('公式是空的');
    let p = 0;
    const peek = () => toks[p] && toks[p].t;
    function expr() {
      let v = term();
      while (peek() === '+' || peek() === '-') { const op = toks[p++].t, r = term(), l = v; v = op === '+' ? o => l(o) + r(o) : o => l(o) - r(o); }
      return v;
    }
    function term() {
      let v = fac();
      while (peek() === '*' || peek() === '/') { const op = toks[p++].t, r = fac(), l = v; v = op === '*' ? o => l(o) * r(o) : o => { const d = r(o); return d ? l(o) / d : 0; }; }
      return v;
    }
    function fac() {
      const k = toks[p++];
      if (!k) throw new Error('公式不完整');
      if (k.t === '-') { const f = fac(); return o => -f(o); }
      if (k.t === 'n') return () => k.v;
      if (k.t === 'id') return o => (typeof o[k.v] === 'number' && isFinite(o[k.v]) ? o[k.v] : 0);
      if (k.t === '(') { const v = expr(); if (peek() !== ')') throw new Error('括号没配对'); p++; return v; }
      throw new Error('公式写法不对');
    }
    const f = expr();
    if (p < toks.length) throw new Error('公式写法不对');
    return f;
  }
  // ---------------- 日期 / 行筛选 ----------------
  function parseDate(s) {
    const m = String(s == null ? '' : s).match(/(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})/);
    if (!m) return '';
    const y = +m[1], mo = +m[2], d = +m[3];
    if (y < 2000 || y > 2100 || mo < 1 || mo > 12 || d < 1 || d > 31) return '';
    return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }
  const ymOf = s => { const d = parseDate(s); return d ? d.slice(0, 7) : ''; };
  const DATE_PRI = [/成交|签约|下单|订单/, /创建|录入|分配|登记|进线|获客/, /日期/, /时间/];
  function detectDateCols(headers, rows) {
    const out = [];
    headers.forEach((h, i) => {
      if (/^月份|^日期$/.test(h)) return;
      const vals = rows.map(r => r[i]).filter(v => String(v || '').trim()).slice(0, 300);
      if (vals.length && vals.filter(v => parseDate(v)).length / vals.length >= 0.8) out.push(i);
    });
    return out;
  }
  function primaryDate(headers, idxs) {
    for (const re of DATE_PRI) { const i = idxs.find(i => re.test(headers[i])); if (i != null) return i; }
    return idxs.length ? idxs[0] : -1;
  }
  const lc = s => String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
  const isMonthCol = h => /^月份/.test(h || '');
  /** 条件：包含/等于/不包含（不分大小写）、日期在…之间；表里没有这一列的条件不作用在这张表上 */
  function rowPass(t, r, conds) {
    for (const f of conds || []) {
      if (!f || !f.col) continue;
      const i = t.headers.indexOf(f.col);
      if (i < 0) continue;
      const v = r[i];
      if (f.op === 'between') {
        const d = parseDate(v) || (/^\d{4}-\d{2}$/.test(v) ? v + '-01' : '');
        if (!d) return false;
        if (f.from && d < f.from) return false;
        if (f.to && d > f.to) return false;
        continue;
      }
      const a = lc(v), b = lc(f.value);
      if (!b) continue;
      if (f.op === 'eq' ? a !== b : f.op === 'not' ? a.includes(b) : !a.includes(b)) return false;
    }
    return true;
  }
  const activeConds = cs => (cs || []).filter(f => f && f.col && (f.op === 'between' ? (f.from || f.to) : String(f.value || '').trim()));

  /** 已抓明细表（附加虚拟列：顾问、所属小组、日期、月份、月份(某日期列)） */
  function rawTables(roster) {
    const ro = new Map((roster || []).map(r => [norm(r.name), r]));
    return Object.values(getRaw()).map(t => {
      let headers = t.headers.slice(), rows = t.rows;
      const base = t.headers;
      const ni = t.nameHeader ? headers.indexOf(t.nameHeader) : -1;
      const add = [];
      if (ni >= 0 && !headers.includes('顾问')) add.push(['顾问', r => r[ni]]);
      if (ni >= 0 && !headers.includes('所属小组')) add.push(['所属小组', r => { const x = ro.get(norm(r[ni])); return (x && x.group) || ''; }]);
      const dc = detectDateCols(base, t.rows);
      for (const d of dc) add.push([`月份(${base[d]})`, r => ymOf(r[d])]);
      const pd = primaryDate(base, dc);
      if (pd >= 0) { add.push(['日期', r => parseDate(r[pd])]); add.push(['月份', r => ymOf(r[pd])]); }
      if (add.length) { headers = headers.concat(add.map(a => a[0])); rows = rows.map(r => r.concat(add.map(a => a[1](r)))); }
      return { name: t.name, pageKey: t.pageKey, headers, rows, time: t.time, dateCol: pd >= 0 ? base[pd] : '', dateCols: dc.map(i => base[i]).concat(pd >= 0 ? ['日期'] : []), range: t.range || '' };
    });
  }
  function knownHeaders() {
    const set = new Set();
    for (const t of Object.values(getRaw())) t.headers.forEach(h => set.add(h));
    for (const s of getSteps()) (s.cols || []).forEach(c => { if (!/^__/.test(c.header)) set.add(c.header); });
    set.add('顾问'); set.add('所属小组');
    return [...set];
  }
  function computeTemplate(tpl, tables) {
    const gb = (tpl.groupBy || []).filter(Boolean);
    if (!gb.length) throw new Error('这个报表还没选分组字段');
    const metrics = (tpl.metrics || []).filter(m => m.src && m.label);
    const calcs = (tpl.calcs || []).filter(c => c.label && (c.formula || c.share || c.delta));
    const labels = metrics.map(m => m.label).concat(calcs.map(c => c.label));
    const deltas = calcs.filter(c => c.delta);
    const warn = [];
    const conds = activeConds(tpl.filters);
    const missCond = new Set();
    const plan = metrics.map(m => {
      const ts = tables.filter(t => (m.src === '__rows' || t.headers.includes(m.src)) && gb.every(g => t.headers.includes(g)) && (!m.table || t.pageKey === m.table || !tables.some(x => x.pageKey === m.table)));
      if (!ts.length) warn.push(`「${m.label}」：还没有同时带「${m.src}」和「${gb.join('、')}」的已抓表`);
      return {
        m, ts: ts.map(t => {
          conds.filter(f => !t.headers.includes(f.col)).forEach(f => missCond.add(`「${t.name}」没有「${f.col}」列，这个条件没用在它上面`));
          const rows = t.rows.filter(r => rowPass(t, r, conds) && rowPass(t, r, activeConds(m.where)));
          return { t, rows, si: m.src === '__rows' ? -1 : t.headers.indexOf(m.src), gi: gb.map(g => t.headers.indexOf(g)) };
        })
      };
    });
    missCond.forEach(w => warn.push(w));
    const formulas = calcs.filter(c => !c.delta).map(c => {
      if (c.share) return { c, f: (o, tot) => (tot && tot[c.share] ? o[c.share] / tot[c.share] : 0) };
      try { return { c, f: compileFormula(c.formula, labels.filter(l => l !== c.label)) }; } catch (e) { warn.push(`计算列「${c.label}」：${e.message}`); return { c, f: () => 0 }; }
    });
    let remap = null;
    const ALL = gb.map((_, i) => i);
    const empty = j => (gb[j] === '所属小组' ? '未分组' : '（空）');
    const build = idxs => {
      const groups = new Map();
      for (const { m, ts } of plan) for (const { rows, si, gi } of ts) for (const r of rows) {
        const kv = idxs.map(j => r[gi[j]] || empty(j));
        if (remap && idxs[0] === 0) kv[0] = remap(kv[0]);
        const k = kv.join('\u0001');
        let g = groups.get(k);
        if (!g) { g = { kv, idxs, vals: {} }; groups.set(k, g); }
        (g.vals[m.label] = g.vals[m.label] || []).push(si < 0 ? 1 : r[si]);
      }
      return [...groups.values()];
    };
    const agg = (m, list) => {
      list = list || [];
      if (m.agg === 'rows') return list.length;
      if (m.agg === 'count') return list.filter(v => String(v).trim()).length;
      if (m.agg === 'distinct') return new Set(list.map(v => String(v).trim()).filter(Boolean)).size;
      const ns = list.map(num).filter(v => v != null);
      if (m.agg === 'sum') return ns.reduce((a, b) => a + b, 0);
      if (!ns.length) return 0;
      if (m.agg === 'avg') return ns.reduce((a, b) => a + b, 0) / ns.length;
      if (m.agg === 'max') return ns.reduce((a, b) => Math.max(a, b));
      if (m.agg === 'min') return ns.reduce((a, b) => Math.min(a, b));
      return 0;
    };
    let TOT = null;
    const finish = g => {
      const o = {};
      gb.forEach(h => { o[h] = ''; });
      g.idxs.forEach((j, n) => { o[gb[j]] = g.kv[n]; });
      for (const m of metrics) o[m.label] = agg(m, g.vals[m.label]);
      for (const { c, f } of formulas) o[c.label] = f(o, TOT);
      for (const c of deltas) o[c.label] = '';
      return o;
    };
    { const t0 = build([])[0]; if (t0) { TOT = finish(t0); for (const c of calcs) if (c.share) TOT[c.label] = TOT[c.share] ? 1 : 0; } }
    const sc = tpl.sort && tpl.sort.col, dir = tpl.sort && tpl.sort.dir === 'asc' ? 1 : -1;
    const cmpBy = (col, d) => (a, b) => {
      if (!col) return 0;
      const x = a[col], y = b[col];
      if (typeof x === 'number' && typeof y === 'number') return (x - y) * d;
      return String(x == null ? '' : x).localeCompare(String(y == null ? '' : y), 'zh') * d;
    };
    const cmp = cmpBy(sc, dir);
    let rows = build(ALL).map(finish);
    if (tpl.mergeSmall && gb.length === 1 && rows.length > 1) {
      const key = sc && typeof rows[0][sc] === 'number' ? sc : (metrics[0] && metrics[0].label);
      const tot = TOT ? TOT[key] : 0;
      const ranked = rows.slice().sort((a, b) => b[key] - a[key]);
      const keep = new Set(ranked.filter((r, i) => i < 15 && (!tot || r[key] / tot >= 0.01)).map(r => r[gb[0]]));
      if (keep.size < rows.length) { remap = v => (keep.has(v) ? v : '其他（合并）'); rows = build(ALL).map(finish); }
    }
    // 环比浮动：同一组内按月份排，本月 − 上月
    const mi = gb.findIndex(isMonthCol);
    if (deltas.length && mi >= 0) {
      const by = new Map();
      for (const r of rows) { const k = gb.filter((_, i) => i !== mi).map(h => r[h]).join('\u0001'); if (!by.has(k)) by.set(k, []); by.get(k).push(r); }
      for (const list of by.values()) {
        list.sort((a, b) => String(a[gb[mi]]).localeCompare(String(b[gb[mi]])));
        list.forEach((r, i) => { for (const c of deltas) r[c.label] = i && typeof r[c.delta] === 'number' && typeof list[i - 1][c.delta] === 'number' ? r[c.delta] - list[i - 1][c.delta] : ''; });
      }
    }
    const fmt = {};
    for (const m of metrics) fmt[m.label] = m.fmt || defFmt(m.agg);
    for (const c of calcs) if (!c.delta) fmt[c.label] = c.fmt || (c.share || /\//.test(c.formula) ? 'pct' : 'money');
    for (const c of deltas) fmt[c.label] = fmt[c.delta] === 'pct' ? 'pctd' : 'deltan';

    // 月份横向展开
    const pc = tpl.pivot && tpl.pivot.col;
    if (pc && gb.includes(pc) && gb.length >= 2) {
      const pi = gb.indexOf(pc);
      const rIdx = ALL.filter(i => i !== pi), rowGb = rIdx.map(i => gb[i]);
      const pvs = [...new Set(rows.map(r => r[pc]))].sort();
      const vl = (tpl.pivot.metrics && tpl.pivot.metrics.length ? tpl.pivot.metrics : labels.filter(l => !deltas.some(c => c.label === l))).filter(l => labels.includes(l));
      const sameYear = new Set(pvs.map(p => p.slice(0, 4))).size <= 1;
      const pvName = p => (/^\d{4}-\d{2}$/.test(p) ? (sameYear ? `${+p.slice(5)}月` : `${p.slice(2, 4)}年${+p.slice(5)}月`) : p);
      const colName = (L, p) => (vl.length === 1 ? pvName(p) : `${L}·${pvName(p)}`);
      const totName = L => (vl.length === 1 ? '合计' : `${L}·合计`);
      const key = o => rowGb.map(h => o[h]).join('\u0001');
      const rowTot = new Map(build(rIdx).map(g => { const o = finish(g); return [key(o), o]; }));
      const out = new Map();
      const blank = o => { for (const L of vl) { for (const p of pvs) o[colName(L, p)] = 0; o[totName(L)] = 0; } };
      for (const r of rows) {
        const k = key(r);
        let o = out.get(k);
        if (!o) { o = {}; rowGb.forEach(h => { o[h] = r[h]; }); blank(o); out.set(k, o); }
        for (const L of vl) o[colName(L, r[pc])] = r[L];
      }
      for (const [k, o] of out) { const t = rowTot.get(k); if (t) for (const L of vl) o[totName(L)] = t[L]; }
      const cols = rowGb.slice();
      const pfmt = {};
      for (const L of vl) { for (const p of pvs) { cols.push(colName(L, p)); pfmt[colName(L, p)] = fmt[L]; } cols.push(totName(L)); pfmt[totName(L)] = fmt[L]; }
      const dcol = [];
      const L0 = vl[0];
      const withDelta = o => {
        if (tpl.pivot.delta && pvs.length >= 2 && L0) o['浮动'] = (o[colName(L0, pvs[pvs.length - 1])] || 0) - (o[colName(L0, pvs[pvs.length - 2])] || 0);
      };
      if (tpl.pivot.delta && pvs.length >= 2 && L0) { cols.push('浮动'); pfmt['浮动'] = fmt[L0] === 'pct' ? 'pctd' : 'deltan'; dcol.push('浮动'); }
      let prow = [...out.values()];
      prow.forEach(withDelta);
      const sortCol = sc && vl.includes(sc) ? totName(sc) : sc && cols.includes(sc) ? sc : totName(L0);
      prow.sort(cmpBy(sortCol, sc ? dir : -1));
      const sums = [];
      if (tpl.total !== false && prow.length && TOT) {
        const o = { __sum: true };
        rowGb.forEach(h => { o[h] = ''; }); o[rowGb[0]] = '合计';
        blank(o);
        for (const g of build([pi])) { const f = finish(g); for (const L of vl) o[colName(L, f[pc])] = f[L]; }
        for (const L of vl) o[totName(L)] = TOT[L];
        withDelta(o);
        sums.push(o);
      }
      return { cols, rows: prow, sums, warn, fmt: pfmt, groupCols: rowGb, calcLabels: [], deltaCols: dcol };
    }

    let subs = [];
    if (gb.length > 1) {
      const sub = new Map(build([0]).map(g => [g.kv[0], finish(g)]));
      subs = [...sub.values()].sort(cmp);
      const order = new Map(subs.map((s, i) => [s[gb[0]], i]));
      const inner = isMonthCol(gb[1]) ? cmpBy(gb[1], 1) : cmp;
      rows.sort((a, b) => (order.get(a[gb[0]]) - order.get(b[gb[0]])) || inner(a, b));
    } else rows.sort(cmp);
    const sums = [];
    if (tpl.total !== false && rows.length) {
      for (const s of subs) sums.push(Object.assign({}, s, { __sum: true, [gb[0]]: s[gb[0]] + ' 小计', [gb[1]]: '' }));
      if (TOT) { const o = Object.assign({}, TOT); o[gb[0]] = '合计'; gb.slice(1).forEach(h => { o[h] = ''; }); o.__sum = true; sums.push(o); }
    }
    return { cols: gb.concat(labels), rows, sums, warn, fmt, groupCols: gb, calcLabels: calcs.filter(c => !c.delta).map(c => c.label), deltaCols: deltas.map(c => c.label) };
  }

  // ---------------- 排行 ----------------
  function computeRanking(sources, roster, settings) {
    const agg = {}, extras = [];
    for (const s of sources) {
      const b = agg[s.label] = agg[s.label] || {};
      for (const [n, v] of Object.entries(s.data || {})) b[norm(n)] = { name: n, v: ((b[norm(n)] && b[norm(n)].v) || 0) + v };
      if (!STD.includes(s.label) && !extras.includes(s.label)) extras.push(s.label);
    }
    const names = new Map();
    for (const r of roster) if (r.name) names.set(norm(r.name), r.name);
    if (!settings.onlyRoster || !roster.length) {
      for (const b of Object.values(agg)) for (const [k, o] of Object.entries(b)) if (!names.has(k)) names.set(k, o.name);
    }
    const g = (f, k) => (agg[f] && agg[f][k] ? agg[f][k].v : 0);
    const has = f => !!agg[f];
    const rows = [];
    for (const [k, disp] of names) {
      const ro = roster.find(r => norm(r.name) === k);
      const row = { 顾问: disp, 所属小组: (ro && ro.group) || '' };
      for (const f of STD) row[f] = g(f, k);
      if (!has('有效客资') && has('总客资数量') && has('无效客资')) row['有效客资'] = row['总客资数量'] - row['无效客资'];
      row['有效客资成单率'] = row['有效客资'] ? row['成单'] / row['有效客资'] : 0;
      row['总成交(元)'] = row['成交额(元)'] + row['变动金额(元)'];
      row['月目标(元)'] = ro ? (num(ro.target) || 0) : 0;
      row['月完成率'] = row['月目标(元)'] ? row['总成交(元)'] / row['月目标(元)'] : 0;
      for (const f of extras) row[f] = g(f, k);
      rows.push(row);
    }
    rows.sort((a, b) => b['总成交(元)'] - a['总成交(元)']);
    rows.forEach((r, i) => { r['排名'] = i + 1; });
    return { cols: FIXED_COLS.concat(extras), rows };
  }
  function fmtCell(c, v) {
    if (v == null || v === '') return '';
    if (c === '有效客资成单率' || c === '月完成率') return (v * 100).toFixed(1) + '%';
    if (typeof v !== 'number') return String(v);
    if (c === '变动金额(元)' || c === '总成交(元)') return fmt2(v);
    if (!Number.isInteger(v)) return fmt2(v);
    return v.toLocaleString('en-US');
  }
  const isGreen = c => c === '排名' || c === '有效客资成单率' || c === '总成交(元)' || c === '月目标(元)';
  // 红色规则：默认 月完成率 < 时间进度（今天几号/30，最多 100%）；可切换为固定百分比
  const nowMs = () => W.__ocqNow || Date.now();
  const today8 = () => { const d = new Date(nowMs() + 8 * 3600e3); return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() }; };
  const timeProgress = () => Math.min(1, today8().d / 30);
  const redCut = st => (st.redMode === 'fixed' ? (num(st.threshold) || 80) / 100 : timeProgress());
  const isRed = (c, r, cut) => c === '月完成率' && !r.__sum && r['月目标(元)'] > 0 && r[c] < cut - 1e-12;
  function infoLine(st) {
    const t = today8(), p2 = n => String(n).padStart(2, '0');
    const tp = (timeProgress() * 100).toFixed(1) + '%';
    return `统计日期 ${t.y}-${p2(t.m)}-${p2(t.d)}，时间进度 ${tp}，` + (st.redMode === 'fixed' ? `月完成率低于 ${num(st.threshold) || 80}% 标红` : '月完成率低于时间进度标红');
  }
  function dateProgress() {
    const t = today8(), p2 = n => String(n).padStart(2, '0');
    return `统计日期 ${t.y}-${p2(t.m)}-${p2(t.d)}，时间进度 ${(timeProgress() * 100).toFixed(1)}%`;
  }
  /** 合计行（+可选 小组小计） */
  function summaryRows(rk, st) {
    const make = (label, rows) => {
      const o = { __sum: true, 排名: '', 顾问: label, 所属小组: '' };
      for (const c of rk.cols) if (!['排名', '顾问', '所属小组', '有效客资成单率', '月完成率'].includes(c)) o[c] = rows.reduce((a, r) => a + (typeof r[c] === 'number' ? r[c] : 0), 0);
      o['有效客资成单率'] = o['有效客资'] ? o['成单'] / o['有效客资'] : 0;
      o['月完成率'] = o['月目标(元)'] ? o['总成交(元)'] / o['月目标(元)'] : 0;
      return o;
    };
    const out = [];
    if (st.groupSubtotal) {
      const groups = [...new Set(rk.rows.map(r => r['所属小组'] || '（未分组）'))];
      for (const g of groups) { const o = make(g + ' 小计', rk.rows.filter(r => (r['所属小组'] || '（未分组）') === g)); o['所属小组'] = g; out.push(o); }
    }
    out.push(make('合计', rk.rows));
    return out;
  }

  // ---------------- 月份日期改写 ----------------
  function rewriteMonthDates(url, nowMs) {
    const re = /(\d{4})-(\d{2})-(\d{2})/g;
    const found = [...String(url).matchAll(re)];
    if (!found.length) return url;
    const lastDay = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
    let minIdx = Infinity;
    for (const f of found) {
      const y = +f[1], m = +f[2], d = +f[3];
      if (m < 1 || m > 12) return url;
      if (d !== 1 && d !== lastDay(y, m)) return url;   // 不是整月区间就不动
      minIdx = Math.min(minIdx, y * 12 + (m - 1));
    }
    const t = new Date((nowMs || Date.now()) + 8 * 3600e3); // UTC+8
    const off = (t.getUTCFullYear() * 12 + t.getUTCMonth()) - minIdx;
    if (!off) return url;
    const pad = n => String(n).padStart(2, '0');
    return String(url).replace(re, (s, y, m, d) => {
      const i = +y * 12 + (+m - 1) + off;
      const ny = Math.floor(i / 12), nm = i % 12 + 1;
      return `${ny}-${pad(nm)}-${pad(+d === 1 ? 1 : lastDay(ny, nm))}`;
    });
  }

  // ---------------- 导出（ExcelJS，统一样式；一个视图可含多张表） ----------------
  const stamp = () => { const d = today8(); return `${d.y}${String(d.m).padStart(2, '0')}${String(d.d).padStart(2, '0')}`; };
  const COL_WIDTH = { '排名': 8.8, '顾问': 8.8, '所属小组': 10.8, '小组': 10.8, '人数': 8.8, '成交额(元)': 14.8, '成单': 8.8, '看位数': 8.8, '总客资数量': 12.8, '无效客资': 10.8, '有效客资': 10.8, '有效客资成单率': 16.8, '变动金额(元)': 16.8, '总成交(元)': 14.8, '月目标(元)': 14.8, '月完成率': 10.8, '小组完成率': 12.8 };
  function numFmtFor(c, rows) {
    if (/率/.test(c)) return '0.0%';
    if (c === '变动金额(元)' || c === '总成交(元)') return '#,##0.00';
    if (c === '成交额(元)' || c === '月目标(元)') return '#,##0';
    if (c === '排名' || c === '人数' || COUNT_FIELDS.includes(c)) return '0';
    return rows.some(r => typeof r[c] === 'number' && !Number.isInteger(r[c])) ? '#,##0.00' : '#,##0';
  }
  const C = { head: 'FF3B3B3B', yellow: 'FFFFFBD6', green: 'FFEAF5EA', sum: 'FFF2F2F2', red: 'FFC00000', border: 'FFBBBBBB' };
  const autoWidth = (c, rows, text) => Math.min(30, Math.max(8.8, ...[c].concat(rows.map(r => text(c, r[c]))).map(s => String(s).replace(/[^\x00-\xff]/g, 'xx').length * 1.1 + 2)));

  /** 小组排名（按名单的所属小组汇总，没分组的进「未分组」） */
  const GROUP_COLS = ['排名', '小组', '人数', '成交额(元)', '成单', '看位数', '总客资数量', '无效客资', '有效客资', '有效客资成单率', '变动金额(元)', '总成交(元)', '月目标(元)', '小组完成率'];
  const SUM_COLS = ['成交额(元)', '成单', '看位数', '总客资数量', '无效客资', '有效客资', '变动金额(元)', '总成交(元)', '月目标(元)'];
  function groupRanking(rk) {
    const m = new Map();
    const mk = g => { const o = { 小组: g, 人数: 0 }; SUM_COLS.forEach(c => { o[c] = 0; }); return o; };
    const add = (o, r) => { o['人数']++; SUM_COLS.forEach(c => { o[c] += typeof r[c] === 'number' ? r[c] : 0; }); };
    const fin = o => { o['有效客资成单率'] = o['有效客资'] ? o['成单'] / o['有效客资'] : 0; o['小组完成率'] = o['月目标(元)'] ? o['总成交(元)'] / o['月目标(元)'] : 0; return o; };
    const tot = mk('合计');
    for (const r of rk.rows) {
      const g = (r['所属小组'] || '').trim() || '未分组';
      if (!m.has(g)) m.set(g, mk(g));
      add(m.get(g), r); add(tot, r);
    }
    const rows = [...m.values()].map(fin).sort((a, b) => b['小组完成率'] - a['小组完成率'] || b['总成交(元)'] - a['总成交(元)']);
    rows.forEach((r, i) => { r['排名'] = i + 1; });
    fin(tot); tot['排名'] = ''; tot.__sum = true;
    return { cols: GROUP_COLS, rows, sums: rows.length ? [tot] : [] };
  }
  const isGreenG = c => c === '排名' || c === '有效客资成单率' || c === '总成交(元)' || c === '月目标(元)';

  function rankingView(rk) {
    const st = getSettings(), cut = redCut(st), info = infoLine(st);
    const main = {
      sheet: '顾问排行', title: st.title || '顾问业绩排行', info, cols: rk.cols, rows: rk.rows, sums: rk.rows.length ? summaryRows(rk, st) : [],
      numFmt: c => numFmtFor(c, rk.rows), text: fmtCell, green: isGreen, red: (c, r) => isRed(c, r, cut),
      width: c => COL_WIDTH[c] || Math.max(10, c.length * 2 + 2), freezeCols: 2
    };
    const gr = groupRanking(rk);
    const grp = {
      sheet: '小组排名', title: '小组完成率排名', info: info.replace('月完成率', '小组完成率'), cols: gr.cols, rows: gr.rows, sums: gr.sums,
      numFmt: c => numFmtFor(c, gr.rows), text: (c, v) => (c === '小组完成率' ? fmtCell('月完成率', v) : fmtCell(c, v)), green: isGreenG,
      red: (c, r) => c === '小组完成率' && !r.__sum && r['月目标(元)'] > 0 && r[c] < cut - 1e-12,
      width: c => COL_WIDTH[c] || 10.8, freezeCols: 2
    };
    return { kind: 'rank', name: '顾问排行', file: `OCQ顾问排行_${stamp()}`, sheets: [main, grp], warn: [] };
  }
  function fmtT(f, v) {
    if (v == null || v === '') return '';
    if (typeof v !== 'number') return String(v);
    if (f === 'pct') return (v * 100).toFixed(1) + '%';
    if (f === 'pctd') return (v > 0 ? '+' : '') + (v * 100).toFixed(1) + '%';
    if (f === 'deltan') return (v > 0 ? '+' : '') + v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (f === 'int') return Math.round(v).toLocaleString('en-US');
    return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  const XF = { money: '#,##0.00', int: '#,##0', pct: '0.0%', pctd: '+0.0%;-0.0%;0.0%', deltan: '+#,##0.00;-#,##0.00;0' };
  function templateView(tpl) {
    const r = computeTemplate(tpl, rawTables(getRoster()));
    const red = tpl.red && tpl.red.col && r.cols.includes(tpl.red.col) ? tpl.red : null;
    const cut = red ? (red.mode === 'fixed' ? (num(red.value) || 0) / (r.fmt[red.col] === 'pct' ? 100 : 1) : timeProgress()) : 0;
    const info = dateProgress() + (red ? `，${red.col}低于${red.mode === 'fixed' ? ' ' + (num(red.value) || 0) + (r.fmt[red.col] === 'pct' ? '%' : '') + ' ' : '时间进度'}标红` : '');
    const text = (c, v) => fmtT(r.fmt[c], v);
    const all = r.rows.concat(r.sums);
    const safe = String(tpl.name || '报表').replace(/[\\/?*\[\]:]/g, '').slice(0, 31) || '报表';
    const sheet = {
      sheet: safe, title: tpl.name || '自定义报表', info, cols: r.cols, rows: r.rows, sums: r.sums,
      numFmt: c => XF[r.fmt[c]] || '@', text, green: c => r.groupCols.includes(c) || r.calcLabels.includes(c),
      red: (c, row) => (r.deltaCols.includes(c) ? typeof row[c] === 'number' && row[c] < -1e-12 : !!red && c === red.col && !row.__sum && typeof row[c] === 'number' && row[c] < cut - 1e-12),
      up: (c, row) => r.deltaCols.includes(c) && typeof row[c] === 'number' && row[c] > 1e-12,
      width: c => autoWidth(c, all, text), freezeCols: r.groupCols.length
    };
    return { kind: 'tpl', name: tpl.name, file: `OCQ_${safe}_${stamp()}`, sheets: [sheet], warn: r.warn };
  }
  const asView = x => (x && x.sheets ? x : rankingView(x));

  async function buildWorkbook(view) {
    const EJ = W.ExcelJS || (typeof ExcelJS !== 'undefined' ? ExcelJS : null);
    if (!EJ) throw new Error('Excel 组件没加载（检查网络后刷新页面）');
    const wb = new EJ.Workbook();
    wb.creator = 'OCQ取数';
    const thin = { style: 'thin', color: { argb: C.border } };
    const box = { top: thin, left: thin, bottom: thin, right: thin };
    for (const v of view.sheets) {
      const n = v.cols.length;
      const ws = wb.addWorksheet(v.sheet, { views: [{ state: 'frozen', ySplit: 3, xSplit: v.freezeCols || 0 }] });
      ws.columns = v.cols.map(c => ({ width: v.width(c) }));
      ws.mergeCells(1, 1, 1, n);
      ws.getCell(1, 1).value = v.title;
      ws.getCell(1, 1).font = { name: 'Microsoft YaHei', size: 14, bold: true };
      ws.getCell(1, 1).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getRow(1).height = 26;
      ws.mergeCells(2, 1, 2, n);
      ws.getCell(2, 1).value = v.info;
      ws.getCell(2, 1).font = { name: 'Microsoft YaHei', size: 10, color: { argb: 'FF595959' } };
      ws.getCell(2, 1).alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getRow(2).height = 18;
      const hr = ws.getRow(3);
      v.cols.forEach((c, j) => {
        const cell = hr.getCell(j + 1);
        cell.value = c;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.head } };
        cell.font = { name: 'Microsoft YaHei', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = box;
      });
      hr.height = 32;
      const fmts = v.cols.map(c => v.numFmt(c));
      const writeRow = (r, rowIdx, isSum) => {
        const row = ws.getRow(rowIdx);
        v.cols.forEach((c, j) => {
          const cell = row.getCell(j + 1);
          const val = r[c];
          cell.value = val == null ? '' : val;
          if (typeof val === 'number') cell.numFmt = fmts[j];
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isSum ? C.sum : v.green(c) ? C.green : C.yellow } };
          const red = v.red(c, r), up = !red && !!(v.up && v.up(c, r));
          cell.font = { name: 'Microsoft YaHei', size: 11, bold: isSum || red || up, color: { argb: red ? C.red : up ? 'FF18794E' : 'FF222222' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = box;
        });
        row.height = 22;
      };
      v.rows.forEach((r, i) => writeRow(r, 4 + i, false));
      v.sums.forEach((r, i) => writeRow(r, 4 + v.rows.length + i, true));
      ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + Math.max(1, v.rows.length), column: n } };
    }
    return wb;
  }
  async function exportExcel(x, opts) {
    const view = asView(x);
    let wb;
    try { wb = await buildWorkbook(view); } catch (e) { alert(e.message); return null; }
    const buf = await wb.xlsx.writeBuffer();
    if (opts && opts.returnBuffer) return buf;
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = view.file + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    return buf;
  }
  function drawCanvas(x) {
    const view = asView(x);
    const H = 34, TOP = 30, GAP = 18, S = 2;
    const parts = view.sheets.map(v => ({ v, widths: v.cols.map(c => Math.round(v.width(c) * 7.2) + 10), all: v.rows.concat(v.sums) }));
    const totalW = Math.max(...parts.map(p => p.widths.reduce((a, b) => a + b, 0)));
    const totalH = parts.reduce((a, p) => a + TOP + H * (p.all.length + 1), 0) + GAP * (parts.length - 1);
    const cv = document.createElement('canvas');
    cv.width = totalW * S; cv.height = totalH * S;
    const g = cv.getContext('2d');
    g.scale(S, S);
    g.fillStyle = '#fff'; g.fillRect(0, 0, totalW, totalH);
    let y0 = 0;
    for (const { v, widths, all } of parts) {
      g.fillStyle = '#595959'; g.font = '13px "Microsoft YaHei", sans-serif'; g.textAlign = 'left'; g.textBaseline = 'middle';
      g.fillText(`${v.title}　${v.info}`, 8, y0 + TOP / 2);
      g.textAlign = 'center';
      [null].concat(all).forEach((r, ri) => {
        let x0 = 0; const y = y0 + TOP + ri * H;
        v.cols.forEach((c, ci) => {
          const head = ri === 0, sum = !head && r.__sum;
          g.fillStyle = head ? '#3b3b3b' : sum ? '#f2f2f2' : v.green(c) ? '#eaf5ea' : '#fffbd6';
          g.fillRect(x0, y, widths[ci], H);
          g.strokeStyle = '#bbbbbb'; g.strokeRect(x0 + 0.5, y + 0.5, widths[ci], H);
          const red = !head && v.red(c, r), up = !head && !red && !!(v.up && v.up(c, r));
          g.fillStyle = head ? '#fff' : red ? '#c00000' : up ? '#18794e' : '#222';
          g.font = `${head || red || sum ? 'bold ' : ''}13px "Microsoft YaHei", sans-serif`;
          g.fillText(head ? c : v.text(c, r[c]), x0 + widths[ci] / 2, y + H / 2);
          x0 += widths[ci];
        });
      });
      y0 += TOP + H * (all.length + 1) + GAP;
    }
    cv.__file = view.file;
    return cv;
  }
  function downloadCanvas(cv) {
    const a = document.createElement('a');
    a.href = cv.toDataURL('image/png');
    a.download = (cv.__file || `OCQ顾问排行_${stamp()}`) + '.png';
    document.body.appendChild(a); a.click(); a.remove();
  }
  async function copyImage(x) {
    const cv = drawCanvas(x);
    try {
      const blob = await new Promise(res => cv.toBlob(res, 'image/png'));
      await Promise.race([
        navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]),
        new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 3000))
      ]);
      return '已复制图片，可直接粘贴到微信/钉钉';
    } catch (e) {
      downloadCanvas(cv);
      return '浏览器不允许直接复制，已改为下载图片';
    }
  }

  // ---------------- 按步骤抓取（手动/回放共用） ----------------
  function resolveStep(step, parsed) {
    const a = analyze(parsed);
    let nameIdx = -1;
    const byKey = k => a.cols.find(c => c.key === k);
    if (step.nameKey && byKey(step.nameKey)) nameIdx = byKey(step.nameKey).idx;
    else if (step.nameHeader) { const c = a.cols.find(c => c.header === step.nameHeader); if (c) nameIdx = c.idx; }
    if (nameIdx < 0) nameIdx = a.nameIdx;
    const picked = [];
    for (const sc of step.cols || []) {
      let c = byKey(sc.key);
      if (!c || c.header !== sc.header) {
        const same = a.cols.filter(x => x.header === sc.header);
        if (same.length) c = same.reduce((p, q) => (q.filled >= p.filled ? q : p));
      }
      if (!c) continue;
      picked.push({ idx: c.idx, key: sc.key, header: c.header, display: c.display, mode: c.numeric ? sc.mode : 'count', label: sc.label });
    }
    return { nameIdx, picked };
  }
  const fallbackIdx = parsed => Math.max(0, parsed.headers.findIndex(h => !/^__/.test(h)));
  async function grabWithStep(step, onProgress) {
    const t0 = Date.now();
    while (Date.now() - t0 < 60000 && !findBestTable()) await sleep(500);
    const t = findBestTable();
    if (!t) return { ok: false, error: '等了 60 秒没看到数据表（是否已登录？）' };
    const { nameIdx, picked } = resolveStep(step, t.parsed);
    const wantCols = (step.cols || []).length > 0;
    if (wantCols && nameIdx < 0) return { ok: false, error: '找不到顾问列' };
    if (wantCols && !picked.length) return { ok: false, error: '步骤里的列在本页都没找到' };
    const res = await scrapeAll(onProgress, nameIdx >= 0 ? nameIdx : fallbackIdx(t.parsed));
    if (!res.ok) return res;
    if (picked.length && namesLookWrong(res.rows, nameIdx)) return { ok: false, error: '没认对顾问列' };
    const nraw = storeRaw(step.pageKey, step.name, res, nameIdx);
    const built = buildSources(res, nameIdx, picked, step.pageKey, location.href, step.id);
    const head = `【${step.name || step.pageKey}】${res.pages} 页，${res.collected}${res.total != null ? '/' + res.total : ''} 条（明细 ${nraw} 行已存）` + (res.incomplete ? `  ⚠ 只抓到 ${res.collected}/${res.total} 条，可能没翻完页` : '');
    return { ok: true, sources: built.sources, lines: [head].concat(built.lines), res: { pages: res.pages, collected: res.collected, total: res.total, incomplete: res.incomplete } };
  }

  // ---------------- 回放（一键取数） ----------------
  function openTab(url) {
    if (typeof GM_openInTab === 'function') {
      try { return GM_openInTab(url, { active: false, insert: true, setParent: true }); } catch (_) {}
    }
    const w = window.open(url, '_blank');
    return { close: () => { try { w && w.close(); } catch (_) {} } };
  }
  function waitForValue(key, timeoutMs) {
    return new Promise(resolve => {
      let done = false, lid = null;
      const finish = v => {
        if (done) return; done = true; clearInterval(iv); clearTimeout(to);
        try { if (lid != null && typeof GM_removeValueChangeListener === 'function') GM_removeValueChangeListener(lid); } catch (_) {}
        resolve(v);
      };
      if (typeof GM_addValueChangeListener === 'function') {
        try { lid = GM_addValueChangeListener(key, (k, o, n) => { if (n) finish(typeof n === 'string' ? JSON.parse(n) : n); }); } catch (_) {}
      }
      const iv = setInterval(() => { const v = gmGet(key, null); if (v) finish(v); }, 800);
      const to = setTimeout(() => finish(null), timeoutMs);
    });
  }
  let replaying = false;
  async function runReplay(onLine) {
    if (replaying) return;
    const steps = getSteps();
    if (!steps.length) { onLine('还没有取数步骤：先在明细表上手动抓一次，会自动保存成步骤。'); return; }
    replaying = true;
    const run = 'r' + Date.now();
    // 替换语义：先清掉这些步骤页面的旧来源
    const pks = new Set(steps.map(s => s.pageKey));
    setSources(getSources().filter(s => !pks.has(s.pageKey)));
    const summary = [];
    try {
      for (const [i, step] of steps.entries()) {
        const url = rewriteMonthDates(step.url);
        onLine(`(${i + 1}/${steps.length}) ${step.name} … 打开中`);
        gmSet('ocq_job', { run, step: Object.assign({}, step, { url }), url: url.split('#')[0], t: Date.now() });
        const openUrl = url.includes('#') ? url : `${url}#ocqAutoGrab=${encodeURIComponent(step.id)}&ocqRun=${run}`;
        const key = `ocq_res_${run}_${step.id}`;
        const tab = openTab(openUrl);
        const r = await waitForValue(key, 240000);
        try { tab && tab.close && tab.close(); } catch (_) {}
        GM_setValue(key, '');
        if (r && r.ok) {
          const msgs = upsertSources(r.sources);
          const l = r.lines.join('\n') + (msgs.length ? '\n   ' + msgs.join('；') : '');
          addLog(l); summary.push(`✅ ${step.name}`);
          onLine(`(${i + 1}/${steps.length}) ${step.name} ✅ ${r.res.collected}${r.res.total != null ? '/' + r.res.total : ''} 条`);
        } else {
          const err = (r && r.error) || '超时（4 分钟没结果）';
          addLog(`❌ 【${step.name}】${err}`); summary.push(`❌ ${step.name}：${err}`);
          onLine(`(${i + 1}/${steps.length}) ${step.name} ❌ ${err}`);
        }
      }
    } finally {
      gmSet('ocq_job', null);
      replaying = false;
    }
    onLine('一键取数完成：' + summary.join('　'));
  }

  // 工作标签页：检测自动抓取标记
  function detectJob() {
    const m = location.hash.match(/ocqAutoGrab=([^&]+)/);
    const job = gmGet('ocq_job', null);
    if (m) {
      const id = decodeURIComponent(m[1]);
      const run = (location.hash.match(/ocqRun=([^&]+)/) || [])[1];
      if (job && job.step && job.step.id === id) return job;
      const step = getSteps().find(s => s.id === id);
      return step ? { run, step } : null;
    }
    if (job && job.step && Date.now() - job.t < 240000 && location.href.split('#')[0] === job.url) return job;
    return null;
  }
  async function runWorker(job) {
    const toast = document.createElement('div');
    toast.style.cssText = 'position:fixed;left:50%;top:12px;transform:translateX(-50%);z-index:2147483647;background:#ff6a00;color:#fff;padding:8px 16px;border-radius:6px;font:14px "Microsoft YaHei",sans-serif;box-shadow:0 2px 10px rgba(0,0,0,.3)';
    toast.textContent = `OCQ取数：自动抓取「${job.step.name}」…`;
    document.documentElement.appendChild(toast);
    let r;
    try {
      r = await grabWithStep(job.step, (p, n, tot, note) => { toast.textContent = `OCQ取数：${job.step.name} ${note || `第${p}页 ${n}/${tot != null ? tot : '?'}`}`; });
    } catch (e) { r = { ok: false, error: e.message }; }
    gmSet(`ocq_res_${job.run}_${job.step.id}`, r);
    toast.textContent = r.ok ? 'OCQ取数：完成，本页即将关闭' : 'OCQ取数：失败 ' + r.error;
    setTimeout(() => { try { window.close(); } catch (_) {} }, r.ok ? 600 : 4000);
  }

  // ---------------- 界面 ----------------
  const host = document.createElement('div');
  host.id = 'ocq-qushu-host';
  host.style.cssText = 'position:fixed;z-index:2147483646;right:0;bottom:0;width:0;height:0;';
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `<style>
  *{box-sizing:border-box;font-family:"Microsoft YaHei",sans-serif}
  .fab{position:fixed;right:22px;bottom:26px;width:58px;height:58px;border-radius:50%;background:#ff6a00;color:#fff;border:none;font-size:16px;font-weight:bold;box-shadow:0 4px 14px rgba(0,0,0,.3);cursor:pointer}
  .panel{position:fixed;top:0;right:-680px;width:660px;max-width:97vw;height:100vh;background:#fff;box-shadow:-4px 0 20px rgba(0,0,0,.25);transition:right .25s;display:flex;flex-direction:column;font-size:13px;color:#222}
  .panel.open{right:0}
  .hd{padding:10px 14px;background:#ff6a00;color:#fff;display:flex;align-items:center;gap:8px}
  .hd b{font-size:15px}.hd .sp{flex:1}.hd button{background:none;border:none;color:#fff;font-size:20px;cursor:pointer}
  .hd .hbtn{font-size:12px;border:1px solid #fff;border-radius:4px;padding:3px 9px}.hd select{padding:3px 6px;border-radius:4px;border:none;font-size:13px;max-width:240px}
  .bd{flex:1;overflow:auto;padding:10px 14px}
  .info{background:#fff7ef;border-radius:6px;padding:8px 10px;margin-bottom:8px}.info b{color:#c55a00}
  .big{display:flex;gap:8px;margin:6px 0}.big .btn{flex:1;font-size:15px;padding:12px 6px;margin:0}
  h4{margin:12px 0 6px;font-size:13px;border-left:3px solid #ff6a00;padding-left:6px}
  table{border-collapse:collapse;width:100%}
  .cols td,.cols th{border-bottom:1px solid #eee;padding:3px 4px;text-align:left;vertical-align:middle}
  .cols th{background:#fafafa;font-weight:normal;color:#666}
  input[type=text]{padding:2px 4px;border:1px solid #ccc;border-radius:3px}
  .cols .custom{width:100px}
  .tag{font-size:11px;padding:1px 5px;border-radius:3px;background:#eef;color:#447}.tag.n{background:#efe;color:#373}
  .muted{color:#888;font-size:12px}
  .btn{padding:6px 12px;border:1px solid #ff6a00;background:#fff;color:#ff6a00;border-radius:4px;cursor:pointer;margin:3px 4px 3px 0}
  .btn.p{background:#ff6a00;color:#fff}.btn:disabled{opacity:.5;cursor:default}
  .grab{width:100%;border-style:dashed;padding:9px}
  .mini{padding:0 6px;border:1px solid #ccc;background:#fff;border-radius:3px;cursor:pointer;margin-left:2px}
  .prog{margin:6px 0;color:#c55a00;font-weight:bold;white-space:pre-wrap}
  .log{white-space:pre-wrap;font-size:12px;color:#444;background:#fafafa;border-radius:4px;padding:6px;max-height:150px;overflow:auto}
  .list div{display:flex;align-items:center;gap:4px;padding:3px 0;border-bottom:1px dashed #eee}
  .list .nm{flex:1}
  .ok{color:#18794e}.warn{color:#a15c00;font-size:12px}
  .rk{font-size:12px;margin-bottom:10px}.rk th{background:#3b3b3b;color:#fff;padding:4px 6px;border:1px solid #888;white-space:nowrap}
  .rk td{border:1px solid #ccc;padding:4px 6px;text-align:center;background:#fffbd6;white-space:nowrap}
  .rk td.g{background:#eaf5ea}.rk td.s{background:#f2f2f2;font-weight:bold}.rk td.red{color:#c00000;font-weight:bold}.rk td.up{color:#18794e;font-weight:bold}
  .rk caption{text-align:left;font-weight:bold;padding:4px 0}
  .rkwrap{overflow:auto;max-height:420px}
  .more{display:inline-block;margin-top:8px;color:#888;font-size:12px;text-decoration:none}
  details{border:1px solid #eee;border-radius:4px;margin:6px 0;padding:4px 8px}summary{cursor:pointer;font-weight:bold;padding:2px 0}
  .editor{border:2px solid #ff6a00;border-radius:6px;padding:8px 10px;margin:8px 0;background:#fffdfa}
  .editor>div{margin:5px 0}.er{display:flex;gap:4px;align-items:center;margin:3px 0}
  .wq{font-size:16px;font-weight:bold;margin:6px 0 8px}.chips{display:flex;flex-wrap:wrap;gap:6px}
  .chip{margin:0;padding:8px 12px;font-size:14px}.chip.on{background:#ff6a00;color:#fff}.chip.on .muted{color:#ffe}
  .checks{display:flex;flex-direction:column;gap:6px}.ck{font-size:14px;padding:6px 8px;border:1px solid #eee;border-radius:4px;cursor:pointer}
  .wbig{font-size:15px;padding:9px 18px}.wnav{display:flex;align-items:center;gap:6px;margin-top:12px}.wnav .sp{flex:1}
  .wsteps{display:flex;align-items:center;gap:4px;margin-bottom:6px}.wsteps span{width:22px;height:22px;border-radius:50%;background:#eee;color:#888;display:inline-flex;align-items:center;justify-content:center;font-size:12px}
  .wsteps span.cur{background:#ff6a00;color:#fff}.wsteps span.done{background:#ffd9bf;color:#c55a00}.wsteps i{width:24px;height:2px;background:#eee}
  .adv{margin-top:10px}.chip.sm{padding:3px 8px;font-size:12px}.fr{margin:4px 0;padding:4px 0;border-bottom:1px dashed #eee}.sec{margin:6px 0}
  textarea{width:100%;font-size:12px}
  </style>
  <button class="fab" id="fab" title="OCQ取数">OCQ</button>
  <div class="panel" id="panel">
    <div class="hd"><b>OCQ取数</b><select id="tplSel" title="选择报表"></select><button class="hbtn" id="tplEdit" style="display:none" title="修改这个报表">编辑</button><span class="sp"></span><button id="close" title="关闭">×</button></div>
    <div class="bd">
      <div class="info"><div id="rinfo"></div><div class="muted" id="status"></div></div>
      <div class="big">
        <button class="btn p" id="replay">⚡ 一键取数</button>
        <button class="btn" id="xlsx">📥 导出Excel</button>
        <button class="btn" id="img">🖼 复制图片</button>
      </div>
      <div id="grabBar" style="display:none"><button class="btn grab" id="go">抓当前这张表</button><div class="muted"><label><input type="checkbox" id="addStep"> 加入每日取数步骤</label>　<span id="tblinfo"></span></div></div>
      <div class="prog" id="prog"></div>
      <div class="editor" id="newBox" style="display:none">
        <div class="wq">快速报表 <span class="muted">点一下就出</span><button class="mini" id="newClose" style="float:right">×</button></div>
        <div class="chips" id="presetBtns"></div>
        <div class="wq" style="font-size:14px;margin-top:10px">Facebook 分析　<span class="muted">月份</span> <input type="month" id="pFrom"> ~ <input type="month" id="pTo"></div>
        <div class="chips" id="fbBtns"></div><div class="warn" id="newMsg"></div>
        <button class="btn p wbig" id="ownBtn" style="margin-top:8px">＋ 自己建一个</button>
      </div>
      <div class="editor" id="wiz" style="display:none"><div id="wizBody"></div></div>
      <div class="warn" id="warn"></div>
      <div class="rkwrap" id="rkwrap"></div>
      <a href="#" class="more" id="moreBtn">⚙ 更多</a>
      <div id="more" style="display:none">
        <div class="sec"><b>我的报表</b><div class="list" id="tplList"></div><a href="#" class="more" id="tplNew">＋ 新建报表</a></div>
        <details id="stepBox"><summary>取数步骤 <span class="muted" id="stepCnt"></span></summary><div class="list" id="steps"></div></details>
        <a href="#" class="more" id="advBtn">高级设置</a>
        <div id="adv" style="display:none">
        <details id="pickBox"><summary>这张表抓哪些列（一般不用改）</summary>
          <div>人员列：<select id="namecol"></select> <button class="mini" id="rescan">重新识别</button></div>
          <table class="cols"><thead><tr><th></th><th>表头</th><th>类型</th><th>怎么算</th><th>算进排行的哪一列</th></tr></thead><tbody id="cols"></tbody></table>
        </details>
        <details><summary>已抓的数据</summary><div class="list" id="sources"></div><div class="muted" id="rawInfo"></div></details>
        <details><summary>报表导出 / 导入（JSON）</summary>
          <div><button class="btn" id="tplExp">导出全部报表</button><button class="btn" id="tplImp">导入</button></div>
          <textarea id="tplJson" rows="3" placeholder="导入：把导出的 JSON 粘到这里再点「导入」"></textarea><div class="muted" id="tplMsg"></div>
        </details>
        <details><summary>顾问名单（小组 / 月目标）</summary>
          <div class="muted">可粘贴整张排行表（含表头），会自动认出 顾问 / 所属小组 / 月目标</div>
          <textarea id="paste" rows="5"></textarea>
          <div><button class="btn" id="imp">导入名单</button><button class="btn" id="clrRoster">清空名单</button><span class="muted" id="pmsg"></span></div>
          <div class="muted" id="rosterInfo"></div>
        </details>
        <details><summary>顾问排行设置（标红 / 小组小计 / 标题）</summary>
          <div>标红：<label><input type="radio" name="rm" id="rmP"> 低于时间进度（今天几号÷30）</label>　<label><input type="radio" name="rm" id="rmF"> 低于固定</label> <input type="text" id="th" style="width:44px">%</div>
          <div><label><input type="checkbox" id="grpSub"> 显示小组小计</label>　<label><input type="checkbox" id="only"> 只显示名单内顾问</label></div>
          <div>标题 <input type="text" id="title" style="width:160px"> <button class="btn" id="saveSet">保存</button> <span class="muted" id="smsg"></span></div>
        </details>
        <details><summary>清空 / 记录</summary>
          <button class="btn" id="clear">清空已抓数据</button><span class="muted">（名单、步骤、报表都保留）</span>
          <div class="log" id="log"></div>
        </details>
        </div>
      </div>
    </div>
  </div>`;
  document.documentElement.appendChild(host);
  const $ = id => root.getElementById(id);

  let scan = null, busy = false, ed = null, scanRange = { hist: false, tag: '' };
  const targetOpts = sel => STD.map(f => `<option value="${f}" ${sel === f ? 'selected' : ''}>${f.replace('(元)', '')}</option>`).join('') + `<option value="__custom" ${sel === '__custom' ? 'selected' : ''}>自定义…</option>`;
  const tm = ms => { const d = new Date(ms + 8 * 3600e3); return `${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')} ${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`; };

  function currentChoices() {
    const out = {};
    root.querySelectorAll('#cols tr[data-key]').forEach(tr => {
      out[tr.dataset.key] = { on: tr.querySelector('.on').checked, mode: tr.querySelector('.mode').value, target: tr.querySelector('.target').value, custom: tr.querySelector('.custom').value.trim() };
    });
    return out;
  }
  function savePrefs() {
    if (!scan) return;
    const p = getPrefs();
    const nc = scan.cols[+$('namecol').value];
    p[pageKey()] = { nameKey: nc ? nc.key : null, cols: currentChoices() };
    setPrefs(p);
    root.querySelectorAll('#cols tr[data-key]').forEach(tr => { tr.querySelector('.custom').style.display = tr.querySelector('.target').value === '__custom' ? '' : 'none'; });
  }
  /** 这页筛的是哪段时间：先看页面上的日期框/网址，再看表里的日期列 */
  function pageRange(parsed) {
    let ds = [...String(location.href).matchAll(/(\d{4})-(\d{2})-(\d{2})/g)].map(m => m[0]);
    document.querySelectorAll('input').forEach(i => { const d = parseDate(i.value); if (d && visible(i)) ds.push(d); });
    let src = 'filter';
    if (!ds.length && parsed) {
      src = 'data';
      const dc = detectDateCols(parsed.headers, parsed.rows);
      const pd = primaryDate(parsed.headers, dc);
      if (pd >= 0) ds = parsed.rows.map(r => parseDate(r[pd])).filter(Boolean);
    }
    if (!ds.length) return { hist: false, tag: '' };
    ds.sort();
    const from = ds[0], to = ds[ds.length - 1], t = today8(), cur = `${t.y}-${p2(t.m)}`;
    const hist = !(from.slice(0, 7) === cur && to.slice(0, 7) === cur);
    const a = from.slice(0, 7), b = to.slice(0, 7);
    const tag = a === b ? a : a.slice(0, 4) === b.slice(0, 4) ? `${a}~${b.slice(5)}` : `${a}~${b}`;
    return { hist, tag, from, to, src };
  }
  function doScan() {
    const t = findBestTable();
    if (!t) { $('grabBar').style.display = 'none'; $('pickBox').style.display = 'none'; scan = null; return; }
    $('grabBar').style.display = ''; $('pickBox').style.display = '';
    scanRange = pageRange(t.parsed);
    $('addStep').checked = !scanRange.hist;
    const a = analyze(t.parsed);
    scan = Object.assign({ parsed: t.parsed }, a);
    const saved = getPrefs()[pageKey()];
    let nameIdx = a.nameIdx;
    if (saved && saved.nameKey) { const c = a.cols.find(c => c.key === saved.nameKey); if (c) nameIdx = c.idx; }
    const tot = parseTotal();
    $('tblinfo').textContent = `这页「${pageKey()}」` + (tot != null ? `共 ${tot} 条` : `${t.parsed.rows.length} 行`) + (getSteps().some(s => s.pageKey === pageKey()) ? ' · 已在取数步骤里' : '') + (scanRange.hist ? ` · 是 ${scanRange.tag} 的数据（存成历史数据）` : '');
    const shown = a.cols.filter(c => !/^__sel/.test(c.header));
    $('namecol').innerHTML = `<option value="-1">（没有）</option>` + shown.map(c => `<option value="${c.idx}" ${c.idx === nameIdx ? 'selected' : ''}>${esc(c.display)}</option>`).join('');
    if (nameIdx < 0) $('namecol').value = '-1';
    $('cols').innerHTML = shown.map(c => {
      const ch = Object.assign({}, a.defaults[c.key], saved && saved.cols && saved.cols[c.key] || {});
      const sample = (t.parsed.rows.find(r => (r[c.idx] || '').trim()) || [])[c.idx] || '';
      return `<tr data-key="${esc(c.key)}">
        <td><input type="checkbox" class="on" ${ch.on ? 'checked' : ''}></td>
        <td title="例：${esc(sample)}">${esc(c.display)}${c.filled ? '' : ' <span class="muted">(空)</span>'}<div class="muted">${esc(String(sample).slice(0, 16))}</div></td>
        <td><span class="tag ${c.numeric ? 'n' : ''}">${c.numeric ? '数字' : '文字'}</span></td>
        <td><select class="mode">${c.numeric ? `<option value="sum" ${ch.mode !== 'count' ? 'selected' : ''}>求和</option>` : ''}<option value="count" ${ch.mode === 'count' || !c.numeric ? 'selected' : ''}>计数</option></select></td>
        <td><select class="target">${targetOpts(ch.target)}</select> <input type="text" class="custom" value="${esc(ch.custom || c.display)}" style="${ch.target === '__custom' ? '' : 'display:none'}"></td>
      </tr>`;
    }).join('');
  }
  function stepCount(s) {
    const r = getRaw()[s.pageKey];
    if (r) return r.rows.length;
    const ss = getSources().filter(x => x.pageKey === s.pageKey);
    return ss.length ? Math.max(...ss.map(x => x.rows || 0)) : null;
  }
  function renderSteps() {
    const st = getSteps();
    $('stepCnt').textContent = st.length ? `（${st.length} 步）` : '';
    $('steps').innerHTML = st.length ? st.map((s, i) => { const n = stepCount(s); return `<div data-i="${i}"><span class="nm" title="${esc(s.url)}">${esc(s.name)}<span class="muted">${n != null ? ` · ${n} 条` : ' · 还没抓'}</span></span><button class="mini" data-a="del" title="删除这一步">×</button></div>`; }).join('') : '<span class="muted">（还没有：打开一张明细表，点「抓当前这张表」）</span>';
    const hist = Object.values(getRaw()).filter(r => r.range && !st.some(s => s.pageKey === r.pageKey));
    if (hist.length) $('steps').innerHTML += '<div class="muted" style="border:none">历史数据（不参加每日取数）：</div>' + hist.map(r => `<div data-raw="${esc(r.pageKey)}"><span class="nm">${esc(r.pageKey)}<span class="muted"> · ${r.rows.length} 条</span></span><button class="mini" data-a="rmraw" title="删除这份历史数据">×</button></div>`).join('');
  }
  function renderSources() {
    const ss = getSources();
    $('sources').innerHTML = ss.length ? ss.map((s, i) => `<div data-i="${i}"><span class="nm">${esc(s.label)} ← ${esc(s.colHeader.replace(/^__.*/, '行数'))}（${s.mode === 'sum' ? '求和' : '计数'}） <span class="muted">${esc(s.pageKey)} · ${s.rows} 条${s.total != null ? '/' + s.total : ''} · ${tm(s.time)}${s.footer ? (s.footer.ok ? ' · 合计✅' : ' · 合计⚠') : ''}</span></span><button class="mini" data-a="rm">×</button></div>`).join('') : '<span class="muted">（还没有）</span>';
    const raw = Object.values(getRaw());
    $('rawInfo').textContent = raw.length ? '明细（自定义报表用）：' + raw.map(t => `${t.name} ${t.rows.length} 行`).join('；') : '';
  }
  function statusText() {
    const pks = new Set(getSources().map(s => s.pageKey).concat(Object.keys(getRaw())));
    const times = getSources().map(s => s.time).concat(Object.values(getRaw()).map(t => t.time));
    return pks.size ? `已抓 ${pks.size} 张表 · 最近更新 ${tm(Math.max(...times))}（UTC+8）` : '还没抓数据：点「⚡ 一键取数」，或打开明细表点「抓当前这张表」';
  }
  function currentView() {
    const st = getSettings();
    const tpl = st.tpl && getTpls().find(t => t.id === st.tpl);
    if (tpl) return templateView(tpl);
    return rankingView(computeRanking(getSources(), getRoster(), st));
  }
  function renderTplSel() {
    const st = getSettings(), tp = getTpls();
    if (st.tpl && !tp.some(t => t.id === st.tpl)) { st.tpl = ''; setSettings(st); }
    $('tplSel').innerHTML = `<option value="">顾问排行（默认）</option>` + tp.map(t => `<option value="${esc(t.id)}">${esc(t.name)}</option>`).join('') + `<option value="__new">＋ 新建报表</option>`;
    $('tplSel').value = st.tpl || '';
    $('tplEdit').style.display = st.tpl ? '' : 'none';
  }
  function renderRanking() {
    $('status').textContent = statusText();
    let v;
    try { v = currentView(); } catch (e) { $('rinfo').textContent = dateProgress(); $('warn').textContent = e.message; $('rkwrap').innerHTML = ''; return null; }
    $('rinfo').innerHTML = esc(v.sheets[0].info).replace(/时间进度 ([\d.]+%)/, '时间进度 <b>$1</b>');
    $('warn').textContent = (v.warn || []).join('；');
    $('rkwrap').innerHTML = v.sheets.map((s, si) => `<table class="rk">${v.sheets.length > 1 || si ? `<caption>${esc(s.title)}</caption>` : ''}<tr>${s.cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr>` +
      (s.rows.length ? s.rows.concat(s.sums).map(r => `<tr>${s.cols.map(c => `<td class="${r.__sum ? 's' : s.green(c) ? 'g' : ''} ${s.red(c, r) ? 'red' : s.up && s.up(c, r) ? 'up' : ''}">${esc(s.text(c, r[c]))}</td>`).join('')}</tr>`).join('') : `<tr><td colspan="${s.cols.length}" class="muted">（还没有数据）</td></tr>`) + '</table>').join('');
    $('log').textContent = getLog().slice().reverse().join('\n');
    return v;
  }
  function renderSettings() {
    const st = getSettings(), ro = getRoster();
    $('th').value = st.threshold; $('title').value = st.title;
    $('rmP').checked = st.redMode !== 'fixed'; $('rmF').checked = st.redMode === 'fixed'; $('grpSub').checked = !!st.groupSubtotal; $('only').checked = !!st.onlyRoster;
    $('rosterInfo').textContent = ro.length ? `名单 ${ro.length} 人：` + ro.slice(0, 8).map(r => `${r.name}/${r.group || '-'}/${r.target || '-'}`).join('；') + (ro.length > 8 ? '…' : '') : '名单为空';
    const tp = getTpls();
    $('tplList').innerHTML = tp.length ? tp.map((t, i) => `<div data-i="${i}"><span class="nm">${esc(t.name)}</span><button class="mini" data-a="edit">编辑</button><button class="mini" data-a="del">删除</button></div>`).join('') : '<span class="muted">（还没有自定义报表）</span>';
  }
  function renderAll() { renderTplSel(); renderSteps(); renderSources(); renderRanking(); renderSettings(); }

  const opt = (list, sel, none) => (none != null ? `<option value="">${none}</option>` : '') + [...new Set(list.concat(sel ? [sel] : []))].map(h => `<option value="${esc(h)}" ${h === sel ? 'selected' : ''}>${esc(h)}</option>`).join('');
  const objOpt = (o, sel) => Object.entries(o).map(([k, v]) => `<option value="${k}" ${k === sel ? 'selected' : ''}>${v}</option>`).join('');
  // ---------- 报表向导（快速报表 + 3 步自建） ----------
  const MONEY_COLS = ['成交金额', '变动金额', '累计收款', '合同未收'];
  const isMoneyName = h => MONEY_COLS.includes(h) || /金额|收款|未收|款项|费用|价格|总价|业绩/.test(h);
  const PRESETS = [['按小组统计', '所属小组'], ['按渠道统计', '渠道来源'], ['按客服统计', '装修客服'], ['按房屋类型统计', '房屋类型'], ['按装修进度统计', '装修进度'], ['按顾问统计', '顾问']];
  /** 每个已抓表的列画像：文字/数字/金额 + 例子 */
  function colProfile() {
    const ro = getRoster();
    const tabs = rawTables(ro);
    const hasGroups = ro.some(r => (r.group || '').trim());
    return tabs.map(t => {
      const cols = t.headers.map((h, i) => {
        const vals = t.rows.map(r => r[i]).filter(v => String(v || '').trim());
        const nums = vals.filter(v => num(v) != null).length;
        const numeric = vals.length > 0 && nums / vals.length >= 0.8 && !/电话|手机|编号|单号|日期|时间|月份|ID|id|^#$|^序号$/.test(h);
        return { h, numeric, money: numeric && isMoneyName(h), example: vals[0] || '', filled: vals.length, distinct: new Set(vals).size };
      }).filter(c => c.filled && !(c.h === '所属小组' && !hasGroups && !t.headers.slice(0, -2).includes('所属小组')));
      return { t, cols };
    });
  }
  const labelFor = (src, agg) => (src === '__rows' ? '单数' : agg === 'sum' ? src : src + AGG_NAMES[agg]);
  const fmtForMetric = (m, prof) => (m.agg === 'rows' || m.agg === 'count' || m.agg === 'distinct' ? 'int' : 'money');
  const fmtForCalc = c => (c.share || /\//.test(c.formula || '') ? 'pct' : 'money');
  /** 选出最合适的那张表：含全部分组列 + 选中数字列最多 + 行最多 */
  function bestTable(prof, gb, wantCols) {
    const ok = prof.filter(p => gb.every(g => p.t.headers.includes(g)));
    ok.sort((a, b) => wantCols.filter(c => b.t.headers.includes(c)).length - wantCols.filter(c => a.t.headers.includes(c)).length || b.t.rows.length - a.t.rows.length);
    return ok[0] || null;
  }
  function availablePresets() {
    const prof = colProfile();
    const out = [];
    for (const [name, g] of PRESETS) {
      const p = bestTable(prof, [g], MONEY_COLS);
      if (!p) continue;
      const c = p.cols.find(x => x.h === g);
      if (!c || c.numeric) continue;
      out.push({ name, g, p });
    }
    return out;
  }
  function buildPreset(name, g) {
    const p = bestTable(colProfile(), [g], MONEY_COLS);
    const money = MONEY_COLS.filter(h => p.t.headers.includes(h));
    const metrics = [{ src: '__rows', agg: 'rows', label: '单数', fmt: 'int', table: p.t.pageKey }].concat(money.map(h => ({ src: h, agg: 'sum', label: h, fmt: 'money', table: p.t.pageKey })));
    return { name, groupBy: [g], metrics, calcs: [{ label: '占比', share: '单数', formula: '', fmt: 'pct' }], sort: { col: money[0] || '单数', dir: 'desc' }, red: {}, total: true, preset: true };
  }
  function applyPreset(name, g) {
    const tp = getTpls();
    const old = tp.find(t => t.preset && t.groupBy[0] === g);
    const t = saveTemplate(Object.assign(buildPreset(name, g), old ? { id: old.id } : {}));
    hideBoxes(); renderAll();
    return t;
  }
  // ---------- FB 报表（带日期范围） ----------
  const p2 = n => String(n).padStart(2, '0');
  function lastFullMonths(n) {
    const t = today8();
    let y = t.y, m = t.m - 1;
    if (m < 1) { m = 12; y--; }
    const toM = `${y}-${p2(m)}`;
    let fy = y, fm = m - (n - 1);
    while (fm < 1) { fm += 12; fy--; }
    return { from: `${fy}-${p2(fm)}`, to: toM };
  }
  const monthEnd = ym => { const [y, m] = ym.split('-').map(Number); return `${ym}-${p2(new Date(Date.UTC(y, m, 0)).getUTCDate())}`; };
  const chanCol = p => p.t.headers.find(h => /渠道/.test(h) && !/^月份/.test(h)) || '';
  const isOrderTab = p => p.t.headers.includes('成交金额');
  const isValidTab = p => !isOrderTab(p) && /有效/.test(p.t.name) && !/无效/.test(p.t.name);
  const validityCol = p => p.t.headers.find(h => /是否有效|有效性|判定结果|线索状态|客资状态|有效状态/.test(h)) || '';
  const isLeadTab = p => !isOrderTab(p) && !isValidTab(p) && !/无效/.test(p.t.name) && (/线索|客资/.test(p.t.name) || !!validityCol(p));
  const houseCol = p => p.t.headers.find(h => /房子|房屋/.test(h) && !/^月份/.test(h)) || '';
  function fbRoles(fromYM, toYM) {
    const from = fromYM + '-01', to = monthEnd(toYM);
    const prof = colProfile();
    const ok = p => chanCol(p) && p.t.headers.includes('日期') && p.t.headers.includes('顾问');
    const score = p => { const i = p.t.headers.indexOf('日期'); return p.t.rows.filter(r => r[i] && r[i] >= from && r[i] <= to).length; };
    const best = f => prof.filter(p => ok(p) && f(p)).sort((a, b) => score(b) - score(a) || b.t.time - a.t.time)[0] || null;
    return { from, to, order: best(isOrderTab), lead: best(isLeadTab), valid: best(isValidTab) };
  }
  function fbConds(R, tabs) {
    const cols = [...new Set(tabs.filter(Boolean).map(chanCol).filter(Boolean))];
    return cols.map(c => ({ col: c, op: 'contains', value: 'facebook' })).concat([{ col: '日期', op: 'between', from: R.from, to: R.to }]);
  }
  function fbPresets(fromYM, toYM) {
    const R = fbRoles(fromYM, toYM), out = [];
    const rng = `${fromYM}~${toYM}`;
    if (R.order) {
      out.push({ key: 'fb1', name: 'FB·每人每月成交', tpl: { name: 'FB·每人每月成交', groupBy: ['顾问', '月份'], filters: fbConds(R, [R.order]), metrics: [{ src: '成交金额', agg: 'sum', label: '成交金额', fmt: 'money', table: R.order.t.pageKey }], calcs: [], pivot: { col: '月份', delta: false }, sort: { col: '成交金额', dir: 'desc' }, red: {}, total: true } });
      const hc = houseCol(R.order);
      if (hc) out.push({ key: 'fb2', name: 'FB·房子情况', tpl: { name: 'FB·房子情况', groupBy: [hc], filters: fbConds(R, [R.order]), metrics: [{ src: '__rows', agg: 'rows', label: '单数', fmt: 'int', table: R.order.t.pageKey }, { src: '成交金额', agg: 'sum', label: '成交金额', fmt: 'money', table: R.order.t.pageKey }], calcs: [{ label: '占比', share: '单数', formula: '', fmt: 'pct' }], sort: { col: '单数', dir: 'desc' }, red: {}, total: true } });
      if (R.lead) {
        const ms = [{ src: '__rows', agg: 'rows', label: '客资', fmt: 'int', table: R.lead.t.pageKey }];
        let rate = '成单/客资', rateName = '成单率(按总客资)';
        if (R.valid) { ms.push({ src: '__rows', agg: 'rows', label: '有效客资', fmt: 'int', table: R.valid.t.pageKey }); rate = '成单/有效客资'; rateName = '有效客资成单率'; }
        else if (validityCol(R.lead)) { const vc = validityCol(R.lead); ms.push({ src: '__rows', agg: 'rows', label: '有效客资', fmt: 'int', table: R.lead.t.pageKey, where: [{ col: vc, op: 'contains', value: '有效' }, { col: vc, op: 'not', value: '无效' }] }); rate = '成单/有效客资'; rateName = '有效客资成单率'; }
        ms.push({ src: '__rows', agg: 'rows', label: '成单', fmt: 'int', table: R.order.t.pageKey });
        out.push({ key: 'fb3', name: 'FB·客资与成单率', tpl: { name: 'FB·客资与成单率', groupBy: ['顾问', '月份'], filters: fbConds(R, [R.order, R.lead, R.valid]), metrics: ms, calcs: [{ label: rateName, formula: rate, fmt: 'pct' }, { label: '浮动', delta: rateName, formula: '', fmt: 'pctd' }], sort: { col: '成单', dir: 'desc' }, red: {}, total: true } });
      }
    }
    const nameOf = pk => (colProfile().find(p => p.t.pageKey === pk) || { t: { name: pk } }).t.name;
    out.forEach(x => { x.tpl.presetKey = x.key; x.tpl.range = rng; x.tables = [...new Set(x.tpl.metrics.map(m => m.table))].map(nameOf); });
    return out;
  }
  function applyFb(key) {
    const from = $('pFrom').value, to = $('pTo').value;
    if (!from || !to || from > to) { $('newMsg').textContent = '请选好起止月份'; return null; }
    const x = fbPresets(from, to).find(p => p.key === key);
    if (!x) return null;
    const old = getTpls().find(t => t.presetKey === key);
    const t = saveTemplate(Object.assign(x.tpl, { name: `${x.name}（${from.slice(5).replace(/^0/, '')}–${to.slice(5).replace(/^0/, '')}月）` }, old ? { id: old.id } : {}));
    hideBoxes(); renderAll();
    return t;
  }
  function renderFbBtns() {
    const from = $('pFrom').value, to = $('pTo').value;
    const ps = from && to ? fbPresets(from, to) : [];
    $('fbBtns').innerHTML = ps.length ? ps.map(x => `<button class="btn chip" data-fb="${x.key}">${esc(x.name)}<span class="muted">　${esc(x.tables.join(' + '))}</span></button>`).join('') : '<span class="muted">需要先抓带「渠道来源」和日期的订单表（如 装修项目列表）；客资报表还要抓线索表</span>';
  }
  function hideBoxes() { $('newBox').style.display = 'none'; $('wiz').style.display = 'none'; ed = null; }
  function openNewBox() {
    hideBoxes();
    $('panel').classList.add('open');
    const ps = availablePresets();
    $('presetBtns').innerHTML = ps.length ? ps.map(x => `<button class="btn chip" data-name="${esc(x.name)}" data-g="${esc(x.g)}">${esc(x.name)}<span class="muted">　${esc(x.p.t.name)}</span></button>`).join('') : '<span class="muted">先「⚡ 一键取数」或抓一张明细表，这里会按抓到的列给出快速报表</span>';
    if (!$('pFrom').value) { const r = lastFullMonths(3); $('pFrom').value = r.from; $('pTo').value = r.to; }
    $('newMsg').textContent = '';
    renderFbBtns();
    $('newBox').style.display = '';
  }

  // 向导状态：ed = 模板草稿；wizStep = 1..3
  let wizStep = 1, wizSub = false;
  function openWizard(tpl) {
    hideBoxes();
    ed = tpl ? JSON.parse(JSON.stringify(tpl)) : { id: '', name: '', groupBy: [], metrics: [], calcs: [], sort: { col: '', dir: 'desc' }, red: { col: '', mode: 'progress', value: '' }, total: true };
    ed.groupBy = ed.groupBy || []; ed.filters = ed.filters || []; ed.metrics = ed.metrics || []; ed.calcs = ed.calcs || []; ed.sort = ed.sort || { dir: 'desc' }; ed.red = ed.red || {};
    ed.__autoName = !tpl;
    wizStep = 1; wizSub = ed.groupBy.length > 1;
    $('panel').classList.add('open');
    $('wiz').style.display = '';
    renderWizard();
  }
  /** 含全部分组列的表：主表（列最多/行最多）排第一 */
  function groupTables(prof) {
    const ok = prof.filter(p => ed.groupBy.every(g => p.t.headers.includes(g)));
    const prim = bestTable(prof, ed.groupBy, MONEY_COLS);
    return prim ? [prim].concat(ok.filter(p => p !== prim)) : ok;
  }
  const tableName = pk => ((getRaw()[pk] || {}).name || pk);
  const baseLabel = m => labelFor(m.src, m.agg) + (m.table && m.table !== ed.__prim ? ' · ' + tableName(m.table) : '');
  const metricLabel = m => baseLabel(m);
  function textCols(prof, within) {
    const m = new Map();
    for (const p of prof) {
      if (within && !within.every(g => p.t.headers.includes(g))) continue;
      const orig = (getRaw()[p.t.pageKey] || {}).headers || [];
      for (const c of p.cols) {
        if (c.numeric || m.has(c.h)) continue;
        if (c.h === '顾问' && !orig.includes('顾问')) continue;          // 已有原始人员列，不重复
        if (c.filled >= 20 && c.distinct / c.filled > 0.9) continue;      // 项目名/编号这类每行都不同，不适合分组
        m.set(c.h, c);
      }
    }
    return [...m.values()];
  }
  function numCols(prof, gb) {
    const m = new Map();
    for (const p of prof) {
      if (!gb.every(g => p.t.headers.includes(g))) continue;
      for (const c of p.cols) if (c.numeric && !m.has(c.h)) m.set(c.h, c);
    }
    return [...m.values()].sort((a, b) => b.money - a.money);
  }
  const chip = (c, on, kind) => `<button class="btn chip ${on ? 'on' : ''}" data-k="${kind}" data-h="${esc(c.h)}">${esc(c.h)}<span class="muted">（例：${esc(String(c.example).slice(0, 12))}）</span></button>`;
  const metricLabels = () => ed.metrics.map(m => m.label).concat(ed.calcs.filter(c => c.label).map(c => c.label));
  function renderWizard() {
    const prof = colProfile();
    const nav = (back, next) => `<div class="wnav">${back ? '<button class="btn wbig" data-w="back">← 上一步</button>' : ''}<button class="btn" data-w="cancel">取消</button><span class="sp"></span>${next}</div>`;
    let h = `<div class="wsteps">${[1, 2, 3].map(i => `<span class="${i === wizStep ? 'cur' : i < wizStep ? 'done' : ''}">${i}</span>`).join('<i></i>')}</div>`;
    if (!prof.length) h += '<div class="warn">还没有抓到明细表：先点「⚡ 一键取数」或打开明细表点「抓当前这张表」。</div>';
    if (wizStep === 1) {
      const tc = textCols(prof);
      h += `<div class="wq">按什么分组？</div><div class="chips">${tc.map(c => chip(c, ed.groupBy[0] === c.h, 'g1')).join('') || '<span class="muted">（没有可分组的文字列）</span>'}</div>`;
      if (ed.groupBy[0]) {
        if (!wizSub) h += `<a href="#" class="more" data-w="sub">＋ 再细分一层</a>`;
        else {
          const tc2 = textCols(prof, [ed.groupBy[0]]).filter(c => c.h !== ed.groupBy[0]);
          h += `<div class="muted" style="margin-top:8px">再细分一层（可不选）<a href="#" class="more" data-w="nosub">　不细分</a></div><div class="chips">${tc2.map(c => chip(c, ed.groupBy[1] === c.h, 'g2')).join('')}</div>`;
        }
      }
      h += filterHtml(prof);
      if (ed.groupBy[0] && !ed.groupBy[1]) h += `<details class="adv" ${ed.mergeSmall ? 'open' : ''}><summary>高级</summary><label><input type="checkbox" id="wMerge" ${ed.mergeSmall ? 'checked' : ''}> 小渠道合并成其他 <span class="muted">（只留占比 ≥1% 的前 15 个）</span></label></details>`;
      h += nav(false, `<button class="btn p wbig" data-w="next" ${ed.groupBy[0] ? '' : 'disabled'}>下一步 →</button>`);
    } else if (wizStep === 2) {
      const tabs = groupTables(prof);
      const has = (pk, src) => ed.metrics.some(m => m.table === pk && m.src === src);
      const key = (pk, src) => esc(pk + '\u0001' + src);
      h += `<div class="wq">要看哪些数？</div>`;
      tabs.forEach((p, ti) => {
        const ncs = p.cols.filter(c => c.numeric).sort((a, b) => b.money - a.money);
        h += `<div class="muted" style="margin:8px 0 4px">${ti ? `来自另一张表「${esc(p.t.name)}」<span class="muted">（按「${esc(ed.groupBy.join(' / '))}」对上）</span>` : `来自「${esc(p.t.name)}」`}</div><div class="checks">` +
          `<label class="ck"><input type="checkbox" data-m="${key(p.t.pageKey, '__rows')}" ${has(p.t.pageKey, '__rows') ? 'checked' : ''}> ${ti ? `单数 · ${esc(p.t.name)}` : '单数（条数）'}</label>` +
          ncs.map(c => `<label class="ck"><input type="checkbox" data-m="${key(p.t.pageKey, c.h)}" ${has(p.t.pageKey, c.h) ? 'checked' : ''}> ${esc(c.h)}${ti ? ' · ' + esc(p.t.name) : ''} <span class="muted">${c.money ? '金额，求和' : '求和'}</span></label>`).join('') + '</div>';
      });
      h += `<details class="adv" ${ed.metrics.some(m => m.agg !== 'sum' && m.agg !== 'rows') || ed.calcs.some(c => c.formula) ? 'open' : ''}><summary>高级</summary>` +
        (ed.metrics.filter(m => m.src !== '__rows').length ? '<div class="muted">怎么算：</div>' + ed.metrics.filter(m => m.src !== '__rows').map(m => `<div class="er">${esc(baseLabel(m))} <select data-agg="${esc(m.table + '\u0001' + m.src)}">${objOpt({ sum: '求和', avg: '平均', max: '最大', min: '最小', distinct: '去重计数' }, m.agg)}</select></div>`).join('') : '') +
        '<div class="muted" style="margin-top:6px">改名（公式里好用，如 成单、有效客资）：</div>' + ed.metrics.map(m => `<div class="er"><span class="muted" style="width:150px">${esc(m.label)}</span> → <input type="text" class="m-name" data-mk="${esc(m.table + '\u0001' + m.src)}" value="${esc(m.label)}" style="width:110px"></div>`).join('') +
        `<div class="muted" style="margin-top:6px">计算列：用上面的名字和 + - * /。例：单均 = ${esc((ed.metrics.find(m => m.src !== '__rows') || { label: '变动金额' }).label)} / 单数</div>` +
        ed.calcs.map((c, i) => c.delta ? `<div class="er" data-i="${i}">${esc(c.label)} = ${esc(c.delta)} 本月 − 上月 <button class="mini" data-w="rmC">×</button></div>` : c.share ? `<div class="er" data-i="${i}">${esc(c.label)} = ${esc(c.share)} ÷ 合计 <button class="mini" data-w="rmC">×</button></div>` : `<div class="er" data-i="${i}"><input type="text" class="c-label" value="${esc(c.label)}" placeholder="名字" style="width:80px"> = <input type="text" class="c-formula" value="${esc(c.formula)}" placeholder="如 变动金额/单数" style="flex:1"><button class="mini" data-w="rmC">×</button></div>`).join('') +
        `<button class="mini" data-w="addC">＋ 计算列</button> <button class="mini" data-w="addShare">＋ 占比（单数÷合计）</button>${ed.groupBy.some(isMonthCol) ? ' <button class="mini" data-w="addDelta">＋ 环比浮动（本月−上月）</button>' : ''}<div class="muted">例：有效客资成单率 = 成单 / 有效客资（会显示成 %）</div><div class="warn" id="wMsg"></div></details>`;
      h += nav(true, `<button class="btn p wbig" data-w="next" ${ed.metrics.length ? '' : 'disabled'}>下一步 →</button>`);
    } else {
      const labs = metricLabels();
      if (!labs.includes(ed.sort.col)) ed.sort.col = labs[0] || '';
      if (ed.__autoName || !ed.name) ed.name = `按${ed.groupBy.join('/')}统计`;
      if (ed.pivot && !ed.groupBy.some(isMonthCol)) ed.pivot = null;
      h += `<div class="wq">按哪个排名？</div><div class="checks">${labs.map(l => `<label class="ck"><input type="radio" name="wsort" value="${esc(l)}" ${ed.sort.col === l ? 'checked' : ''}> ${esc(l)}</label>`).join('')}</div>` +
        `<div style="margin:6px 0"><label><input type="radio" name="wdir" value="desc" ${ed.sort.dir !== 'asc' ? 'checked' : ''}> 从大到小</label>　<label><input type="radio" name="wdir" value="asc" ${ed.sort.dir === 'asc' ? 'checked' : ''}> 从小到大</label></div>` +
        `<div class="wq" style="font-size:14px">报表名称</div><input type="text" id="wName" value="${esc(ed.name)}" style="width:100%;font-size:15px;padding:6px">` +
        `<details class="adv" ${ed.red.col ? 'open' : ''}><summary>高级</summary>` +
        `<div class="er">标红：<select id="wRedCol">${opt(labs, ed.red.col, '（不标红）')}</select> 低于 <select id="wRedMode">${objOpt({ progress: '时间进度', fixed: '固定值' }, ed.red.mode || 'progress')}</select> <input type="text" id="wRedVal" value="${esc(ed.red.value || '')}" style="width:60px;${ed.red.mode === 'fixed' ? '' : 'display:none'}" placeholder="如 80"></div>` +
        `<div><label><input type="checkbox" id="wTotal" ${ed.total !== false ? 'checked' : ''}> 最后加合计行</label></div>` +
        (ed.groupBy.length >= 2 && ed.groupBy.some(isMonthCol) ? `<div><label><input type="checkbox" id="wPivot" ${ed.pivot ? 'checked' : ''}> 月份横向展开（每月一列 + 合计）</label> <span id="wPivDeltaBox" style="${ed.pivot ? '' : 'display:none'}"><label><input type="checkbox" id="wPivDelta" ${ed.pivot && ed.pivot.delta ? 'checked' : ''}> 加「浮动」列（最后一个月 − 上个月）</label></span></div>` : '') + `</details>`;
      h += nav(true, `<button class="btn p wbig" data-w="done">完成</button>`) + '<div class="warn" id="wMsg"></div>';
    }
    $('wizBody').innerHTML = h;
  }
  function filterCols(prof) {
    const set = new Set(), dates = new Set();
    for (const p of prof) { p.t.headers.forEach(h => { if (!/^月份|^顾问$|^所属小组$/.test(h)) set.add(h); }); (p.t.dateCols || []).forEach(h => dates.add(h)); }
    return { cols: [...set], dates };
  }
  function topValues(prof, col) {
    const cnt = new Map();
    for (const p of prof) { const i = p.t.headers.indexOf(col); if (i < 0) continue; for (const r of p.t.rows) { const v = String(r[i] || '').trim(); if (v) cnt.set(v, (cnt.get(v) || 0) + 1); } }
    return [...cnt.entries()].sort((a, b) => b[1] - a[1]).slice(0, 12).map(x => x[0]);
  }
  function filterHtml(prof) {
    const fc = filterCols(prof);
    ed.filters = ed.filters || [];
    let h = `<div class="wq" style="font-size:14px;margin-top:14px">只看哪些行？ <span class="muted">（可不设；不分大小写，多个条件要同时满足）</span></div>`;
    h += ed.filters.map((f, i) => {
      const isD = fc.dates.has(f.col);
      const vals = f.col && !isD ? topValues(prof, f.col) : [];
      return `<div class="fr" data-f="${i}"><div class="er"><select class="f-col">${opt(fc.cols, f.col, '（选一列）')}</select>` +
        (isD ? ` 在 <input type="date" class="f-from" value="${esc(f.from || '')}"> ~ <input type="date" class="f-to" value="${esc(f.to || '')}"> 之间`
          : ` <select class="f-op">${objOpt({ contains: '包含', eq: '等于', not: '不包含' }, f.op || 'contains')}</select> <input type="text" class="f-val" value="${esc(f.value || '')}" placeholder="如 facebook" style="width:120px">`) +
        ` <button class="mini" data-w="rmF">×</button></div>` +
        (vals.length ? `<div class="chips">${vals.map(v => `<button class="btn chip sm" data-w="fv" data-v="${esc(v)}">${esc(v.slice(0, 18))}</button>`).join('')}</div>` : '') + '</div>';
    }).join('');
    return h + `<a href="#" class="more" data-w="addF">＋ 加条件</a>`;
  }
  function readFilters() {
    const fc = filterCols(colProfile());
    ed.filters = [...$('wizBody').querySelectorAll('.fr')].map(r => {
      const col = r.querySelector('.f-col').value;
      if (fc.dates.has(col)) return { col, op: 'between', from: (r.querySelector('.f-from') || {}).value || '', to: (r.querySelector('.f-to') || {}).value || '' };
      return { col, op: (r.querySelector('.f-op') || {}).value || 'contains', value: ((r.querySelector('.f-val') || {}).value || '').trim() };
    });
  }
  function readWizard() {
    const q = s => $('wizBody').querySelector(s);
    if (wizStep === 1 && q('.fr,[data-w=addF]')) readFilters();
    if (wizStep === 2) {
      root.querySelectorAll('#wizBody .m-name').forEach(inp => { const [pk, src] = inp.dataset.mk.split('\u0001'); const m = ed.metrics.find(x => x.table === pk && x.src === src); const v = inp.value.trim(); if (m && v && v !== m.label) { ed.calcs.forEach(c => { if (c.delta === m.label) c.delta = v; if (c.share === m.label) c.share = v; }); m.label = v; m.__named = true; } });
      ed.calcs = ed.calcs.map((c, i) => {
        if (c.share || c.delta) return c;
        const r = $('wizBody').querySelector(`.er[data-i="${i}"]`);
        return r ? { label: r.querySelector('.c-label').value.trim(), formula: r.querySelector('.c-formula').value.trim(), fmt: '' } : c;
      });
    }
    if (wizStep === 3) {
      const s = q('input[name=wsort]:checked'); if (s) ed.sort.col = s.value;
      const d = q('input[name=wdir]:checked'); if (d) ed.sort.dir = d.value;
      const nm = q('#wName').value.trim(); if (nm !== ed.name) { ed.name = nm; ed.__autoName = false; }
      ed.red = { col: q('#wRedCol').value, mode: q('#wRedMode').value, value: q('#wRedVal').value.trim() };
      ed.total = q('#wTotal').checked;
      const pv = q('#wPivot');
      ed.pivot = pv && pv.checked ? { col: ed.groupBy.find(isMonthCol), delta: !!(q('#wPivDelta') && q('#wPivDelta').checked) } : null;
    }
  }
  /** 单数要限定在一张表上（不然多张表的行会混在一起数） */
  function fixRowsTable() {
    const prof = colProfile();
    const p = bestTable(prof, ed.groupBy, ed.metrics.filter(m => m.src !== '__rows').map(m => m.src));
    for (const m of ed.metrics) if (m.src === '__rows' && !m.table) m.table = p ? p.t.pageKey : '';
  }
  function finishWizard() {
    readWizard();
    ed.calcs = ed.calcs.filter(c => c.label && (c.formula || c.share || c.delta));
    ed.filters = activeConds(ed.filters);
    const err = validateTpl(Object.assign({}, ed, { metrics: ed.metrics.map(m => Object.assign({}, m)) }));
    if (err) { const w = $('wizBody').querySelector('#wMsg'); if (w) w.textContent = err; return null; }
    fixRowsTable();
    for (const m of ed.metrics) m.fmt = fmtForMetric(m);
    for (const c of ed.calcs) c.fmt = c.delta ? '' : fmtForCalc(c);
    for (const m of ed.metrics) delete m.__named;
    const t = Object.assign({}, ed); delete t.__autoName; delete t.preset; delete t.__prim;
    const saved = saveTemplate(t);
    hideBoxes(); renderAll();
    return saved;
  }
  function validateTpl(t) {
    if (!t.name) return '请填报表名称';
    if (!t.groupBy.length) return '请选「按什么分组」';
    if (!t.metrics.length) return '至少要选一个数';
    const labels = t.metrics.map(m => m.label).concat(t.calcs.filter(c => c.label).map(c => c.label));
    if (new Set(labels).size !== labels.length) return '名字有重复';
    for (const c of t.calcs.filter(c => !c.share && !c.delta && (c.label || c.formula))) {
      if (!c.label || !c.formula) return '计算列要同时填名字和公式';
      try { compileFormula(c.formula, labels.filter(l => l !== c.label)); } catch (e) { return `计算列「${c.label}」：${e.message}`; }
    }
    return '';
  }
  function saveTemplate(t) {
    t = Object.assign({}, t, { metrics: (t.metrics || []).filter(m => m.src), calcs: (t.calcs || []).filter(c => c.label && (c.formula || c.share || c.delta)) });
    if (!t.id) t.id = 't' + Date.now() + Math.floor(Math.random() * 1000);
    const tp = getTpls();
    const i = tp.findIndex(x => x.id === t.id);
    if (i >= 0) tp[i] = t; else tp.push(t);
    setTpls(tp);
    const st = getSettings(); st.tpl = t.id; setSettings(st);
    return t;
  }
  $('wizBody').addEventListener('click', e => {
    const b = e.target.closest('[data-w],[data-k]');
    if (!b) return;
    e.preventDefault();
    if (wizStep === 1) { const mg = $('wizBody').querySelector('#wMerge'); if (mg) ed.mergeSmall = mg.checked; }
    if (b.dataset.k === 'g1') { const h = b.dataset.h; if (ed.groupBy[0] !== h) { ed.groupBy = [h]; ed.metrics = []; } }
    else if (b.dataset.k === 'g2') { ed.groupBy = [ed.groupBy[0]].concat(ed.groupBy[1] === b.dataset.h ? [] : [b.dataset.h]); }
    else {
      const w = b.dataset.w;
      readWizard();
      if (w === 'sub') wizSub = true;
      else if (w === 'nosub') { wizSub = false; ed.groupBy = ed.groupBy.slice(0, 1); }
      else if (w === 'next') {
        if (wizStep === 1) {
          const prim = bestTable(colProfile(), ed.groupBy, MONEY_COLS);
          ed.__prim = prim ? prim.t.pageKey : '';
          // 旧报表/换了分组：给没指定表的指标补上表
          const prof = colProfile();
          ed.metrics = ed.metrics.map(m => { if (m.table) return m; const p = groupTables(prof).find(p => m.src === '__rows' || p.t.headers.includes(m.src)); return Object.assign({}, m, { table: p ? p.t.pageKey : '' }); });
        }
        if (wizStep === 1 && !ed.metrics.length) ed.metrics = [{ src: '__rows', agg: 'rows', label: '单数', table: ed.__prim }];
        if (wizStep === 2) {
          ed.calcs = ed.calcs.filter(c => c.share || c.delta || c.label || c.formula);
          const err = validateTpl(Object.assign({}, ed, { name: 'x' }));
          if (err) { $('wizBody').querySelector('#wMsg').textContent = err; return; }
        }
        wizStep = Math.min(3, wizStep + 1);
      } else if (w === 'back') wizStep = Math.max(1, wizStep - 1);
      else if (w === 'cancel') { hideBoxes(); renderTplSel(); return; }
      else if (w === 'done') { finishWizard(); return; }
      else if (w === 'addC') ed.calcs.push({ label: '', formula: '' });
      else if (w === 'addDelta') { const x = (ed.calcs.filter(c => !c.delta && !c.share).slice(-1)[0] || ed.metrics[0] || {}).label; if (x && !ed.calcs.some(c => c.delta === x)) ed.calcs.push({ label: x + '浮动', delta: x, formula: '' }); }
      else if (w === 'addF') ed.filters.push({ col: '', op: 'contains', value: '' });
      else if (w === 'rmF') ed.filters.splice(+b.closest('.fr').dataset.f, 1);
      else if (w === 'fv') { const f = ed.filters[+b.closest('.fr').dataset.f]; f.value = b.dataset.v; if (f.op === 'between') f.op = 'contains'; }
      else if (w === 'addShare') { if (!ed.calcs.some(c => c.share)) ed.calcs.push({ label: '占比', share: '单数', formula: '' }); }
      else if (w === 'rmC') ed.calcs.splice(+b.closest('.er').dataset.i, 1);
    }
    renderWizard();
  });
  $('wizBody').addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.m != null) {
      const [pk, src] = t.dataset.m.split('\u0001');
      if (t.checked) { if (!ed.metrics.some(m => m.table === pk && m.src === src)) { const m = { src, agg: src === '__rows' ? 'rows' : 'sum', table: pk }; m.label = metricLabel(m); ed.metrics.push(m); } }
      else ed.metrics = ed.metrics.filter(m => !(m.table === pk && m.src === src));
      readWizard(); renderWizard();
    } else if (t.dataset.agg != null) {
      const [pk, src] = t.dataset.agg.split('\u0001');
      const m = ed.metrics.find(x => x.table === pk && x.src === src); m.agg = t.value; if (!m.__named) m.label = metricLabel(m);
      readWizard(); renderWizard();
    } else if (t.id === 'wMerge') { ed.mergeSmall = t.checked; }
    else if (t.classList.contains('f-col')) { readFilters(); const f = ed.filters[+t.closest('.fr').dataset.f]; f.value = ''; f.from = f.from || ''; renderWizard(); }
    else if (t.id === 'wPivot') { const d = $('wizBody').querySelector('#wPivDeltaBox'); if (d) d.style.display = t.checked ? '' : 'none'; }
    else if (t.id === 'wRedMode') $('wizBody').querySelector('#wRedVal').style.display = t.value === 'fixed' ? '' : 'none';
  });
  $('fbBtns').addEventListener('click', e => { const b = e.target.closest('[data-fb]'); if (b) applyFb(b.dataset.fb); });
  $('pFrom').onchange = renderFbBtns; $('pTo').onchange = renderFbBtns;
  $('presetBtns').addEventListener('click', e => { const b = e.target.closest('[data-g]'); if (b) applyPreset(b.dataset.name, b.dataset.g); });
  $('ownBtn').onclick = () => openWizard(null);
  $('newClose').onclick = () => { hideBoxes(); renderTplSel(); };
  $('tplSel').onchange = () => {
    const v = $('tplSel').value;
    if (v === '__new') { openNewBox(); $('tplSel').value = getSettings().tpl || ''; return; }
    const st = getSettings(); st.tpl = v; setSettings(st); renderRanking();
  };
  $('tplNew').onclick = e => { e.preventDefault(); openNewBox(); };
  $('tplEdit').onclick = () => { const st = getSettings(); const t = getTpls().find(x => x.id === st.tpl); if (t) openWizard(t); };
  const getUI = () => gmGet('ocq_ui', {});
  const setUI = u => gmSet('ocq_ui', u);
  function applyUI() { const u = getUI(); $('adv').style.display = u.adv ? '' : 'none'; $('stepBox').open = !!u.steps; }
  $('advBtn').onclick = e => { e.preventDefault(); const u = getUI(); u.adv = !u.adv; setUI(u); applyUI(); };
  $('stepBox').addEventListener('toggle', () => { const u = getUI(); if (u.steps !== $('stepBox').open) { u.steps = $('stepBox').open; setUI(u); } });
  const openEditor = openWizard;
  $('tplList').addEventListener('click', e => {
    const a = e.target.dataset.a; if (!a) return;
    const i = +e.target.closest('[data-i]').dataset.i; const tp = getTpls(); const t = tp[i];
    if (a === 'edit') openWizard(t);
    else if (a === 'dup') { tp.push(Object.assign(JSON.parse(JSON.stringify(t)), { id: 't' + Date.now(), name: t.name + ' 副本', preset: false })); setTpls(tp); renderAll(); }
    else if (a === 'del') { if (!confirm(`删除报表「${t.name}」？`)) return; tp.splice(i, 1); setTpls(tp); renderAll(); }
  });
  $('tplExp').onclick = () => {
    const s = JSON.stringify(getTpls(), null, 2);
    $('tplJson').value = s;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([s], { type: 'application/json' }));
    a.download = `OCQ报表模板_${stamp()}.json`;
    document.body.appendChild(a); a.click(); a.remove();
    $('tplMsg').textContent = `已导出 ${getTpls().length} 个报表`;
  };
  function importTemplates(text) {
    let arr = JSON.parse(text);
    if (!Array.isArray(arr)) arr = [arr];
    const tp = getTpls(); let n = 0;
    for (const x of arr) {
      if (!x || typeof x !== 'object' || !x.name || !Array.isArray(x.groupBy) || !Array.isArray(x.metrics)) continue;
      const t = { id: x.id && !tp.some(y => y.id === x.id) ? String(x.id) : 't' + Date.now() + n, name: String(x.name), groupBy: x.groupBy.map(String).slice(0, 2), metrics: x.metrics.map(m => ({ src: String(m.src || ''), agg: AGG_NAMES[m.agg] ? m.agg : 'sum', label: String(m.label || ''), fmt: FMT_NAMES[m.fmt] ? m.fmt : defFmt(m.agg), table: m.table ? String(m.table) : '', where: Array.isArray(m.where) ? m.where : undefined })), calcs: (x.calcs || []).map(c => ({ label: String(c.label || ''), formula: String(c.formula || ''), share: c.share ? String(c.share) : '', delta: c.delta ? String(c.delta) : '', fmt: FMT_NAMES[c.fmt] || c.fmt === 'pctd' ? c.fmt : '' })), filters: Array.isArray(x.filters) ? x.filters : [], pivot: x.pivot && x.pivot.col ? { col: String(x.pivot.col), delta: !!x.pivot.delta } : null, presetKey: x.presetKey || '', preset: !!x.preset, sort: x.sort || {}, red: x.red || {}, total: x.total !== false, mergeSmall: !!x.mergeSmall };
      tp.push(t); n++;
    }
    setTpls(tp);
    return n;
  }
  $('tplImp').onclick = () => { try { const n = importTemplates($('tplJson').value); $('tplMsg').textContent = `导入了 ${n} 个报表`; renderAll(); } catch (e) { $('tplMsg').textContent = '不是有效的 JSON：' + e.message; } };

  // ---------- 抓取 ----------
  function saveStepFrom(nameIdx, picked) {
    const pk = pageKey();
    const steps = getSteps();
    const nc = nameIdx >= 0 ? scan.cols[nameIdx] : null;
    const step = {
      pageKey: pk, url: location.href.replace(/#?ocqAutoGrab=[^&]*(&ocqRun=[^&]*)?/, ''),
      nameKey: nc ? nc.key : null, nameHeader: nc ? nc.header : '',
      cols: picked.map(p => ({ key: p.key, header: p.header, mode: p.mode, label: p.label }))
    };
    const i = steps.findIndex(s => s.pageKey === pk);
    if (i >= 0) steps[i] = Object.assign(steps[i], step);
    else steps.push(Object.assign({ id: 's' + Date.now(), name: pk }, step));
    setSteps(steps);
    return i >= 0 ? steps[i] : steps[steps.length - 1];
  }
  function askPick(msg) {
    $('prog').textContent = msg;

  }
  async function startGrab() {
    if (busy) return;
    if (!scan) doScan();
    if (!scan) return;
    savePrefs();
    const choices = currentChoices();
    const nameIdx = +$('namecol').value;
    const picked = [];
    for (const c of scan.cols) {
      const ch = choices[c.key];
      if (!ch || !ch.on || nameIdx < 0) continue;
      const label = labelOf(ch);
      if (!label) continue;
      picked.push({ idx: c.idx, key: c.key, header: c.header, display: c.display, mode: c.numeric ? ch.mode : 'count', label });
    }
    busy = true; $('go').disabled = true;
    try {
      const res = await scrapeAll((p, n, tot, note) => { $('prog').textContent = note || `正在抓第 ${p} 页（${n}/${tot != null ? tot : '?'}）`; }, nameIdx >= 0 ? nameIdx : fallbackIdx(scan.parsed));
      if (!res.ok) { $('prog').textContent = res.error; return; }
      if (picked.length && namesLookWrong(res.rows, nameIdx)) { askPick('没认出哪列是顾问：到「⚙ 更多 › 高级设置 › 这张表抓哪些列」选好「人员列」再抓'); return; }
      const addStep = $('addStep').checked;
      let rg = pageRange(null);
      if (!rg.tag) rg = pageRange({ headers: res.headers, rows: res.rows });
      const histOnly = !addStep && rg.hist;
      const step = addStep ? saveStepFrom(nameIdx, picked) : null;
      const key = histOnly ? `${pageKey()}·${rg.tag}` : pageKey();
      const nraw = storeRaw(key, step ? step.name : key, res, nameIdx, rg.tag);
      const built = buildSources(res, nameIdx, picked, pageKey(), location.href, step ? step.id : null);
      const msgs = histOnly ? [] : upsertSources(built.sources);
      const head = `【${pageKey()}】${res.pages} 页，${res.collected}${res.total != null ? '/' + res.total : ''} 条（明细 ${nraw} 行已存）` + (res.incomplete ? `  ⚠ 只抓到 ${res.collected}/${res.total} 条，可能没翻完页` : '');
      addLog([head].concat(built.lines).join('\n') + (msgs.length ? '\n   ' + msgs.join('；') : ''));
      const fcBad = built.sources.some(s => s.footer && !s.footer.ok);
      $('prog').textContent = (res.incomplete ? `⚠ 只抓到 ${res.collected}/${res.total} 条` : (histOnly ? `✅ 抓好了：${nraw} 条，存为「${key}」（历史数据，不加入每日取数）` : step ? `✅ 抓好了：${nraw} 条，已记进取数步骤` : `✅ 抓好了：${nraw} 条（没加入每日取数）`)) + (fcBad ? '\n⚠ 和页面合计对不上，详见「更多 › 清空 / 记录」' : '');
      W.__ocqLastResult = { res: { pages: res.pages, collected: res.collected, total: res.total, incomplete: res.incomplete }, lines: built.lines, msgs, nraw };
      renderAll(); doScan();
    } catch (e) {
      $('prog').textContent = '出错：' + e.message;
    } finally {
      busy = false; $('go').disabled = false;
    }
  }
  function collapseProg(lines) {
    const steps = lines.filter(l => /^\(\d+\/\d+\)/.test(l) && /[✅❌]/.test(l));
    if (!steps.length) return false;
    const ok = steps.filter(l => l.includes('✅')).length, bad = steps.length - ok;
    const d = new Date(nowMs() + 8 * 3600e3), hm = `${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`;
    $('prog').innerHTML = `${bad ? '⚠' : '✅'} 已更新 ${ok} 张表${bad ? `，${bad} 张失败` : ''} · ${hm} <a href="#" class="more" id="progDet">详情</a><div class="log" id="progLog" style="display:none">${esc(steps.join('\n'))}</div>`;
    $('progDet').onclick = e => { e.preventDefault(); const x = $('progLog'); x.style.display = x.style.display === 'none' ? '' : 'none'; };
    return true;
  }
  async function startReplay() {
    $('panel').classList.add('open'); renderAll();
    $('replay').disabled = true;
    const lines = [];
    try {
      await runReplay(l => { lines.push(l); $('prog').textContent = lines.filter(x => !/^一键取数完成/.test(x)).slice(-1)[0] || l; renderSources(); renderRanking(); });
    } finally {
      $('replay').disabled = false;
      renderAll();
      if (!collapseProg(lines) && lines.length) $('prog').textContent = lines[lines.length - 1];
    }
  }

  function openPanel() { $('panel').classList.add('open'); doScan(); renderAll(); }
  $('fab').onclick = () => ($('panel').classList.contains('open') ? $('panel').classList.remove('open') : openPanel());
  $('replay').onclick = startReplay;
  $('close').onclick = () => $('panel').classList.remove('open');
  $('moreBtn').onclick = e => { e.preventDefault(); const m = $('more'); m.style.display = m.style.display === 'none' ? '' : 'none'; applyUI(); renderSettings(); renderSteps(); };
  $('rescan').onclick = doScan;
  $('cols').addEventListener('change', savePrefs);
  $('cols').addEventListener('input', e => { if (e.target.classList.contains('custom')) savePrefs(); });
  $('namecol').onchange = savePrefs;
  $('go').onclick = startGrab;
  $('steps').addEventListener('click', e => {
    const a = e.target.dataset.a; if (!a) return;
    if (a === 'rmraw') { const k = e.target.closest('[data-raw]').dataset.raw; if (!confirm(`删除历史数据「${k}」？`)) return; const r = getRaw(); delete r[k]; setRaw(r); renderAll(); return; }
    const i = +e.target.closest('[data-i]').dataset.i; const st = getSteps();
    if (a === 'up' && i > 0) [st[i - 1], st[i]] = [st[i], st[i - 1]];
    else if (a === 'down' && i < st.length - 1) [st[i + 1], st[i]] = [st[i], st[i + 1]];
    else if (a === 'del') { if (!confirm(`删除步骤「${st[i].name}」？`)) return; st.splice(i, 1); }
    setSteps(st); renderSteps();
  });
  $('steps').addEventListener('change', e => {
    if (!e.target.classList.contains('stepname')) return;
    const i = +e.target.closest('[data-i]').dataset.i; const st = getSteps();
    st[i].name = e.target.value.trim() || st[i].pageKey; setSteps(st);
  });
  $('sources').addEventListener('click', e => {
    if (e.target.dataset.a !== 'rm') return;
    const i = +e.target.closest('[data-i]').dataset.i; const ss = getSources();
    ss.splice(i, 1); setSources(ss); renderSources(); renderRanking();
  });
  $('xlsx').onclick = async () => { const v = renderRanking(); if (!v || !v.sheets[0].rows.length) return alert('还没有数据'); await exportExcel(v); };
  $('img').onclick = async () => { const v = renderRanking(); if (!v || !v.sheets[0].rows.length) return alert('还没有数据'); $('prog').textContent = await copyImage(v); };
  $('clear').onclick = () => { if (!confirm('清空已抓的数据？（名单、步骤、报表都保留）')) return; setSources([]); setRaw({}); setLog([]); $('prog').textContent = ''; renderAll(); };
  $('imp').onclick = () => {
    const r = parsePaste($('paste').value);
    if (r.error) { $('pmsg').textContent = r.error; $('pmsg').className = 'warn'; return; }
    if (!r.rows.length) { $('pmsg').textContent = '没识别到顾问'; $('pmsg').className = 'warn'; return; }
    const ro = getRoster();
    for (const x of r.rows) { const e = ro.find(y => y.name === x.name); if (e) Object.assign(e, x); else ro.push(x); }
    setRoster(ro);
    $('pmsg').className = 'ok';
    $('pmsg').textContent = `识别到 ${r.rows.length} 人`;
    renderSettings(); renderRanking();
  };
  $('clrRoster').onclick = () => { if (confirm('清空顾问名单？')) { setRoster([]); renderSettings(); renderRanking(); } };
  $('saveSet').onclick = () => { const st = getSettings(); st.threshold = num($('th').value) || 80; st.redMode = $('rmF').checked ? 'fixed' : 'progress'; st.groupSubtotal = $('grpSub').checked; st.onlyRoster = $('only').checked; st.title = $('title').value.trim() || '顾问业绩排行'; setSettings(st); renderRanking(); $('smsg').textContent = '已保存'; };

  // 工作模式（回放打开的后台标签页）
  const job = detectJob();
  if (job) { host.style.display = 'none'; runWorker(job); }

  window.__ocqUS = W.__ocqUS = {
    openPanel, doScan, startGrab, startReplay, runReplay, getSources, getSteps, setSteps, computeRanking, getRoster, getSettings,
    parsePaste, num, analyze, exportExcel, drawCanvas, summaryRows, infoLine, timeProgress, redCut, setSettingsForTest: setSettings, findBestTable, rewriteMonthDates, root, isBusy: () => busy, isReplaying: () => replaying, FIXED_COLS,
    getRaw, getTpls, saveTemplate, availablePresets, applyPreset, collapseProg, fbPresets, applyFb, pageRange, parseDate, lastFullMonths, getUI: () => gmGet('ocq_ui', {}), buildPreset, openNewBox, openWizard, importTemplates, computeTemplate, compileFormula, rawTables, currentView, templateView, rankingView, groupRanking, openEditor, setRosterForTest: setRoster, renderAll
  };
})();
