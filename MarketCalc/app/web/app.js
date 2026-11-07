'use strict';

/* ============================================================
   МАРКЕТ — калькулятор юнит-экономики Яндекс Маркета
   ============================================================ */

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const LS_KEY = 'marketcalc.ui.v1';

const $fmt = {
  money: v => Math.round(v || 0).toLocaleString('ru-RU') + ' ₽',
  money2: v => (v || 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₽',
  pct: v => (v * 100).toFixed(1).replace('.', ',') + '%',
  int: v => Math.round(v || 0).toLocaleString('ru-RU'),
  qty: v => (Math.abs(v - Math.round(v)) < 0.01 ? String(Math.round(v)) : (v).toFixed(1).replace('.', ',')),
};

const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function fmtDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value || ''));
  return match ? `${match[3]}.${match[2]}.${match[1]}` : String(value || '');
}

const DEFAULT_UI = {
  page: 'overview',
  selected: null,
  search: '',
  scheme: 'all',
  category: 'all',
  sortKey: 'marginRate',
  sortDir: 'desc',
  targets: null,
  rounding: null,
  edits: {},
  calcSearch: '',
};

let ui = loadUI();
let bundle = null;
let lastTotals = {};
let saveTimer = null;

function loadUI() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) return Object.assign({}, DEFAULT_UI, JSON.parse(raw));
  } catch (e) { /* ignore */ }
  return JSON.parse(JSON.stringify(DEFAULT_UI));
}

function saveUI() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    try { localStorage.setItem(LS_KEY, JSON.stringify(ui)); } catch (e) { /* ignore */ }
  }, 150);
}

/* ------------------------------------------------------------ расчёт */
function schemeOf(sku) {
  return (bundle.tariffs.schemes || {})[sku.scheme] || bundle.tariffs.schemes.FBY || {};
}

function tierPrice(tiers, volume) {
  const list = tiers || [];
  for (const tier of list) {
    const upper = tier.volume_to_l;
    if (upper == null || volume <= Number(upper)) return Number(tier.price_rub) || 0;
  }
  return list.length ? Number(list[list.length - 1].price_rub) || 0 : 0;
}

function effectiveSku(sku) {
  const edits = ui.edits[sku.offer_id] || {};
  return Object.assign({}, sku, edits);
}

function commissionRate(sku) {
  if (sku.commission_rate != null) return Number(sku.commission_rate);
  const cfg = schemeOf(sku);
  const byCat = cfg.commission_by_category || {};
  return Number(byCat[String(sku.category_id)] ?? cfg.default_commission ?? 0.17);
}

function logisticsCost(sku) {
  if (sku.logistics_rub != null) return Number(sku.logistics_rub);
  return tierPrice(schemeOf(sku).logistics_tiers, Number(sku.volume_l) || 0);
}

function calc(sku, price, drrRate) {
  const cfg = schemeOf(sku);
  const commission = price * commissionRate(sku);
  const logistics = logisticsCost(sku);

  const lm = cfg.last_mile || { rate: 0.055, min_rub: 25, max_rub: 500 };
  const lmRate = Number(sku.last_mile_rate ?? lm.rate);
  const lastMile = Math.min(Number(lm.max_rub), Math.max(Number(lm.min_rub), price * lmRate));

  const acquiring = price * Number(sku.acquiring_rate ?? (cfg.acquiring || {}).rate ?? 0);
  const returnsCfg = cfg.returns || {};
  const returnsRate = Number(sku.return_rate) || 0;
  const returns = returnsRate * (Number(returnsCfg.fee_per_return_rub) || 0
    + (returnsCfg.commission_refunded === false ? commission : 0));

  const storage = Number((cfg.storage || {}).rate_per_liter_day || 0)
    * (Number(sku.volume_l) || 0) * (Number(sku.storage_days) || 0);

  const utilization = Number((cfg.utilization || {}).fee_per_unit_rub || 0) * returnsRate;
  const fines = price * Number((cfg.fines || {}).rate || 0);
  const delivery = Number(sku.delivery_to_warehouse_rub ?? cfg.delivery_to_warehouse_rub ?? 0);

  const cogs = Number(sku.cogs_rub) || 0;
  const packaging = Number(sku.packaging_rub) || 0;
  const other = price * Number(sku.other_rate ?? bundle.settings.defaults.other_rate ?? 0.01);
  const ads = price * drrRate;

  const taxCfg = bundle.settings.tax || {};
  const vatInPrice = Boolean(sku.vat_included) || Boolean(taxCfg.vat_in_price);
  const revenue = vatInPrice ? price / (1 + Number(taxCfg.vat_rate || 0)) : price;
  const beforeTax = commission + logistics + lastMile + acquiring + storage + returns
    + utilization + fines + delivery + ads + other + cogs + packaging;
  const tax = taxCfg.regime === 'USN_PROFIT'
    ? Math.max(0, revenue - beforeTax) * Number(taxCfg.rate || 0)
    : revenue * Number(sku.tax_rate ?? taxCfg.rate ?? 0);

  const items = [
    ['commission', 'Комиссия площадки', commission],
    ['logistics', 'Логистика', logistics],
    ['lastMile', 'Последняя миля', lastMile],
    ['acquiring', 'Эквайринг', acquiring],
    ['storage', 'Хранение', storage],
    ['returns', 'Возвраты', returns],
    ['util', 'Утилизация и штрафы', utilization + fines],
    ['delivery', 'Доставка до склада', delivery],
    ['ads', 'Реклама / ДРР', ads],
    ['tax', 'Налоги', tax],
    ['cogs', 'Себестоимость', cogs],
    ['pack', 'Упаковка', packaging],
    ['other', 'Прочее', other],
  ].map(([key, label, value]) => ({ key, label, value, share: price ? value / price : 0 }));

  const total = items.reduce((acc, item) => acc + item.value, 0);
  const margin = revenue - total;
  return {
    price, revenue, items, total, margin,
    rate: revenue ? margin / revenue : 0,
    roi: cogs ? margin / cogs : 0,
  };
}

function marginRateAt(sku, price, drr) {
  return calc(sku, price, drr).rate;
}

function roundPrice(value, mode) {
  if (mode === 'none') return Math.round(value * 100) / 100;
  if (mode === 'up_1') return Math.ceil(value);
  if (mode === 'up_10') return Math.ceil(value / 10) * 10;
  if (mode === 'up_50') return Math.ceil(value / 50) * 50;
  if (mode === 'psych_99') return Math.ceil((value + 1) / 100) * 100 - 1;
  return Math.ceil(value / 10) * 10;
}

function roundingMode() {
  return ui.rounding || (bundle.settings.rounding || {}).mode || 'up_10';
}

function solvePrice(sku, targetRate, drr) {
  if (marginRateAt(sku, 1_000_000, drr) < targetRate) return null;
  let low = 0;
  let high = 1_000_000;
  for (let i = 0; i < 60; i += 1) {
    const mid = (low + high) / 2;
    if (marginRateAt(sku, mid, drr) < targetRate) low = mid;
    else high = mid;
  }
  let price = roundPrice(high, roundingMode());
  const minPrice = Number(sku.min_price_rub) || 0;
  if (minPrice && price < minPrice) price = minPrice;
  return price;
}

function breakEvenPrice(sku, drr, minMargin = 0) {
  const marginAt = price => calc(sku, price, drr).margin;
  if (marginAt(1_000_000) < minMargin) return null;
  let low = 0;
  let high = 1_000_000;
  for (let i = 0; i < 60; i += 1) {
    const mid = (low + high) / 2;
    if (marginAt(mid) < minMargin) low = mid;
    else high = mid;
  }
  return roundPrice(high, 'up_1');
}

function targets() {
  if (!ui.targets) {
    ui.targets = {
      margin: Number(bundle.settings.targets.margin_rate) || 0.2,
      drr: Number(bundle.settings.targets.drr_rate) || 0.1,
      minMargin: Number(bundle.settings.targets.min_margin_rub) || 0,
    };
  }
  return ui.targets;
}

function drrFor(sku) {
  const eff = effectiveSku(sku);
  return Number(eff.drr_rate ?? targets().drr);
}

function statusOf(margin, rate) {
  const t = targets();
  if (margin < 0) return { key: 'loss', label: 'Убыток' };
  if (margin < t.minMargin || rate + 1e-9 < t.margin) return { key: 'risk', label: 'Ниже цели' };
  return { key: 'ok', label: 'В норме' };
}

function promosFor(offerId) {
  return (bundle.promos || []).filter(p => p.offer_id === offerId);
}

function promoPrice(sku, promo, currentPrice) {
  if (promo.promo_price_rub != null) return Number(promo.promo_price_rub);
  return currentPrice * (1 - Number(promo.seller_discount_rate || 0));
}

/* ------------------------------------------------------------ форматирование */
function shareBar(value, cls = '') {  const width = Math.max(0, Math.min(100, value * 100));
  return `<div class="mbar"><div class="mbar-track"><i class="${cls}" style="width:${width}%"></i></div>
    <span class="mono">${$fmt.pct(value)}</span></div>`;
}

function statusChip(key, label) {
  return `<span class="status ${key}">${esc(label)}</span>`;
}

/* ------------------------------------------------------------ обзор */
function filteredSkus() {
  const q = (ui.search || '').trim().toLowerCase();
  return (bundle.skus || []).filter(sku => {
    if (ui.scheme !== 'all' && sku.scheme !== ui.scheme) return false;
    if (ui.category !== 'all' && String(sku.category_id) !== String(ui.category)) return false;
    if (q && !(`${sku.name} ${sku.offer_id} ${sku.category_name || ''}`.toLowerCase().includes(q))) return false;
    return true;
  });
}

function overviewRows() {
  return filteredSkus().map(sku => {
    const eff = effectiveSku(sku);
    const price = Number(eff.current_price_rub) || 0;
    const drr = drrFor(sku);
    const current = calc(eff, price, drr);
    const targetPrice = solvePrice(eff, targets().margin, targets().drr);
    const targetMargin = targetPrice ? calc(eff, targetPrice, targets().drr).margin : null;
    const status = statusOf(current.margin, current.rate);
    return { sku, eff, price, current, targetPrice, targetMargin, status };
  });
}

function renderOverview() {
  const rows = overviewRows();
  const all = bundle.skus || [];
  const revenue = rows.reduce((a, r) => a + r.price, 0);
  const profit = rows.reduce((a, r) => a + r.current.margin, 0);
  const avgRate = revenue ? profit / revenue : 0;
  const below = rows.filter(r => r.status.key !== 'ok').length;

  const promos = [];
  all.forEach(sku => {
    const eff = effectiveSku(sku);
    const price = Number(eff.current_price_rub) || 0;
    promosFor(sku.offer_id).forEach(promo => {
      const pp = promoPrice(eff, promo, price);
      const res = calc(eff, pp, Number(promo.drr_rate || targets().drr));
      if (res.margin < 0) promos.push({ sku, promo, res });
    });
  });

  $('#ovSub').textContent = `Текущие цены · цель по марже ${$fmt.pct(targets().margin)} · ДРР ${$fmt.pct(targets().drr)}`;
  $('#ovKpis').innerHTML = `
    <div class="kpi"><span class="kpi-label">SKU в портфеле</span>
      <strong class="kpi-value">${all.length}</strong>
      <span class="kpi-note">${rows.length} в фильтре</span></div>
    <div class="kpi"><span class="kpi-label">Средняя маржинальность</span>
      <strong class="kpi-value ${avgRate >= targets().margin ? 'good' : 'bad'}">${$fmt.pct(avgRate)}</strong>
      <span class="kpi-note">взвешено по цене</span></div>
    <div class="kpi"><span class="kpi-label">Требуют внимания</span>
      <strong class="kpi-value ${below ? 'bad' : 'good'}">${below}</strong>
      <span class="kpi-note">из ${rows.length} SKU</span></div>
    <div class="kpi"><span class="kpi-label">Убыточные акции</span>
      <strong class="kpi-value ${promos.length ? 'bad' : 'good'}">${promos.length}</strong>
      <span class="kpi-note">цена входа ниже нуля</span></div>`;

  renderOverviewChips();

  const sortKey = ui.sortKey;
  const dir = ui.sortDir === 'asc' ? 1 : -1;
  const sorted = rows.slice().sort((a, b) => {
    const get = r => {
      if (sortKey === 'name') return r.sku.name.toLowerCase();
      if (sortKey === 'price') return r.price;
      if (sortKey === 'margin') return r.current.margin;
      if (sortKey === 'target') return r.targetPrice || 0;
      return r.current.rate;
    };
    const va = get(a), vb = get(b);
    return va > vb ? dir : va < vb ? -dir : 0;
  });

  const head = `
    <thead><tr>
      <th data-sort="name">Товар</th>
      <th>Категория</th>
      <th data-sort="price" class="num">Цена</th>
      <th class="num">Расходы</th>
      <th data-sort="margin" class="num">Маржа</th>
      <th data-sort="marginRate">Маржинальность</th>
      <th data-sort="target" class="num">Цена под цель</th>
      <th>Статус</th>
    </tr></thead>`;

  const body = sorted.map((r, index) => `
    <tr data-sku="${esc(r.sku.offer_id)}" style="animation-delay:${Math.min(index * 8, 160)}ms">
      <td><div class="t-name">${esc(r.sku.name)}</div><div class="t-sub">${esc(r.sku.offer_id)} · ${esc(r.sku.category_name || '—')}</div></td>
      <td class="muted">${esc(r.sku.category_name || '—')}</td>
      <td class="num mono">${$fmt.money(r.price)}</td>
      <td class="num mono muted">${$fmt.money(r.current.total)}</td>
      <td class="num mono strong" style="color:${r.current.margin < 0 ? 'var(--red)' : 'var(--text)'}">${$fmt.money(r.current.margin)}</td>
      <td>${shareBar(r.current.rate, r.status.key === 'ok' ? 'good' : r.status.key === 'loss' ? 'bad' : '')}</td>
      <td class="num mono">${r.targetPrice ? $fmt.money(r.targetPrice) : '—'}</td>
      <td>${statusChip(r.status.key, r.status.label)}</td>
    </tr>`).join('');

  $('#ovTable').innerHTML = head + `<tbody>${body || `<tr><td colspan="8"><div class="empty">Ничего не найдено</div></td></tr>`}</tbody>`;
  $('#ovFoot').innerHTML = `<span>Показано ${rows.length} из ${all.length} SKU</span><span>Клик по строке открывает калькулятор</span>`;

  $$('#ovTable th[data-sort]').forEach(th => th.addEventListener('click', () => {
    const key = th.dataset.sort;
    if (ui.sortKey === key) ui.sortDir = ui.sortDir === 'asc' ? 'desc' : 'asc';
    else { ui.sortKey = key; ui.sortDir = key === 'name' ? 'asc' : 'desc'; }
    saveUI();
    renderOverview();
  }));

  $$('#ovTable tbody tr[data-sku]').forEach(tr => tr.addEventListener('click', () => {
    ui.selected = tr.dataset.sku;
    ui.page = 'calc';
    saveUI();
    render();
  }));
}

function renderOverviewChips() {
  const schemes = ['all'].concat(Array.from(new Set((bundle.skus || []).map(s => s.scheme))));
  const labels = { all: 'Все схемы', FBY: 'FBY', FBS: 'FBS', DBS: 'DBS' };
  $('#ovChips').innerHTML = schemes.map(s =>
    `<button class="chip ${ui.scheme === s ? 'active' : ''}" data-scheme="${esc(s)}">${labels[s] || s}</button>`).join('');
  $$('#ovChips .chip').forEach(chip => chip.addEventListener('click', () => {
    ui.scheme = chip.dataset.scheme;
    saveUI();
    renderOverview();
  }));

  const categories = Array.from(new Map((bundle.skus || [])
    .map(s => [String(s.category_id), s.category_name || s.category_id])).entries());
  const select = $('#ovCategory');
  const current = ui.category;
  select.innerHTML = `<option value="all">Все категории</option>` + categories.map(([id, name]) =>
    `<option value="${esc(id)}">${esc(name)}</option>`).join('');
  select.value = current || 'all';
  if (select.value !== (current || 'all')) { ui.category = 'all'; select.value = 'all'; }
  select.onchange = () => { ui.category = select.value; saveUI(); renderOverview(); };
}

/* ------------------------------------------------------------ калькулятор */
function selectedSku() {
  const skus = bundle.skus || [];
  return skus.find(s => s.offer_id === ui.selected) || skus[0] || null;
}

function renderSkuList() {
  const q = (ui.calcSearch || '').trim().toLowerCase();
  const list = (bundle.skus || []).filter(s => !q || `${s.name} ${s.offer_id}`.toLowerCase().includes(q));
  const listEl = $('#skuList');
  const scrollTop = listEl.scrollTop;
  $('#skuCount').textContent = list.length;
  listEl.innerHTML = list.map(sku => {
    const eff = effectiveSku(sku);
    const res = calc(eff, Number(eff.current_price_rub) || 0, drrFor(sku));
    const status = statusOf(res.margin, res.rate);
    return `<button class="sku-item ${sku.offer_id === ui.selected ? 'active' : ''}" data-sku="${esc(sku.offer_id)}">
      <div>
        <div class="s-name">${esc(sku.name)}</div>
        <div class="s-meta">${esc(sku.offer_id)} · ${$fmt.money(eff.current_price_rub)}</div>
      </div>
      <div class="s-val ${status.key === 'ok' ? 'good' : status.key === 'loss' ? 'bad' : 'risk'}">${$fmt.pct(res.rate)}</div>
    </button>`;
  }).join('') || '<div class="empty">Ничего не найдено</div>';

  listEl.scrollTop = scrollTop;

  $$('#skuList .sku-item').forEach(btn => btn.addEventListener('click', () => {
    ui.selected = btn.dataset.sku;
    saveUI();
    renderCalc();
  }));
}

function paramField(label, key, value, opts = {}) {
  const step = opts.step || 'any';
  const hint = opts.hint ? `<span class="hint">${esc(opts.hint)}</span>` : '';
  return `<div class="field">
    <label>${esc(label)}</label>
    <input type="number" step="${step}" min="0" data-param="${key}" value="${value}">
    ${hint}
  </div>`;
}

function renderCalc() {
  renderSkuList();
  const sku = selectedSku();
  if (!sku) {
    $('#calcMain').innerHTML = '<div class="card empty">Нет товаров. Загрузите данные.</div>';
    return;
  }
  ui.selected = sku.offer_id;
  const eff = effectiveSku(sku);
  const price = Number(eff.current_price_rub) || 0;
  const drr = drrFor(sku);
  const current = calc(eff, price, drr);
  const status = statusOf(current.margin, current.rate);
  const t = targets();
  const targetPrice = solvePrice(eff, t.margin, t.drr);
  const targetCalc = targetPrice ? calc(eff, targetPrice, t.drr) : null;
  const breakEven = breakEvenPrice(eff, drr, t.minMargin);
  const commission = commissionRate(eff);
  const logistics = logisticsCost(eff);

  const paramsHtml = [
    paramField('Цена продажи, ₽', 'current_price_rub', Math.round(price), { step: 10 }),
    paramField('Себестоимость, ₽', 'cogs_rub', Math.round(eff.cogs_rub || 0), { step: 10 }),
    paramField('Объём, л', 'volume_l', eff.volume_l || 0, { step: 0.1 }),
    paramField('Хранение, дней', 'storage_days', eff.storage_days || 0, { step: 1 }),
    paramField('Возвраты, %', 'return_rate', +(Number(eff.return_rate || 0) * 100).toFixed(1), { step: 0.5, hint: 'доля заказов' }),
    paramField('Комиссия, %', 'commission_rate', +(commission * 100).toFixed(1), { step: 0.5, hint: 'по категории' }),
    paramField('Логистика, ₽', 'logistics_rub', Math.round(logistics), { step: 5, hint: 'по объёму' }),
    paramField('ДРР, %', 'drr_rate', +(drr * 100).toFixed(1), { step: 1, hint: 'доля рекламы' }),
    paramField('Мин. цена, ₽', 'min_price_rub', Math.round(eff.min_price_rub || 0), { step: 10 }),
  ].join('');

  const breakdown = current.items.filter(item => item.value > 0.005).map(item => `
    <div class="brow">
      <div class="b-name"><span>${esc(item.label)}</span><div class="b-bar"><i style="width:${Math.min(100, item.share * 140)}%"></i></div></div>
      <div class="b-val">${$fmt.money2(item.value)}</div>
      <div class="b-share">${$fmt.pct(item.share)}</div>
    </div>`).join('');

  const promoList = promosFor(sku.offer_id);
  const promosHtml = promoList.length ? promoList.map((promo, index) => {
    const pp = promoPrice(eff, promo, price);
    const pdrr = Number(promo.drr_rate || t.drr);
    const res = calc(eff, pp, pdrr);
    const verdict = statusOf(res.margin, res.rate);
    const be = breakEvenPrice(eff, pdrr, t.minMargin);
    const maxDiscount = be && price > 0 ? Math.max(0, 1 - be / price) : null;
    const recommended = solvePrice(eff, t.margin, pdrr);
    return `<div class="promo" style="animation-delay:${index * 30}ms">
      <div class="promo-top">
        <div>
          <div class="promo-name">${esc(promo.promo_name || promo.promo_id)}</div>
          <div class="promo-dates">${esc(fmtDate(promo.start_date))}${promo.end_date ? ' — ' + esc(fmtDate(promo.end_date)) : ''} · ${esc(promo.promo_type)}</div>
        </div>
        ${statusChip(verdict.key, verdict.label)}
      </div>
      <div class="promo-grid">
        <div class="promo-cell"><div class="l">Цена акции</div><div class="v">${$fmt.money(pp)}</div></div>
        <div class="promo-cell"><div class="l">Маржа</div><div class="v" style="color:${res.margin < 0 ? 'var(--red)' : 'var(--text)'}">${$fmt.money(res.margin)}</div></div>
        <div class="promo-cell"><div class="l">Маржинальность</div><div class="v">${$fmt.pct(res.rate)}</div></div>
      </div>
      <div class="promo-stats">
        <span>Макс. скидка при марже ${$fmt.money(t.minMargin)}: <b>${maxDiscount != null ? $fmt.pct(maxDiscount) : '—'}</b></span>
        <span>Цена под цель: <b>${recommended ? $fmt.money(recommended) : '—'}</b></span>
      </div>
      <div class="promo-actions">
        <button class="price-preset" data-apply-promo="${esc(promo.promo_id)}" data-price="${Math.round(pp)}">Поставить цену акции</button>
      </div>
    </div>`;
  }).join('') : '<div class="promo-empty">Для этого SKU нет акций. Добавьте их на листе «Акции» и загрузите данные.</div>';

  const delta = targetPrice && price > 0 ? (targetPrice - price) / price : null;

  $('#calcMain').innerHTML = `
    <div class="calc-head">
      <div>
        <h2>${esc(sku.name)}</h2>
        <div class="chips-row">
          <span class="badge accent">${esc(sku.scheme)}</span>
          <span class="badge">${esc(sku.category_name || 'Без категории')}</span>
          <span class="badge">${esc(sku.offer_id)}</span>
          <span class="badge">${$fmt.qty(eff.volume_l || 0)} л</span>
        </div>
      </div>
      <button class="btn btn-ghost" id="btnResetSku"><svg class="icon"><use href="#i-refresh"/></svg><span>Сбросить правки</span></button>
    </div>

    <div class="calc-kpis">
      <div class="kpi"><span class="kpi-label">Маржа с единицы</span>
        <strong class="kpi-value ${current.margin < 0 ? 'bad' : ''}">${$fmt.money(current.margin)}</strong>
        <span class="kpi-note">${esc(status.label)}</span></div>
      <div class="kpi"><span class="kpi-label">Маржинальность</span>
        <strong class="kpi-value ${status.key === 'ok' ? 'good' : status.key === 'loss' ? 'bad' : ''}">${$fmt.pct(current.rate)}</strong>
        <span class="kpi-note">цель ${$fmt.pct(t.margin)}</span></div>
      <div class="kpi"><span class="kpi-label">ROI к себестоимости</span>
        <strong class="kpi-value">${$fmt.pct(current.roi)}</strong>
        <span class="kpi-note">прибыль / закупка</span></div>
      <div class="kpi"><span class="kpi-label">Безубыточная цена</span>
        <strong class="kpi-value">${breakEven ? $fmt.money(breakEven) : '—'}</strong>
        <span class="kpi-note">маржа ${$fmt.money(t.minMargin)}</span></div>
    </div>

    <div class="card panel">
      <h3>Параметры расчёта <button class="reset-link" id="btnResetParams">вернуть исходные</button></h3>
      <div class="param-grid">${paramsHtml}</div>
      <div class="price-row">
        <button class="price-preset" data-price-factor="-0.05">−5%</button>
        <button class="price-preset" data-price-factor="-0.1">−10%</button>
        <button class="price-preset" data-price-factor="0.05">+5%</button>
        <button class="price-preset" data-price-base="1">Текущая цена</button>
      </div>
    </div>

    <div class="card panel">
      <h3>Из чего складывается расход <span class="muted" style="font-weight:400;font-size:10px">на единицу</span></h3>
      ${breakdown}
      <div class="brow total">
        <div class="b-name">Итого расходов</div>
        <div class="b-val">${$fmt.money2(current.total)}</div>
        <div class="b-share">${$fmt.pct(price ? current.total / price : 0)}</div>
      </div>
    </div>

    <div class="scenarios">
      <div class="card panel">
        <h3>Цена под целевую маржу</h3>
        <div class="target-box">
          <div class="target-price">${targetPrice ? $fmt.money(targetPrice) : 'Недостижимо'}</div>
          <div class="target-meta">
            ${targetCalc ? `<span>Маржа <b>${$fmt.money(targetCalc.margin)}</b></span>
            <span>Маржинальность <b>${$fmt.pct(targetCalc.rate)}</b></span>` : ''}
            ${delta != null ? `<span>К текущей <b style="color:${delta >= 0 ? 'var(--green)' : 'var(--red)'}">${delta >= 0 ? '+' : ''}${(delta * 100).toFixed(1)}%</b></span>` : ''}
          </div>
          <div class="mini-grid">
            <div class="field"><label>Целевая маржа, %</label>
              <input type="number" step="1" min="0" max="80" data-target="margin" value="${+(t.margin * 100).toFixed(1)}"></div>
            <div class="field"><label>Целевой ДРР, %</label>
              <input type="number" step="1" min="0" max="60" data-target="drr" value="${+(t.drr * 100).toFixed(1)}"></div>
            <div class="field"><label>Мин. маржа, ₽</label>
              <input type="number" step="10" min="0" data-target="minMargin" value="${Math.round(t.minMargin)}"></div>
            <div class="field"><label>Округление</label>
              <select data-target="rounding">
                <option value="up_10" ${roundingMode() === 'up_10' ? 'selected' : ''}>до 10 ₽</option>
                <option value="up_50" ${roundingMode() === 'up_50' ? 'selected' : ''}>до 50 ₽</option>
                <option value="psych_99" ${roundingMode() === 'psych_99' ? 'selected' : ''}>до 99 ₽</option>
                <option value="up_1" ${roundingMode() === 'up_1' ? 'selected' : ''}>до 1 ₽</option>
                <option value="none" ${roundingMode() === 'none' ? 'selected' : ''}>без округления</option>
              </select></div>
          </div>
          ${delta != null && delta < 0 ? '<div class="hint-note">Текущая маржа уже выше цели — расчётная цена ниже текущей. Снижение допустимо.</div>' : ''}
          ${targetPrice ? `<button class="btn btn-primary" data-apply-price="${targetPrice}">Применить цену</button>` : ''}
        </div>
      </div>

      <div class="card panel">
        <h3>Акции и цены входа</h3>
        <div class="promo-list">${promosHtml}</div>
      </div>
    </div>`;

  bindCalcEvents(sku);
}

function bindCalcEvents(sku) {
  $$('#calcMain [data-param]').forEach(input => {
    input.addEventListener('change', () => {
      const key = input.dataset.param;
      let value = Number(String(input.value).replace(',', '.'));
      if (Number.isNaN(value) || value < 0) return;
      if (['return_rate', 'commission_rate', 'drr_rate'].includes(key)) value /= 100;
      if (!ui.edits[sku.offer_id]) ui.edits[sku.offer_id] = {};
      ui.edits[sku.offer_id][key] = value;
      saveUI();
      renderCalc();
      renderSkuList();
    });
  });

  $$('#calcMain [data-target]').forEach(input => {
    input.addEventListener('change', () => {
      const key = input.dataset.target;
      if (key === 'rounding') {
        ui.rounding = input.value;
      } else {
        const value = Number(String(input.value).replace(',', '.'));
        if (Number.isNaN(value)) return;
        if (key === 'minMargin') ui.targets.minMargin = value;
        else ui.targets[key] = value / 100;
      }
      saveUI();
      renderCalc();
    });
  });

  $$('#calcMain [data-price-factor]').forEach(btn => btn.addEventListener('click', () => {
    const factor = Number(btn.dataset.priceFactor);
    const eff = effectiveSku(sku);
    const next = Math.max(1, Math.round((Number(eff.current_price_rub) || 0) * (1 + factor)));
    setSkuEdit(sku, 'current_price_rub', next);
  }));

  $$('#calcMain [data-price-base]').forEach(btn => btn.addEventListener('click', () => {
    if (!ui.edits[sku.offer_id]) return;
    delete ui.edits[sku.offer_id].current_price_rub;
    saveUI();
    renderCalc();
    renderSkuList();
  }));

  $$('#calcMain [data-apply-price]').forEach(btn => btn.addEventListener('click', () => {
    setSkuEdit(sku, 'current_price_rub', Number(btn.dataset.applyPrice));
    toast(`Цена обновлена: ${$fmt.money(Number(btn.dataset.applyPrice))}`);
  }));

  $$('#calcMain [data-apply-promo]').forEach(btn => btn.addEventListener('click', () => {
    setSkuEdit(sku, 'current_price_rub', Number(btn.dataset.price));
    toast(`Цена акции применена: ${$fmt.money(Number(btn.dataset.price))}`);
  }));

  const resetSku = $('#btnResetSku');
  if (resetSku) resetSku.addEventListener('click', () => {
    delete ui.edits[sku.offer_id];
    saveUI();
    renderCalc();
    renderSkuList();
    toast('Правки по SKU сброшены');
  });

  const resetParams = $('#btnResetParams');
  if (resetParams) resetParams.addEventListener('click', () => {
    delete ui.edits[sku.offer_id];
    saveUI();
    renderCalc();
    renderSkuList();
    toast('Параметры возвращены к данным из таблицы');
  });
}

function setSkuEdit(sku, key, value) {
  if (!ui.edits[sku.offer_id]) ui.edits[sku.offer_id] = {};
  ui.edits[sku.offer_id][key] = value;
  saveUI();
  renderCalc();
  renderSkuList();
}

/* ------------------------------------------------------------ тарифы */
function renderTariffs() {
  const t = bundle.tariffs;
  const settings = bundle.settings;
  const schemeNames = Object.keys(t.schemes || {});
  const cards = [];

  cards.push(`
    <div class="card tariff-card full">
      <h3 style="margin:0 0 10px;font-size:11px;letter-spacing:.09em;color:var(--dim);text-transform:uppercase">Комиссии по категориям</h3>
      <table class="mini-table">
        <thead><tr><th>Схема</th><th>Категория</th><th>Ставка</th></tr></thead>
        <tbody>
          ${schemeNames.map(scheme => Object.entries(t.schemes[scheme].commission_by_category || {})
            .map(([cat, rate]) => {
              const name = (bundle.skus.find(s => String(s.category_id) === String(cat)) || {}).category_name || cat;
              return `<tr><td>${scheme}</td><td>${esc(name)}</td><td>${$fmt.pct(rate)}</td></tr>`;
            }).join('')).join('')}
          ${schemeNames.map(scheme => `<tr><td>${scheme}</td><td class="muted">по умолчанию</td><td>${$fmt.pct(t.schemes[scheme].default_commission)}</td></tr>`).join('')}
        </tbody>
      </table>
    </div>`);

  cards.push(`
    <div class="card tariff-card">
      <h3>Логистика по объёму</h3>
      ${schemeNames.map(scheme => `
        <div class="muted" style="font-size:10px;margin:6px 0 2px">${scheme}</div>
        <table class="mini-table"><tbody>
          ${(t.schemes[scheme].logistics_tiers || []).map(tier =>
            `<tr><td>${tier.volume_from_l}–${tier.volume_to_l == null ? '∞' : tier.volume_to_l} л</td><td>${$fmt.money(tier.price_rub)}</td></tr>`).join('')}
        </tbody></table>`).join('')}
    </div>`);

  const other = [];
  schemeNames.forEach(scheme => {
    const cfg = t.schemes[scheme];
    other.push(`<div class="muted" style="font-size:10px;margin:8px 0 2px">${scheme}</div>`);
    other.push(`<div class="kv"><span>Последняя миля</span><b>${$fmt.pct(cfg.last_mile.rate)} · мин ${$fmt.money(cfg.last_mile.min_rub)} · макс ${$fmt.money(cfg.last_mile.max_rub)}</b></div>`);
    other.push(`<div class="kv"><span>Эквайринг</span><b>${$fmt.pct(cfg.acquiring.rate)}${cfg.acquiring.included_in_commission ? ' (в комиссии)' : ''}</b></div>`);
    other.push(`<div class="kv"><span>Хранение</span><b>${$fmt.money2(cfg.storage.rate_per_liter_day)} / л в сутки</b></div>`);
    other.push(`<div class="kv"><span>Возврат</span><b>${$fmt.money(cfg.returns.fee_per_return_rub)}${cfg.returns.commission_refunded ? ' · комиссия возвращается' : ''}</b></div>`);
    other.push(`<div class="kv"><span>Утилизация</span><b>${$fmt.money(cfg.utilization.fee_per_unit_rub)} / ед.</b></div>`);
    other.push(`<div class="kv"><span>Штрафы</span><b>${$fmt.pct(cfg.fines.rate)} от выручки</b></div>`);
  });
  cards.push(`<div class="card tariff-card"><h3>Прочие условия</h3>${other.join('')}</div>`);

  cards.push(`
    <div class="card tariff-card">
      <h3>Налоги и цели</h3>
      <div class="kv"><span>Налоговый режим</span><b>${esc(settings.tax.regime)}</b></div>
      <div class="kv"><span>Ставка налога</span><b>${$fmt.pct(settings.tax.rate)}</b></div>
      <div class="kv"><span>НДС</span><b>${$fmt.pct(settings.tax.vat_rate)}${settings.tax.vat_in_price ? ' · в цене' : ''}</b></div>
      <div class="kv"><span>Целевая маржинальность</span><b>${$fmt.pct(ui.targets.margin)}</b></div>
      <div class="kv"><span>Целевой ДРР</span><b>${$fmt.pct(ui.targets.drr)}</b></div>
      <div class="kv"><span>Минимальная маржа</span><b>${$fmt.money(ui.targets.minMargin)}</b></div>
      <div class="kv"><span>Округление цены</span><b>${esc(roundingMode())}</b></div>
    </div>`);

  const meta = bundle.meta || {};
  const warnings = meta.warnings || [];

  $('#tariffsContent').innerHTML = `
    <div class="page-head" style="margin-bottom:14px">
      <div><h1>Тарифы и источник данных</h1>
      <p class="sub">Условия площадки, на которых считаются расходы</p></div>
    </div>

    <div class="card source-card">
      <div>
        <div class="title">${esc(meta.source || 'Демо-данные')}</div>
        <div class="meta">${meta.version ? 'Версия: ' + esc(meta.version) + ' · ' : ''}${(bundle.skus || []).length} SKU · ${(bundle.promos || []).length} акций · ${esc(meta.tariffs_source || '')}</div>
        ${warnings.length ? `<ul class="warn-list">${warnings.map(w => `<li><svg class="icon"><use href="#i-alert"/></svg>${esc(w)}</li>`).join('')}</ul>` : ''}
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-ghost" data-tpl><svg class="icon"><use href="#i-doc"/></svg><span>Шаблон</span></button>
        <button class="btn btn-outline" data-load><svg class="icon"><use href="#i-upload"/></svg><span>Загрузить данные</span></button>
      </div>
    </div>

    <div class="api-card">
      <div class="ic"><svg class="icon"><use href="#i-percent"/></svg></div>
      <div>
        <b>API площадки — следующий этап</b>
        <span>Первая версия работает локально на загруженных таблицах. Модель данных уже совпадает с полями API:
        комиссия, логистика, последняя миля, эквайринг, хранение, возвраты. Подключение заменит ручную загрузку без изменения расчётов.</span>
      </div>
    </div>

    <div class="tariff-grid">${cards.join('')}</div>`;

  $$('#tariffsContent [data-tpl]').forEach(b => b.addEventListener('click', downloadTemplate));
  $$('#tariffsContent [data-load]').forEach(b => b.addEventListener('click', openModal));
}

/* ------------------------------------------------------------ страницы */
function render() {
  if (!bundle) return;
  targets();

  $$('.tab').forEach(tab => tab.classList.toggle('active', tab.dataset.page === ui.page));
  $$('.page').forEach(page => page.classList.toggle('active', page.id === `page-${ui.page}`));

  const meta = bundle.meta || {};
  $('#sourceName').textContent = meta.source || 'Демо-данные';
  $('#sourceCount').textContent = `${(bundle.skus || []).length} SKU`;
  $('#sourceChip').title = `${meta.source || ''}${meta.version ? ' · ' + meta.version : ''}`;

  if (ui.page === 'overview') renderOverview();
  if (ui.page === 'calc') renderCalc();
  if (ui.page === 'tariffs') renderTariffs();
}

function switchPage(page) {
  ui.page = page;
  saveUI();
  render();
  $('.main-scroll').scrollTop = 0;
}

/* ------------------------------------------------------------ экспорт */
function buildExportPayload() {
  const t = targets();
  const rows = (bundle.skus || []).map(sku => {
    const eff = effectiveSku(sku);
    const price = Number(eff.current_price_rub) || 0;
    const drr = drrFor(sku);
    const current = calc(eff, price, drr);
    const targetPrice = solvePrice(eff, t.margin, t.drr);
    const targetMargin = targetPrice ? calc(eff, targetPrice, t.drr).margin : null;
    const status = statusOf(current.margin, current.rate);
    return {
      name: sku.name,
      category: sku.category_name,
      scheme: sku.scheme,
      price,
      total: current.total,
      margin: current.margin,
      marginRate: current.rate,
      targetPrice: targetPrice || null,
      targetMargin,
      status: status.label,
    };
  });

  const promos = [];
  (bundle.skus || []).forEach(sku => {
    const eff = effectiveSku(sku);
    const price = Number(eff.current_price_rub) || 0;
    promosFor(sku.offer_id).forEach(promo => {
      const pp = promoPrice(eff, promo, price);
      const pdrr = Number(promo.drr_rate || t.drr);
      const res = calc(eff, pp, pdrr);
      const be = breakEvenPrice(eff, pdrr, t.minMargin);
      const maxDiscount = be && price > 0 ? Math.max(0, 1 - be / price) : null;
      promos.push({
        name: sku.name,
        promo: promo.promo_name,
        price: pp,
        total: res.total,
        margin: res.margin,
        marginRate: res.rate,
        maxDiscount,
        verdict: statusOf(res.margin, res.rate).label,
      });
    });
  });

  return {
    meta: bundle.meta,
    targetMargin: +(t.margin * 100).toFixed(1),
    targetDrr: +(t.drr * 100).toFixed(1),
    rows,
    promos,
  };
}

async function exportReport() {
  if (!bundle) return;
  showLoading('Собираю отчёт…');
  try {
    const res = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(buildExportPayload()),
    });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || 'Ошибка экспорта');
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Юнит-экономика_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    toast('Отчёт сохранён в загрузки');
  } catch (err) {
    toast(err.message, 'err');
  } finally {
    hideLoading();
  }
}

/* ------------------------------------------------------------ данные */
function applyBundle(data) {
  bundle = data;
  targets();
  ui.search = '';
  ui.calcSearch = '';
  ui.scheme = 'all';
  ui.category = 'all';
  if ($('#ovSearch')) $('#ovSearch').value = '';
  if ($('#skuSearch')) $('#skuSearch').value = '';
  if (!bundle.skus.some(s => s.offer_id === ui.selected)) {
    ui.selected = bundle.skus.length ? bundle.skus[0].offer_id : null;
  }
  ui.edits = {};
  render();
}

async function loadData() {
  try {
    const res = await fetch('/api/data');
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    bundle = data;
    targets();
    if (!ui.selected) ui.selected = bundle.skus.length ? bundle.skus[0].offer_id : null;
    render();
  } catch (e) {
    toast('Не удалось загрузить данные. Перезапустите приложение.', 'err');
    $('#sourceName').textContent = 'Нет данных';
  }
}

async function uploadFile(file) {
  if (!file) return;
  if (!/\.(xlsx|csv)$/i.test(file.name)) {
    $('#modalError').textContent = 'Поддерживаются только .xlsx и .csv. Файл .xls сохраните как .xlsx.';
    $('#modalError').hidden = false;
    return;
  }
  $('#modalError').hidden = true;
  showLoading('Читаю таблицу…');
  try {
    const res = await fetch('/api/upload', {
      method: 'POST',
      headers: { 'X-Filename': encodeURIComponent(file.name) },
      body: file,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Не удалось прочитать файл');
    applyBundle(data);
    closeModal();
    toast(`Загружено ${data.skus.length} SKU`);
    (data.meta.warnings || []).slice(0, 3).forEach(w => toast(w));
  } catch (err) {
    $('#modalError').textContent = err.message;
    $('#modalError').hidden = false;
  } finally {
    hideLoading();
    $('#fileInput').value = '';
  }
}

async function resetData() {
  showLoading('Возвращаю демо-данные…');
  try {
    const res = await fetch('/api/reset', { method: 'POST' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    applyBundle(await res.json());
    closeModal();
    toast('Загружены демо-данные');
  } catch (e) {
    toast('Не удалось вернуть демо-данные', 'err');
  } finally {
    hideLoading();
  }
}

function downloadTemplate() {
  const link = document.createElement('a');
  link.href = '/api/template';
  link.download = 'Шаблон_данных_Маркет.xlsx';
  document.body.appendChild(link);
  link.click();
  link.remove();
}

/* ------------------------------------------------------------ UI */
function openModal() {
  $('#modalError').hidden = true;
  $('#loadModal').classList.add('show');
  $('#loadModal').setAttribute('aria-hidden', 'false');
}
function closeModal() {
  $('#loadModal').classList.remove('show');
  $('#loadModal').setAttribute('aria-hidden', 'true');
}
function showLoading(text) {
  $('#loadingText').textContent = text || 'Секунду…';
  $('#loading').hidden = false;
}
function hideLoading() { $('#loading').hidden = true; }

function toast(message, kind = 'ok') {
  const el = document.createElement('div');
  el.className = `toast ${kind}`;
  el.innerHTML = `<svg class="icon"><use href="#${kind === 'err' ? 'i-alert' : 'i-check'}"/></svg><span>${esc(message)}</span>`;
  $('#toasts').appendChild(el);
  setTimeout(() => {
    el.classList.add('out');
    setTimeout(() => el.remove(), 260);
  }, kind === 'err' ? 6000 : 3600);
}

function ripple(e) {
  const button = e.target.closest('.btn, .chip, .tab, .price-preset');
  if (!button) return;
  const rect = button.getBoundingClientRect();
  const size = Math.max(rect.width, rect.height);
  const span = document.createElement('span');
  span.className = 'ripple';
  span.style.width = span.style.height = size + 'px';
  span.style.left = e.clientX - rect.left - size / 2 + 'px';
  span.style.top = e.clientY - rect.top - size / 2 + 'px';
  button.appendChild(span);
  setTimeout(() => span.remove(), 600);
}

function bind() {
  $$('.tab').forEach(tab => tab.addEventListener('click', () => switchPage(tab.dataset.page)));

  $('#ovSearch').addEventListener('input', e => { ui.search = e.target.value; saveUI(); renderOverview(); });
  $('#skuSearch').addEventListener('input', e => { ui.calcSearch = e.target.value; saveUI(); renderSkuList(); });

  $('#btnLoad').addEventListener('click', openModal);
  $('#modalClose').addEventListener('click', closeModal);
  $('#loadModal').addEventListener('click', e => { if (e.target === $('#loadModal')) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

  $('#dropzone').addEventListener('click', () => $('#fileInput').click());
  $('#fileInput').addEventListener('change', e => uploadFile(e.target.files[0]));
  ['dragenter', 'dragover'].forEach(evt => $('#dropzone').addEventListener(evt, e => {
    e.preventDefault();
    $('#dropzone').classList.add('drag');
  }));
  ['dragleave', 'drop'].forEach(evt => $('#dropzone').addEventListener(evt, e => {
    e.preventDefault();
    $('#dropzone').classList.remove('drag');
  }));
  $('#dropzone').addEventListener('drop', e => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) uploadFile(file);
  });

  $('#btnTemplate').addEventListener('click', downloadTemplate);
  $('#btnTemplate2').addEventListener('click', downloadTemplate);
  $('#btnResetData').addEventListener('click', resetData);
  $('#btnExport').addEventListener('click', exportReport);

  document.addEventListener('pointerdown', ripple);
}

async function init() {
  bind();
  await loadData();
  setInterval(() => fetch('/api/ping').catch(() => {}), 4000);
}

document.addEventListener('DOMContentLoaded', init);
