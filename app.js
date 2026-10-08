'use strict';

// ===================== НАСТРОЙКИ =====================

const CONFIG = {
  // Excel-файл с данными (лежит рядом с index.html в репозитории)
  file: 'Свод договоренностей.xlsx',
  // Названия листов. Если лист не найден по имени, берётся 1-й и 2-й лист.
  fieldsSheet: 'Поля',
  regionsSheet: 'Регионы',
  // Сколько дней запись считается «новой»
  newDays: 14,
  // Иконки разделов листа «Поля» (по началу заголовка столбца)
  fieldIcons: [
    [/групп|тем/i, '📂'],
    [/итог/i, '✅'],
    [/объект|адрес/i, '📍'],
    [/источник/i, '🔗'],
    [/причин/i, '🔒'],
    [/ручн|публикац/i, '✍️'],
  ],
  // Разные написания одного региона в файле → одно название на сайте
  regionAliases: {
    'ростов': 'Ростовская область',
    'башкирия': 'Республика Башкортостан',
    'башкортостан': 'Республика Башкортостан',
    'якутия': 'Республика Саха (Якутия)',
  },
  // Дополнительные слова, по которым регион ищется в общих правилах раздела «Поля»
  regionMentions: {
    'Республика Башкортостан': ['башкир'],
    'Республика Саха (Якутия)': ['якут'],
  },
};

const GLOBAL = 'Все регионы';

// Быстрые фильтры раздела «Регионы» (поиск по ключевым словам в тексте)
const TOPICS = [
  { id: 'fake', label: 'Фейк', re: /фейк|фэйк/ },
  { id: 'sc', label: 'Закрытие по СЦ', re: /(^|[^а-яё])сц([^а-яё]|$)/ },
  { id: 'common', label: 'Общий ответ', re: /общ(ий|им|ие|ими|его|ему)\s+ответ|един(ым|ый)\s+ответ|одним\s+ответом|перекличк/ },
  { id: 'screen', label: 'Скрин / ручная публикация', re: /скрин|ручн/ },
  { id: 'addr', label: 'Адрес / объект', re: /адрес|объект/ },
  { id: 'theme', label: 'Тема', re: /(^|[^а-яё])(тем[аеуыо]|тематик)/ },
  { id: 'result', label: 'Итог / тип', re: /итог|(^|[^а-яё])тип/ },
  { id: 'max', label: 'MAX', re: /(^|[^а-яёa-z])(макс[ае]?|max|мах)([^а-яёa-z]|$)/ },
  { id: 'vdl', label: 'ВДЛ / эфир', re: /вдл|эфир/ },
];

// ===================== СОСТОЯНИЕ =====================

const state = {
  data: null,         // { fields, entries, regions, lastDate }
  siteData: null,     // данные с сайта (чтобы вернуться после предпросмотра)
  preview: null,      // имя файла в режиме предпросмотра
  sourceInfo: '',
  tab: 'fields',
  q: '',
  region: '',         // '' — все записи
  topic: '',
  onlyNew: false,
  sort: 'new',
  col: '',            // '' — все разделы «Поля»
  target: '',         // id записи, к которой надо прокрутить
  expanded: new Set(),
};

const $ = (id) => document.getElementById(id);

// ===================== РАЗБОР EXCEL =====================

const norm = (s) => String(s ?? '').toLowerCase().replace(/ё/g, 'е');
const clean = (v) => (v === null || v === undefined) ? '' : String(v).replace(/\r\n?/g, '\n').trim();

function hashId(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

function findSheet(wb, name, fallbackIndex) {
  const key = norm(name).trim();
  const found = wb.SheetNames.find((n) => norm(n).trim() === key);
  const sheetName = found || wb.SheetNames[fallbackIndex];
  return sheetName ? { name: sheetName, rows: XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null, blankrows: false }), exact: !!found } : null;
}

function parseDate(v) {
  if (v === null || v === undefined || v === '') return null;
  let y, m, d;
  if (typeof v === 'number') {
    const dt = new Date(Math.round((v - 25569) * 864e5));
    y = dt.getUTCFullYear(); m = dt.getUTCMonth() + 1; d = dt.getUTCDate();
  } else if (v instanceof Date) {
    y = v.getFullYear(); m = v.getMonth() + 1; d = v.getDate();
  } else {
    const s = String(v).trim();
    let r = s.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
    if (r) { d = +r[1]; m = +r[2]; y = +r[3]; if (y < 100) y += 2000; }
    else if ((r = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/))) { y = +r[1]; m = +r[2]; d = +r[3]; }
    else return { iso: null, text: s };
  }
  const p = (n) => String(n).padStart(2, '0');
  return { iso: `${y}-${p(m)}-${p(d)}`, text: `${p(d)}.${p(m)}.${y}` };
}

function cleanRegion(raw) {
  const s = clean(raw).replace(/\s+/g, ' ').replace(/[\s!.]+$/, '').trim();
  const key = norm(s);
  if (!key) return '';
  if (key === 'все регионы' || key === 'все') return GLOBAL;
  return CONFIG.regionAliases[key] || s;
}

function parseWorkbook(wb) {
  const warnings = [];

  // ----- Лист «Поля»: каждый столбец — отдельный раздел правил -----
  const fs = findSheet(wb, CONFIG.fieldsSheet, 0);
  const fields = [];
  if (!fs) warnings.push('Не найден лист «Поля».');
  else {
    if (!fs.exact) warnings.push(`Лист «${CONFIG.fieldsSheet}» не найден по имени, взят лист «${fs.name}».`);
    const headIdx = fs.rows.findIndex((r) => r && r.filter((c) => clean(c)).length >= 2);
    if (headIdx < 0) warnings.push('На листе «Поля» не найдена строка заголовков.');
    else {
      const head = fs.rows[headIdx];
      head.forEach((h, ci) => {
        const title = clean(h);
        if (!title) return;
        const icon = (CONFIG.fieldIcons.find(([re]) => re.test(title)) || [null, '📌'])[1];
        const items = [];
        for (let ri = headIdx + 1; ri < fs.rows.length; ri++) {
          const text = clean(fs.rows[ri] && fs.rows[ri][ci]);
          if (text) items.push({ id: 'f' + hashId(title + text), text, search: norm(text) });
        }
        fields.push({ title, icon, items });
      });
    }
  }

  // ----- Лист «Регионы»: дата | регион | особенность -----
  const rs = findSheet(wb, CONFIG.regionsSheet, 1);
  const entries = [];
  if (!rs) warnings.push('Не найден лист «Регионы».');
  else {
    if (!rs.exact) warnings.push(`Лист «${CONFIG.regionsSheet}» не найден по имени, взят лист «${rs.name}».`);
    let headIdx = rs.rows.findIndex((r) => r && r.some((c) => /регион/i.test(clean(c))));
    let cDate = 0, cRegion = 1, cText = 2;
    if (headIdx < 0) { headIdx = 0; warnings.push('На листе «Регионы» не найден заголовок «Регион», использую столбцы A, B, C.'); }
    else {
      const head = rs.rows[headIdx].map(clean);
      cRegion = head.findIndex((h) => /регион/i.test(h));
      const d = head.findIndex((h) => /дата/i.test(h));
      if (d >= 0) cDate = d;
      let t = head.findIndex((h) => /особенн|описан|суть|текст/i.test(h));
      if (t < 0) t = head.findIndex((h, i) => h && i !== cRegion && i !== cDate);
      if (t >= 0) cText = t;
    }
    let prevRegion = '';
    const seen = new Map();
    for (let ri = headIdx + 1; ri < rs.rows.length; ri++) {
      const row = rs.rows[ri] || [];
      const text = clean(row[cText]);
      let region = cleanRegion(row[cRegion]);
      if (!text) continue;
      if (!region) region = prevRegion; // объединённые ячейки: регион из строки выше
      if (!region) { region = 'Без региона'; warnings.push(`Запись без региона: «${text.slice(0, 60)}…»`); }
      prevRegion = region;
      const date = parseDate(row[cDate]);
      let id = 'e' + hashId(region + '|' + text);
      if (seen.has(id)) id += '-' + seen.get(id);
      seen.set(id, (seen.get(id) || 0) + 1);
      entries.push({ id, order: entries.length, region, date, text });
    }
  }

  // Объединяем написания региона, отличающиеся только регистром/пробелами
  const variants = new Map();
  entries.forEach((e) => {
    const k = norm(e.region);
    if (!variants.has(k)) variants.set(k, new Map());
    const m = variants.get(k);
    m.set(e.region, (m.get(e.region) || 0) + 1);
  });
  const display = new Map();
  variants.forEach((m, k) => {
    const list = [...m.entries()].sort((a, b) => b[1] - a[1]);
    const nice = list.find(([n]) => n !== n.toUpperCase() || n.length <= 5) || list[0];
    display.set(k, nice[0]);
  });

  const today = new Date();
  const newLimit = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) - CONFIG.newDays * 864e5)
    .toISOString().slice(0, 10);
  let lastDate = null;

  entries.forEach((e) => {
    e.region = display.get(norm(e.region));
    e.isGlobal = e.region === GLOBAL;
    e.isNew = !!(e.date && e.date.iso && e.date.iso >= newLimit);
    e.search = norm(e.region + ' ' + e.text + ' ' + (e.date ? e.date.text : ''));
    e.topics = TOPICS.filter((t) => t.re.test(norm(e.text))).map((t) => t.id);
    if (e.date && e.date.iso && (!lastDate || e.date.iso > lastDate.iso)) lastDate = e.date;
  });

  const regions = [...new Set(entries.map((e) => e.region))]
    .filter((r) => r !== GLOBAL)
    .sort((a, b) => a.localeCompare(b, 'ru'));

  return { fields, entries, regions, lastDate, warnings };
}

// ===================== ВСПОМОГАТЕЛЬНОЕ =====================

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function plural(n, one, few, many) {
  const m10 = n % 10, m100 = n % 100;
  if (m100 >= 11 && m100 <= 19) return many;
  if (m10 === 1) return one;
  if (m10 >= 2 && m10 <= 4) return few;
  return many;
}

function terms() {
  return norm(state.q).split(/\s+/).filter(Boolean);
}

function matches(searchText, ts) {
  return ts.every((t) => searchText.includes(t));
}

function highlighter(ts) {
  if (!ts.length) return null;
  const parts = [...ts].sort((a, b) => b.length - a.length)
    .map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/е/g, '[её]'));
  return new RegExp('(' + parts.join('|') + ')', 'gi');
}

function hl(text, re) {
  if (!re) return esc(text);
  return text.split(re).map((p, i) => (i % 2 ? `<mark>${esc(p)}</mark>` : esc(p))).join('');
}

const URL_RE = /(https?:\/\/[^\s<>"«»]+)/g;
const INC_RE = /(\b\d{5,8}\b)/g;

// Текст записи → HTML: ссылки, номера инцидентов (копируются по клику), подсветка поиска
function richText(text, re) {
  return text.split(URL_RE).map((part, i) => {
    if (i % 2) {
      let url = part, tail = '';
      const m = url.match(/[.,;:)\]]+$/);
      if (m) { tail = m[0]; url = url.slice(0, -tail.length); }
      let shown = url.replace(/^https?:\/\//, '');
      if (shown.length > 60) shown = shown.slice(0, 57) + '…';
      return `<a href="${esc(url)}" target="_blank" rel="noopener noreferrer">${hl(shown, re)}</a>${hl(tail, re)}`;
    }
    return part.split(INC_RE).map((p, j) =>
      j % 2 ? `<span class="inc" data-copy="${p}" title="Скопировать номер">${hl(p, re)}</span>` : hl(p, re)
    ).join('');
  }).join('');
}

function regionMatcher(region) {
  if (!region || region === GLOBAL) return null;
  const generic = /^(область|обл|край|республика|респ|автономный|автономная|округ|ао|город|г|санкт)$/i;
  const words = region.replace(/[()]/g, ' ').split(/[\s-]+/).filter((w) => w && !generic.test(w));
  const parts = [];
  words.forEach((w) => {
    if (w.length <= 4 && w === w.toUpperCase()) parts.push(`(^|[^а-яёa-z])${norm(w)}([^а-яёa-z]|$)`);
    else if (w.length >= 4) parts.push(`(^|[^а-яё])${norm(w).slice(0, Math.max(4, w.length - 3))}`);
  });
  (CONFIG.regionMentions[region] || []).forEach((a) => parts.push(`(^|[^а-яё])${norm(a)}`));
  return parts.length ? new RegExp(parts.join('|')) : null;
}

function fmtCount(n, one, few, many) { return `${n} ${plural(n, one, few, many)}`; }

function toast(msg) {
  const t = $('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.remove('show'), 1600);
}

async function copy(text, msg) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const ta = document.createElement('textarea');
    ta.value = text; document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); } finally { ta.remove(); }
  }
  toast(msg || 'Скопировано');
}

// ===================== АДРЕС СТРАНИЦЫ (#...) =====================

function readHash() {
  const p = new URLSearchParams(location.hash.slice(1));
  state.tab = p.get('tab') === 'regions' ? 'regions' : 'fields';
  state.q = p.get('q') || '';
  state.region = p.get('region') || '';
  state.topic = p.get('topic') || '';
  state.onlyNew = p.get('new') === '1';
  state.sort = p.get('sort') === 'old' ? 'old' : 'new';
  state.col = p.get('col') || '';
  state.target = p.get('id') || '';
  if (state.target) state.tab = state.target.startsWith('f') ? 'fields' : 'regions';
}

function writeHash() {
  const p = new URLSearchParams();
  if (state.tab !== 'fields') p.set('tab', state.tab);
  if (state.q) p.set('q', state.q);
  if (state.region) p.set('region', state.region);
  if (state.topic) p.set('topic', state.topic);
  if (state.onlyNew) p.set('new', '1');
  if (state.sort !== 'new') p.set('sort', state.sort);
  if (state.col) p.set('col', state.col);
  const h = p.toString();
  history.replaceState(null, '', h ? '#' + h : location.pathname + location.search);
}

function linkTo(id) {
  return location.origin + location.pathname + '#id=' + id;
}

// ===================== ОТРИСОВКА =====================

function render() {
  if (!state.data) return;
  writeHash();
  const ts = terms();
  const re = highlighter(ts);

  document.querySelectorAll('.tab').forEach((b) => b.classList.toggle('active', b.dataset.tab === state.tab));
  $('view-fields').classList.toggle('active', state.tab === 'fields');
  $('view-regions').classList.toggle('active', state.tab === 'regions');
  $('search-box').classList.toggle('has-value', !!state.q);
  if ($('search').value !== state.q) $('search').value = state.q;

  // Счётчики на вкладках
  const d = state.data;
  const fieldsTotal = d.fields.reduce((s, c) => s + c.items.length, 0);
  const fieldsHit = d.fields.reduce((s, c) => s + c.items.filter((it) => matches(it.search, ts)).length, 0);
  const regionsHit = d.entries.filter((e) => matches(e.search, ts)).length;
  $('count-fields').textContent = ts.length ? fieldsHit : fieldsTotal;
  $('count-regions').textContent = ts.length ? regionsHit : d.entries.length;

  renderFields(ts, re);
  renderRegions(ts, re);

  if (state.target) {
    const el = document.getElementById(state.target);
    state.target = '';
    if (el) requestAnimationFrame(() => {
      el.scrollIntoView({ block: 'center' });
      el.classList.add('target');
      setTimeout(() => el.classList.remove('target'), 2500);
    });
  }
}

function renderFields(ts, re) {
  const d = state.data;
  const colIdx = state.col === '' ? -1 : +state.col;

  $('fields-toolbar').innerHTML =
    `<button class="chip ${colIdx < 0 ? 'active' : ''}" data-col="">Все разделы</button>` +
    d.fields.map((c, i) => {
      const n = c.items.filter((it) => matches(it.search, ts)).length;
      return `<button class="chip ${colIdx === i ? 'active' : ''}" data-col="${i}">${c.icon} ${esc(c.title)}<span class="n">${n}</span></button>`;
    }).join('');

  const cards = d.fields.map((c, i) => {
    if (colIdx >= 0 && colIdx !== i) return '';
    const items = c.items.filter((it) => matches(it.search, ts));
    if (!items.length && ts.length) return '';
    return `<article class="field-card">
      <h2><span>${c.icon}</span>${esc(c.title)}<span class="n">${items.length}</span></h2>
      <ol>${items.length ? items.map((it) => `<li id="${it.id}">${richText(it.text, re)}<button class="icon-btn" data-copy-text="${it.id}" title="Скопировать текст">⧉</button></li>`).join('')
        : '<li style="color:var(--muted)">Пока нет правил</li>'}</ol>
    </article>`;
  }).join('');

  $('fields-content').innerHTML = cards.trim()
    ? `<div class="fields-grid">${cards}</div>`
    : `<div class="empty">В разделе «Поля» ничего не найдено по запросу «${esc(state.q)}».${regionsHitHint(ts)}</div>`;
}

function regionsHitHint(ts) {
  const n = state.data.entries.filter((e) => matches(e.search, ts)).length;
  return n ? `<br><br><button class="link-btn" data-goto="regions">В разделе «Регионы» ${plural(n, 'найдена', 'найдено', 'найдено')} ${fmtCount(n, 'запись', 'записи', 'записей')} →</button>` : '';
}

function sortEntries(list) {
  const dir = state.sort === 'old' ? 1 : -1;
  return [...list].sort((a, b) => {
    const ad = a.date && a.date.iso, bd = b.date && b.date.iso;
    if (ad && bd && ad !== bd) return ad < bd ? -dir : dir;
    if (ad && !bd) return dir < 0 ? -1 : 1;  // записи без даты — в конце (или в начале при «сначала старые»)
    if (!ad && bd) return dir < 0 ? 1 : -1;
    return (a.order - b.order) * (dir < 0 ? -1 : 1);
  });
}

function entryCard(e, re) {
  const long = !state.q && !state.expanded.has(e.id) && (e.text.length > 550 || e.text.split('\n').length > 8);
  return `<article class="card ${e.isGlobal ? 'global' : ''}" id="${e.id}">
    <div class="card-head">
      <button class="card-region" data-region="${esc(e.region)}" title="Показать все записи региона">${e.isGlobal ? '🌍 ' : ''}${hl(e.region, re)}</button>
      ${e.date ? `<span class="badge">📅 ${esc(e.date.text)}</span>` : '<span class="badge">без даты</span>'}
      ${e.isNew ? '<span class="badge new">Новое</span>' : ''}
      <span class="card-actions">
        <button class="icon-btn" data-link="${e.id}" title="Скопировать ссылку на запись">🔗</button>
        <button class="icon-btn" data-copy-text="${e.id}" title="Скопировать текст">⧉</button>
      </span>
    </div>
    <div class="card-text ${long ? 'clamped' : ''}">${richText(e.text, re)}</div>
    ${long ? `<button class="more-btn" data-expand="${e.id}">Показать полностью ↓</button>` : ''}
  </article>`;
}

function renderRegions(ts, re) {
  const d = state.data;
  const base = d.entries.filter((e) =>
    matches(e.search, ts) &&
    (!state.topic || e.topics.includes(state.topic)) &&
    (!state.onlyNew || e.isNew));

  // ----- список регионов -----
  const counts = new Map();
  base.forEach((e) => counts.set(e.region, (counts.get(e.region) || 0) + 1));
  const rq = norm($('region-search').value.trim());
  const item = (value, label, n, cls = '') =>
    `<li><button class="${cls} ${state.region === value ? 'active' : ''}" data-region="${esc(value)}" ${n === 0 && state.region !== value ? 'style="opacity:.45"' : ''}>${label}<span class="n">${n}</span></button></li>`;
  let listHtml = '';
  if (!rq) {
    listHtml += item('', 'Все записи', base.length);
    listHtml += item(GLOBAL, '🌍 Для всех регионов', counts.get(GLOBAL) || 0, 'global-item');
    listHtml += '<li class="sep"></li>';
  }
  const shownRegions = d.regions.filter((r) => !rq || norm(r).includes(rq));
  listHtml += shownRegions.map((r) => item(r, esc(r), counts.get(r) || 0)).join('');
  if (rq && !shownRegions.length) listHtml += '<li style="padding:8px 10px;color:var(--muted)">Регион не найден</li>';
  $('region-list').innerHTML = listHtml;
  if (state.region !== renderRegions.lastRegion) {
    renderRegions.lastRegion = state.region;
    const act = $('region-list').querySelector('button.active');
    if (act) act.scrollIntoView({ block: 'nearest' });
  }

  const sel = $('region-select');
  sel.innerHTML =
    `<option value="">Все записи (${base.length})</option>` +
    `<option value="${GLOBAL}">🌍 Для всех регионов (${counts.get(GLOBAL) || 0})</option>` +
    d.regions.map((r) => `<option value="${esc(r)}">${esc(r)} (${counts.get(r) || 0})</option>`).join('');
  sel.value = state.region;

  // ----- панель фильтров -----
  const newCount = d.entries.filter((e) => e.isNew && matches(e.search, ts) && (!state.topic || e.topics.includes(state.topic))).length;
  const topicCount = (id) => d.entries.filter((e) => e.topics.includes(id) && matches(e.search, ts) && (!state.onlyNew || e.isNew)).length;
  $('regions-toolbar').innerHTML =
    `<button class="chip new-chip ${state.onlyNew ? 'active' : ''}" data-new="1" title="Записи за последние ${CONFIG.newDays} дней">🆕 Новые<span class="n">${newCount}</span></button>` +
    TOPICS.map((t) => `<button class="chip ${state.topic === t.id ? 'active' : ''}" data-topic="${t.id}" title="Записи, где встречается: ${esc(t.label.toLowerCase())}">${esc(t.label)}<span class="n">${topicCount(t.id)}</span></button>`).join('') +
    `<span class="spacer"></span>
     <select class="control" id="sort" aria-label="Сортировка">
       <option value="new" ${state.sort === 'new' ? 'selected' : ''}>Сначала новые</option>
       <option value="old" ${state.sort === 'old' ? 'selected' : ''}>Сначала старые</option>
     </select>`;

  // ----- записи -----
  const filtersOn = ts.length || state.topic || state.onlyNew;
  let html = '';
  let info = '';

  if (!state.region) {
    const list = sortEntries(base);
    info = filtersOn ? `Найдено: ${fmtCount(list.length, 'запись', 'записи', 'записей')}` : `Всего ${fmtCount(list.length, 'запись', 'записи', 'записей')} по ${fmtCount(d.regions.length, 'региону', 'регионам', 'регионам')}`;
    html = list.length ? `<div class="cards">${list.map((e) => entryCard(e, re)).join('')}</div>` : '';
  } else if (state.region === GLOBAL) {
    const list = sortEntries(base.filter((e) => e.isGlobal));
    info = `Договорённости, действующие для всех регионов: ${list.length}`;
    html = list.length ? `<div class="cards">${list.map((e) => entryCard(e, re)).join('')}</div>` : '';
  } else {
    const own = sortEntries(base.filter((e) => e.region === state.region));
    const glob = sortEntries(base.filter((e) => e.isGlobal));
    const mre = regionMatcher(state.region);
    const mentions = [];
    if (mre) d.fields.forEach((c) => c.items.forEach((it) => {
      if (mre.test(it.search) && matches(it.search, ts)) mentions.push({ c, it });
    }));
    info = `${state.region}: ${fmtCount(own.length, 'договорённость', 'договорённости', 'договорённостей')}` +
      (glob.length ? ` + ${glob.length} для всех регионов` : '') +
      (mentions.length ? ` + ${fmtCount(mentions.length, 'упоминание', 'упоминания', 'упоминаний')} в разделе «Поля»` : '');
    if (own.length) html += `<h2 class="section-title">${esc(state.region)}</h2><div class="cards">${own.map((e) => entryCard(e, re)).join('')}</div>`;
    else html += `<h2 class="section-title">${esc(state.region)}</h2><div class="empty">${filtersOn ? 'Нет записей по выбранным фильтрам' : 'Для региона нет отдельных договорённостей'}</div>`;
    if (mentions.length) html += `<h2 class="section-title">Упоминания в разделе «Поля»</h2><div class="cards">${mentions.map(({ c, it }) =>
      `<article class="card" id="m-${it.id}"><div class="card-head"><span class="badge">${c.icon} ${esc(c.title)}</span>
        <span class="card-actions"><button class="icon-btn" data-link="${it.id}" title="Скопировать ссылку на правило">🔗</button></span></div>
        <div class="card-text">${richText(it.text, re)}</div></article>`).join('')}</div>`;
    if (glob.length) html += `<h2 class="section-title">🌍 Действуют для всех регионов</h2><div class="cards">${glob.map((e) => entryCard(e, re)).join('')}</div>`;
  }

  $('regions-info').textContent = info;
  $('regions-content').innerHTML = html || `<div class="empty">Ничего не найдено.${filtersOn ? '<br><br><button class="link-btn" data-reset="1">Сбросить фильтры</button>' : ''}</div>`;
}

function renderMeta() {
  const d = state.data;
  const parts = [
    fmtCount(d.entries.length, 'договорённость', 'договорённости', 'договорённостей'),
    fmtCount(d.regions.length, 'регион', 'региона', 'регионов'),
  ];
  if (d.lastDate) parts.push('последняя запись ' + d.lastDate.text);
  $('meta').textContent = parts.join(' · ');
}

// ===================== ЗАГРУЗКА =====================

function setBanner(html, cls) {
  $('banner-slot').innerHTML = html ? `<div class="banner ${cls || ''}">${html}</div>` : '';
}

function useData(data) {
  state.data = data;
  $('loading').style.display = 'none';
  renderMeta();
  render();
}

async function loadSite() {
  try {
    if (typeof XLSX === 'undefined') throw new Error('не загрузилась библиотека чтения Excel (vendor/xlsx.mini.min.js)');
    const res = await fetch(encodeURI(CONFIG.file), { cache: 'no-cache' });
    if (!res.ok) throw new Error(`файл «${CONFIG.file}» не найден (ошибка ${res.status})`);
    const buf = await res.arrayBuffer();
    const data = parseWorkbook(XLSX.read(buf, { type: 'array' }));
    state.siteData = data;
    const lm = res.headers.get('Last-Modified');
    state.sourceInfo = `Источник: ${CONFIG.file}` +
      (lm ? ` · опубликован ${new Date(lm).toLocaleString('ru-RU', { dateStyle: 'long', timeStyle: 'short' })}` : '');
    $('source-info').textContent = state.sourceInfo;
    useData(data);
  } catch (err) {
    $('loading').style.display = 'none';
    const local = location.protocol === 'file:';
    setBanner(`<span class="grow"><b>Не удалось загрузить данные:</b> ${esc(err.message)}.` +
      (local ? ' Страница открыта прямо с диска, а браузер в этом случае не даёт читать файлы. Выберите Excel-файл вручную.' : '') +
      `</span><button class="btn" data-pick="1">Выбрать файл…</button>`, 'error');
  }
}

async function loadPreview(file) {
  try {
    const buf = await file.arrayBuffer();
    const data = parseWorkbook(XLSX.read(buf, { type: 'array' }));
    state.preview = file.name;
    const warn = data.warnings.length
      ? `<br><b>Замечания:</b><ul style="margin:4px 0 0 18px;padding:0">${data.warnings.slice(0, 8).map((w) => `<li>${esc(w)}</li>`).join('')}</ul>` : '';
    const nameWarn = file.name !== CONFIG.file
      ? `<br>⚠️ Файл называется «${esc(file.name)}». Перед загрузкой на GitHub переименуйте его в <b>${esc(CONFIG.file)}</b>.` : '';
    setBanner(`<span class="grow"><b>Предпросмотр файла «${esc(file.name)}»</b>: ${fmtCount(data.entries.length, 'запись', 'записи', 'записей')} в «Регионах», ` +
      `${fmtCount(data.fields.reduce((s, c) => s + c.items.length, 0), 'правило', 'правила', 'правил')} в «Полях». ` +
      `Эти данные видны только вам, на сайте пока прежняя версия.${nameWarn}${warn}</span>` +
      (state.siteData ? '<button class="btn ghost" data-exit-preview="1">Вернуть данные сайта</button>' : ''), 'preview');
    $('source-info').textContent = `Предпросмотр: ${file.name}`;
    useData(data);
  } catch (err) {
    toast('Не удалось прочитать файл: ' + err.message);
  }
}

// ===================== СОБЫТИЯ =====================

// Если на текущей вкладке ничего не найдено, а на другой есть, переключаемся на неё
function autoSwitchTab() {
  const ts = terms();
  if (!ts.length || !state.data) return;
  const f = state.data.fields.some((c) => c.items.some((it) => matches(it.search, ts)));
  const r = state.data.entries.some((e) => matches(e.search, ts));
  if (state.tab === 'fields' && !f && r) state.tab = 'regions';
  else if (state.tab === 'regions' && !r && f) state.tab = 'fields';
}

function setHeaderHeight() {
  document.documentElement.style.setProperty('--header-h', document.querySelector('.header').offsetHeight + 'px');
}

function bind() {
  let timer;
  $('search').addEventListener('input', (e) => {
    clearTimeout(timer);
    timer = setTimeout(() => { state.q = e.target.value; autoSwitchTab(); render(); }, 120);
  });
  $('search').addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { state.q = ''; render(); }
  });
  $('search-clear').addEventListener('click', () => { state.q = ''; render(); $('search').focus(); });

  document.addEventListener('keydown', (e) => {
    if (e.key === '/' && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) {
      e.preventDefault(); $('search').focus(); $('search').select();
    }
  });

  $('region-search').addEventListener('input', () => render());
  $('region-select').addEventListener('change', (e) => { state.region = e.target.value; render(); window.scrollTo({ top: 0 }); });

  $('help-open').addEventListener('click', () => $('help').showModal());
  $('help-close').addEventListener('click', () => $('help').close());

  $('preview-open').addEventListener('click', () => $('preview-file').click());
  $('preview-file').addEventListener('change', (e) => {
    const f = e.target.files[0];
    if (f) loadPreview(f);
    e.target.value = '';
  });

  window.addEventListener('resize', setHeaderHeight);
  window.addEventListener('hashchange', () => { readHash(); render(); });

  document.addEventListener('change', (e) => {
    if (e.target.id === 'sort') { state.sort = e.target.value; render(); }
  });

  document.addEventListener('click', (e) => {
    const b = e.target.closest('button, .inc');
    if (!b) return;
    const ds = b.dataset;
    if (b.classList.contains('tab')) { state.tab = ds.tab; render(); return; }
    if ('goto' in ds) { state.tab = ds.goto; render(); return; }
    if ('col' in ds) { state.col = ds.col; render(); return; }
    if ('region' in ds) {
      state.region = ds.region; state.tab = 'regions';
      $('region-search').value = '';
      render(); window.scrollTo({ top: 0 }); return;
    }
    if ('topic' in ds) { state.topic = state.topic === ds.topic ? '' : ds.topic; render(); return; }
    if ('new' in ds) { state.onlyNew = !state.onlyNew; render(); return; }
    if ('reset' in ds) { state.q = ''; state.topic = ''; state.onlyNew = false; render(); return; }
    if ('expand' in ds) { state.expanded.add(ds.expand); render(); return; }
    if ('copy' in ds) { copy(ds.copy, `Номер ${ds.copy} скопирован`); return; }
    if ('copyText' in ds) {
      const id = ds.copyText;
      const e1 = state.data.entries.find((x) => x.id === id);
      const f1 = state.data.fields.flatMap((c) => c.items).find((x) => x.id === id);
      const text = e1 ? `${e1.region}${e1.date ? ' (' + e1.date.text + ')' : ''}: ${e1.text}` : f1 ? f1.text : '';
      copy(text, 'Текст скопирован');
      return;
    }
    if ('link' in ds) { copy(linkTo(ds.link), 'Ссылка на запись скопирована'); return; }
    if ('pick' in ds) { $('preview-file').click(); return; }
    if ('exitPreview' in ds) {
      state.preview = null; setBanner('');
      $('source-info').textContent = state.sourceInfo;
      useData(state.siteData); return;
    }
  });
}

// ===================== СТАРТ =====================

readHash();
bind();
setHeaderHeight();
loadSite();
