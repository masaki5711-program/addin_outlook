import { render, extractInputs, textToHtml } from './lib/render.js';
import { buildSalutation } from './lib/salutation.js';
import { parseRecipients, invalidAddresses, missingRecipients, formatRecipients, toOfficeRecipients } from './lib/recipients.js';
import * as store from './lib/store.js';

const CHIPS = [
  ['宛名', '{宛名}'],
  ['自分の名前', '{自分}'],
  ['件名', '{件名}'],
  ['今日', '{今日}'],
  ['明日', '{明日}'],
  ['3営業日後', '{3営業日後}'],
  ['1週間後', '{今日+7}'],
  ['来週月曜', '{来週月曜}'],
  ['入力欄', '{?項目名}'],
];

const $ = (id) => document.getElementById(id);

const state = {
  item: null,
  isRead: false,
  templates: [],
  settings: null,
  ctx: {},
  to: [], // 作成中メールの今の To / CC
  cc: [],
  current: null, // 確認画面で開いているテンプレート
  editing: null, // 編集中のテンプレート id（新規は null）
  lastField: null, // 差し込みボタンの挿入先
};

Office.onReady(async () => {
  state.item = Office.context.mailbox.item;
  state.isRead = typeof state.item.displayReplyForm === 'function';
  state.templates = store.loadTemplates();
  state.settings = store.loadSettings();
  bindEvents();
  await refreshContext();
  $('loading').hidden = true;
  showList();
});

// Office.js のコールバック API を Promise にする
function call(fn) {
  return new Promise((resolve, reject) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))),
  );
}

async function refreshContext() {
  const { item, isRead, settings } = state;
  const myName = Office.context.mailbox.userProfile.displayName ?? '';
  try {
    if (isRead) {
      const from = item.from ? [item.from] : [];
      state.ctx = { myName, subject: item.subject ?? '', salutation: buildSalutation(from, settings) };
    } else {
      const [to, cc, subject] = await Promise.all([
        call((cb) => item.to.getAsync(cb)),
        call((cb) => item.cc.getAsync(cb)),
        call((cb) => item.subject.getAsync(cb)),
      ]);
      Object.assign(state, { to, cc });
      state.ctx = { myName, subject, salutation: buildSalutation(to, settings) };
    }
  } catch {
    state.ctx = { myName, subject: '', salutation: '' };
  }
}

// テンプレートの To / CC のうち、まだ入っていない人（To・CC のどちらかに居れば追加しない）
function recipientsToAdd(tpl) {
  const to = missingRecipients([...state.to, ...state.cc], parseRecipients(tpl.to));
  const cc = missingRecipients([...state.to, ...state.cc, ...to], parseRecipients(tpl.cc));
  return { to, cc };
}

// {宛名} はテンプレートで追加する To も含めて作る
function previewCtx() {
  if (state.isRead || !state.current) return state.ctx;
  const { to } = recipientsToAdd(state.current);
  return { ...state.ctx, salutation: buildSalutation([...state.to, ...to], state.settings) };
}

function el(tag, props = {}, ...children) {
  const e = Object.assign(document.createElement(tag), props);
  e.append(...children);
  return e;
}

function show(viewId) {
  for (const v of document.querySelectorAll('.view')) v.hidden = v.id !== viewId;
  window.scrollTo(0, 0);
}

function toast(message, isError = false) {
  const t = $('toast');
  t.textContent = message;
  t.className = isError ? 'error' : '';
  t.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => (t.hidden = true), isError ? 5000 : 2500);
}

// 削除など取り返しのつかない操作は 2 回押しで実行する（アドインでは confirm() が使えない）
function armed(button, label, action) {
  if (button.classList.contains('armed')) {
    disarm(button);
    action();
    return;
  }
  button.dataset.label ??= button.textContent;
  button.textContent = label;
  button.classList.add('armed');
  clearTimeout(button.timer);
  button.timer = setTimeout(() => disarm(button), 4000);
}

function disarm(button) {
  clearTimeout(button.timer);
  button.classList.remove('armed');
  if (button.dataset.label) button.textContent = button.dataset.label;
}

async function persist(list) {
  try {
    await store.saveTemplates(list);
    state.templates = list;
    return true;
  } catch (e) {
    toast(`保存できませんでした：${e.message}`, true);
    return false;
  }
}

function categories() {
  return [...new Set(state.templates.map((t) => t.category).filter(Boolean))];
}

// ---------- 一覧 ----------

function showList() {
  const q = $('search').value.trim().toLowerCase();
  const list = $('list');
  list.replaceChildren();

  const { salutation } = state.ctx;
  $('ctx-info').textContent = salutation
    ? `{宛名} → ${salutation}`
    : state.isRead ? '' : '宛先（To）を入れると {宛名} に名前が入ります';

  const hits = state.templates.filter(
    (t) => !q || [t.name, t.category, t.subject, t.body].join('\n').toLowerCase().includes(q),
  );
  if (!hits.length) {
    list.append(el('p', { className: 'empty' }, q ? '該当するテンプレートがありません' : '＋ からテンプレートを追加してください'));
  }

  const groups = new Map();
  for (const t of hits) {
    const key = t.category || 'その他';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(t);
  }
  for (const [category, items] of groups) {
    list.append(el('h3', {}, category));
    for (const t of items) {
      const btn = el('button', { className: 'tpl' },
        el('span', { className: 'tpl-name' }, t.name),
        el('span', { className: 'tpl-sub' }, summary(t)));
      btn.addEventListener('click', () => showDetail(t.id));
      list.append(btn);
    }
  }
  show('view-list');
}

function summary(t) {
  return t.subject || t.body.split('\n').find((l) => l.trim() && !l.includes('{宛名}')) || '';
}

// ---------- 確認・挿入 ----------

async function showDetail(id) {
  const tpl = state.templates.find((t) => t.id === id);
  if (!tpl) return showList();
  state.current = tpl;
  await refreshContext(); // 宛先は開いた後に変わることがあるので毎回取り直す

  $('d-name').textContent = tpl.name;
  const box = $('d-inputs');
  box.replaceChildren();
  for (const { label, defaultValue } of extractInputs(`${tpl.subject ?? ''}\n${tpl.body}`)) {
    const input = el('input', { value: defaultValue, autocomplete: 'off' });
    input.dataset.label = label;
    input.addEventListener('input', updatePreview);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !state.isRead) insertIntoBody();
    });
    box.append(el('label', {}, label, input));
  }

  const hasSubject = Boolean(tpl.subject);
  $('d-subject-row').hidden = state.isRead || !hasSubject;
  $('d-set-subject').checked = hasSubject && !state.ctx.subject;
  $('d-actions-compose').hidden = state.isRead;
  $('d-actions-read').hidden = !state.isRead;
  showRecipients(tpl);

  updatePreview();
  show('view-detail');
  box.querySelector('input')?.focus();
}

function inputValues() {
  const values = {};
  for (const i of $('d-inputs').querySelectorAll('input')) values[i.dataset.label] = i.value.trim();
  return values;
}

function showRecipients(tpl) {
  const box = $('d-recipients');
  box.replaceChildren();
  if (tpl.to?.trim() || tpl.cc?.trim()) {
    if (state.isRead) {
      box.append(el('div', {}, '※ 返信では To / CC は追加されません。返信画面でこのテンプレートを挿入すると追加されます'));
    } else {
      const { to, cc } = recipientsToAdd(tpl);
      if (to.length) box.append(el('div', {}, `To に追加：${formatRecipients(to)}`));
      if (cc.length) box.append(el('div', {}, `CC に追加：${formatRecipients(cc)}`));
      if (!to.length && !cc.length) box.append(el('div', {}, 'To / CC：すでに入っています'));
    }
  }
  box.hidden = !box.childElementCount;
}

function rendered() {
  const tpl = state.current;
  const ctx = previewCtx();
  const values = inputValues();
  return {
    subject: tpl.subject ? render(tpl.subject, ctx, values) : '',
    body: render(tpl.body, ctx, values),
  };
}

function updatePreview() {
  const r = rendered();
  $('d-subject').textContent = r.subject ? `件名：${r.subject}` : '';
  $('d-subject').hidden = !r.subject || state.isRead;
  $('d-body').textContent = r.body;
}

async function insertIntoBody() {
  const { item } = state;
  try {
    // 宛先を先に入れ、Outlook が名前を解決した後の表示名で {宛名} を作る
    await refreshContext();
    const { to, cc } = recipientsToAdd(state.current);
    if (to.length) await call((cb) => item.to.addAsync(toOfficeRecipients(to), cb));
    if (cc.length) await call((cb) => item.cc.addAsync(toOfficeRecipients(cc), cb));
    if (to.length || cc.length) await refreshContext();

    const r = rendered();
    const type = await call((cb) => item.body.getTypeAsync(cb));
    const html = type === Office.CoercionType.Html;
    await call((cb) =>
      item.body.setSelectedDataAsync(html ? textToHtml(r.body) : r.body,
        { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text }, cb));
    if (r.subject && $('d-set-subject').checked) {
      await call((cb) => item.subject.setAsync(r.subject, cb));
    }
    const added = to.length + cc.length;
    toast(added ? `挿入しました（宛先を ${added} 件追加）` : '挿入しました');
    showRecipients(state.current);
    updatePreview();
  } catch (e) {
    toast(`挿入できませんでした：${e.message}`, true);
  }
}

function reply(all) {
  const html = textToHtml(rendered().body);
  if (all) state.item.displayReplyAllForm(html);
  else state.item.displayReplyForm(html);
}

// ---------- 編集 ----------

function showEditor(id) {
  const tpl = id ? state.templates.find((t) => t.id === id) : null;
  state.editing = tpl?.id ?? null;
  $('e-title').textContent = tpl ? 'テンプレートを編集' : '新しいテンプレート';
  $('e-name').value = tpl?.name ?? '';
  $('e-category').value = tpl?.category ?? '';
  $('e-to').value = tpl?.to ?? '';
  $('e-cc').value = tpl?.cc ?? '';
  $('e-subject').value = tpl?.subject ?? '';
  $('e-body').value = tpl?.body ?? '{宛名}\n\nいつもお世話になっております。{自分}です。\n\n\n\nよろしくお願いいたします。';
  $('categories').replaceChildren(...categories().map((c) => el('option', { value: c })));
  for (const b of ['e-up', 'e-down', 'e-delete']) $(b).hidden = !tpl;
  disarm($('e-delete'));
  state.lastField = $('e-body');
  show('view-edit');
  $('e-name').focus();
}

function insertAtCaret(field, text) {
  const start = field.selectionStart ?? field.value.length;
  const end = field.selectionEnd ?? start;
  field.setRangeText(text, start, end, 'end');
  field.focus();
}

async function saveTemplate() {
  const data = {
    name: $('e-name').value.trim(),
    category: $('e-category').value.trim(),
    to: $('e-to').value.trim(),
    cc: $('e-cc').value.trim(),
    subject: $('e-subject').value.trim(),
    body: $('e-body').value.replace(/\s+$/, ''),
  };
  if (!data.name || !data.body) return toast('名前と本文を入力してください', true);
  const invalid = invalidAddresses([...parseRecipients(data.to), ...parseRecipients(data.cc)]);
  if (invalid.length) return toast(`メールアドレスの形式が正しくありません：${invalid.join(', ')}`, true);

  const list = [...state.templates];
  const i = list.findIndex((t) => t.id === state.editing);
  const id = i >= 0 ? state.editing : store.newId();
  if (i >= 0) list[i] = { ...list[i], ...data };
  else list.push({ id, ...data });

  if (await persist(list)) {
    toast('保存しました');
    showDetail(id);
  }
}

async function moveTemplate(step) {
  const list = [...state.templates];
  const i = list.findIndex((t) => t.id === state.editing);
  const j = i + step;
  if (i < 0 || j < 0 || j >= list.length) return;
  [list[i], list[j]] = [list[j], list[i]];
  if (await persist(list)) toast(`${step < 0 ? '上' : '下'}へ移動しました`);
}

async function deleteTemplate() {
  const list = state.templates.filter((t) => t.id !== state.editing);
  if (await persist(list)) {
    toast('削除しました');
    showList();
  }
}

// ---------- 設定 ----------

function showSettings() {
  const s = state.settings;
  $('s-style').value = s.style;
  $('s-honorific').value = s.honorific;
  $('s-max').value = s.maxNames;
  $('s-group').value = s.groupText;
  $('s-unknown').value = s.unknownName;
  $('s-export-text').hidden = true;
  $('s-copy').hidden = true;
  $('s-import-text').value = '';
  for (const b of ['s-import-replace', 's-reset']) disarm($(b));
  updateUsage();
  previewSalutation();
  show('view-settings');
}

function settingsFromForm() {
  return {
    style: $('s-style').value,
    honorific: $('s-honorific').value.trim(),
    maxNames: Math.max(1, Number($('s-max').value) || 3),
    groupText: $('s-group').value.trim() || '関係者各位',
    unknownName: $('s-unknown').value.trim() || 'ご担当者',
  };
}

async function previewSalutation() {
  const { item, isRead } = state;
  try {
    const recipients = isRead ? (item.from ? [item.from] : []) : await call((cb) => item.to.getAsync(cb));
    $('s-preview').textContent = buildSalutation(recipients, settingsFromForm()) || '（宛先がありません）';
  } catch {
    $('s-preview').textContent = '';
  }
}

function updateUsage() {
  const used = store.usageBytes(state.templates, state.settings);
  $('s-usage').textContent =
    `テンプレート ${state.templates.length} 件／保存容量 ${(used / 1024).toFixed(1)}KB（上限 約${store.LIMIT_BYTES / 1024}KB）`;
}

async function saveSettings() {
  const settings = settingsFromForm();
  try {
    await store.saveSettings(settings);
    state.settings = settings;
    await refreshContext();
    toast('設定を保存しました');
  } catch (e) {
    toast(`保存できませんでした：${e.message}`, true);
  }
}

function exportTemplates() {
  const area = $('s-export-text');
  area.value = store.exportJson(state.templates);
  area.hidden = false;
  $('s-copy').hidden = false;
  area.select();
}

async function copyExport() {
  const area = $('s-export-text');
  area.select();
  try {
    await navigator.clipboard.writeText(area.value);
  } catch {
    document.execCommand('copy'); // クリップボード API が使えない環境向け
  }
  toast('コピーしました');
}

async function importTemplates(replace) {
  let imported;
  try {
    imported = store.parseImport($('s-import-text').value);
  } catch (e) {
    return toast(e.message, true);
  }
  const list = replace ? imported : [...state.templates, ...imported];
  if (await persist(list)) {
    toast(`${imported.length} 件を読み込みました`);
    $('s-import-text').value = '';
    updateUsage();
  }
}

async function resetTemplates() {
  if (await persist(store.defaultTemplates())) {
    toast('サンプルに戻しました');
    updateUsage();
  }
}

// ---------- イベント ----------

function bindEvents() {
  $('btn-new').addEventListener('click', () => showEditor(null));
  $('btn-settings').addEventListener('click', showSettings);
  $('search').addEventListener('input', showList);
  for (const b of document.querySelectorAll('[data-back]')) b.addEventListener('click', showList);

  $('d-insert').addEventListener('click', insertIntoBody);
  $('d-reply').addEventListener('click', () => reply(false));
  $('d-reply-all').addEventListener('click', () => reply(true));
  $('d-edit').addEventListener('click', () => showEditor(state.current?.id));

  for (const f of ['e-subject', 'e-body']) $(f).addEventListener('focus', (e) => (state.lastField = e.target));
  $('e-chips').append(...CHIPS.map(([label, text]) => {
    const b = el('button', { type: 'button', title: text }, label);
    b.addEventListener('mousedown', (e) => e.preventDefault()); // 入力欄のカーソル位置を保つ
    b.addEventListener('click', () => insertAtCaret(state.lastField ?? $('e-body'), text));
    return b;
  }));
  $('e-save').addEventListener('click', saveTemplate);
  $('e-up').addEventListener('click', () => moveTemplate(-1));
  $('e-down').addEventListener('click', () => moveTemplate(1));
  $('e-delete').addEventListener('click', (e) => armed(e.currentTarget, 'もう一度押すと削除', deleteTemplate));

  for (const f of ['s-style', 's-honorific', 's-max', 's-group', 's-unknown']) $(f).addEventListener('input', previewSalutation);
  $('s-save').addEventListener('click', saveSettings);
  $('s-export').addEventListener('click', exportTemplates);
  $('s-copy').addEventListener('click', copyExport);
  $('s-import-add').addEventListener('click', () => importTemplates(false));
  $('s-import-replace').addEventListener('click', (e) => armed(e.currentTarget, 'もう一度押すと全て置き換え', () => importTemplates(true)));
  $('s-reset').addEventListener('click', (e) => armed(e.currentTarget, 'もう一度押すと初期化', resetTemplates));
}
