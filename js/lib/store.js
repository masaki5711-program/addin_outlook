// テンプレートと設定の保存（roamingSettings = メールボックスに保存され、別の PC でも共有される）

import { DEFAULT_TEMPLATES } from './defaults.js';
import { DEFAULT_SALUTATION } from './salutation.js';

// roamingSettings の上限は 32KB。余裕を持たせて判定する
export const LIMIT_BYTES = 30 * 1024;

const rs = () => Office.context.roamingSettings;

export function newId() {
  return globalThis.crypto?.randomUUID?.() ?? `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

export function defaultTemplates() {
  return DEFAULT_TEMPLATES.map((t) => ({ id: newId(), ...t }));
}

export function loadTemplates() {
  const v = rs().get('templates');
  return Array.isArray(v) ? v : defaultTemplates();
}

export function loadSettings() {
  return { ...DEFAULT_SALUTATION, ...(rs().get('settings') ?? {}) };
}

export function usageBytes(templates, settings) {
  return new TextEncoder().encode(JSON.stringify({ templates, settings })).length;
}

function checkSize(templates, settings) {
  if (usageBytes(templates, settings) > LIMIT_BYTES) {
    throw new Error('保存できる容量（約30KB）を超えています。不要なテンプレートを削除してください');
  }
}

function saveAsync() {
  return new Promise((resolve, reject) =>
    rs().saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error))),
  );
}

export async function saveTemplates(templates) {
  checkSize(templates, loadSettings());
  rs().set('templates', templates);
  await saveAsync();
}

export async function saveSettings(settings) {
  checkSize(loadTemplates(), settings);
  rs().set('settings', settings);
  await saveAsync();
}

export function exportJson(templates) {
  const list = templates.map(({ category, name, subject, body }) => ({ category, name, subject, body }));
  return JSON.stringify({ app: 'mail-assist', version: 1, templates: list }, null, 2);
}

export function parseImport(text) {
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    throw new Error('JSON の形式が正しくありません');
  }
  const list = Array.isArray(data) ? data : data?.templates;
  if (!Array.isArray(list)) throw new Error('templates の一覧が見つかりません');
  return list.map((t, i) => {
    if (typeof t?.name !== 'string' || typeof t?.body !== 'string') {
      throw new Error(`${i + 1} 件目に name または body がありません`);
    }
    return { id: newId(), category: String(t.category ?? ''), name: t.name, subject: String(t.subject ?? ''), body: t.body };
  });
}
