import { test } from 'node:test';
import assert from 'node:assert/strict';
import { render, extractInputs, textToHtml, nextWeekday, addBusinessDays, formatDate } from '../js/lib/render.js';
import { buildSalutation, cleanName } from '../js/lib/salutation.js';
import { parseImport, exportJson } from '../js/lib/store.js';

const THU = new Date(2026, 9, 15); // 2026-10-15 (木)

test('日付の変数', () => {
  const out = render('{今日}/{明日}/{明後日}/{今日+7}/{今日-1}', { now: THU });
  assert.equal(out, '10月15日(木)/10月16日(金)/10月17日(土)/10月22日(木)/10月14日(水)');
});

test('営業日は土日を飛ばす', () => {
  assert.equal(formatDate(addBusinessDays(THU, 3)), '10月20日(火)');
  assert.equal(render('{1営業日後}', { now: new Date(2026, 9, 16) }), '10月19日(月)');
});

test('来週の曜日（月曜始まり）', () => {
  assert.equal(render('{来週月曜}{来週金曜}{来週日曜}', { now: THU }), '10月19日(月)10月23日(金)10月25日(日)');
  // 日曜日に実行しても「来週」は翌日からの週
  assert.equal(formatDate(nextWeekday(new Date(2026, 9, 18), 1)), '10月19日(月)');
});

test('メール情報の変数', () => {
  const out = render('{宛名}\n{自分}です。件名:{件名} {不明}', { salutation: '山田様', myName: '藤川', subject: 'お見積り' });
  assert.equal(out, '山田様\n藤川です。件名:お見積り {不明}');
});

test('入力欄', () => {
  const text = '{?資料名}を送ります。{?補足:よろしく} {?資料名}';
  assert.deepEqual(extractInputs(text), [
    { label: '資料名', defaultValue: '' },
    { label: '補足', defaultValue: 'よろしく' },
  ]);
  assert.equal(render(text, {}, { 資料名: '見積書' }), '見積書を送ります。よろしく 見積書');
  assert.equal(render(text, {}, {}), '［資料名］を送ります。よろしく ［資料名］');
  // 入力値の中の {…} は展開しない
  assert.equal(render('{?a}', {}, { a: '{今日}' }), '{今日}');
});

test('HTML 変換', () => {
  assert.equal(textToHtml('a<b>\n\nc'), '<div>a&lt;b&gt;</div><div><br></div><div>c</div>');
});

test('表示名の整理', () => {
  assert.equal(cleanName('山田 太郎 / Taro Yamada'), '山田 太郎');
  assert.equal(cleanName('Taro Yamada (山田 太郎)'), '山田 太郎');
  assert.equal(cleanName('山田　太郎（営業部）'), '山田 太郎');
  assert.equal(cleanName('【社外】株式会社ABC 佐藤 花子'), 'ABC 佐藤 花子');
  assert.equal(cleanName('John Smith'), 'John Smith');
});

test('宛名', () => {
  const p = (displayName, emailAddress = 'x@example.com') => ({ displayName, emailAddress });
  assert.equal(buildSalutation([p('山田 太郎')]), '山田様');
  assert.equal(buildSalutation([p('山田 太郎')], { style: 'full' }), '山田 太郎様');
  assert.equal(buildSalutation([p('山田 太郎'), p('佐藤 花子')]), '山田様、佐藤様');
  assert.equal(buildSalutation([p('John Smith')]), 'John Smith様');
  assert.equal(buildSalutation([p('a@example.com'), p('')]), 'ご担当者様');
  assert.equal(buildSalutation([p('A'), p('B'), p('C'), p('D')]), '関係者各位');
  assert.equal(buildSalutation([p('山田 太郎')], { honorific: 'さん' }), '山田さん');
  assert.equal(buildSalutation([]), '');
});

test('書き出し → 読み込み', () => {
  const json = exportJson([{ id: '1', category: '社外', name: 'お礼', subject: 's', body: 'b' }]);
  const list = parseImport(json);
  assert.equal(list.length, 1);
  assert.equal(list[0].name, 'お礼');
  assert.ok(list[0].id && list[0].id !== '1');
  assert.throws(() => parseImport('{'), /JSON/);
  assert.throws(() => parseImport('[{"name":"x"}]'), /body/);
});
