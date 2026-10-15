// リボンのボタンから作業ウィンドウを開かずに実行する処理

import { textToHtml } from './lib/render.js';
import { buildSalutation } from './lib/salutation.js';
import { loadSettings } from './lib/store.js';

function notify(message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('mailAssist', {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message,
    icon: 'Icon.16x16',
    persistent: false,
  });
}

// 本文の先頭に「山田様」などの宛名を入れる
function insertSalutation(event) {
  const item = Office.context.mailbox.item;
  item.to.getAsync((to) => {
    const text = to.status === Office.AsyncResultStatus.Succeeded ? buildSalutation(to.value, loadSettings()) : '';
    if (!text) {
      notify('宛先（To）を入力してから「宛名を挿入」を押してください');
      event.completed();
      return;
    }
    item.body.getTypeAsync((type) => {
      const html = type.value === Office.CoercionType.Html;
      item.body.prependAsync(
        html ? textToHtml(`${text}\n`) : `${text}\n\n`,
        { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text },
        () => event.completed(),
      );
    });
  });
}

// 古いクラシック Outlook はグローバル関数を、新しい環境は associate で登録した関数を呼ぶ
globalThis.insertSalutation = insertSalutation;
Office.onReady(() => Office.actions?.associate('insertSalutation', insertSalutation));
