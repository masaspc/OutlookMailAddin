/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { SettingsStorage } from '../storage/SettingsStorage';
import { CheckerManager } from '../checker/CheckerManager';
import { MailData, CheckResult } from '../types';

Office.onReady(() => {
  // Office.jsが初期化された後の処理
  console.log('Commands initialized');
});

/**
 * メールデータを取得する
 */
async function getMailData(item: Office.MessageCompose): Promise<MailData> {
  return new Promise((resolve, reject) => {
    // 件名を取得
    const subject = item.subject.getAsync((subjectResult) => {
      if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error('件名の取得に失敗しました'));
        return;
      }

      const subjectValue = subjectResult.value || '';

      // 本文を取得
      item.body.getAsync(Office.CoercionType.Html, async (bodyResult) => {
        if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error('本文の取得に失敗しました'));
          return;
        }

        const body = bodyResult.value;

        try {
          // 宛先を取得
          const to = await getRecipients(item.to);
          const cc = await getRecipients(item.cc);
          const bcc = await getRecipients(item.bcc);

          // 添付ファイルの取得（MessageComposeでは直接アクセスできないため、
          // 添付ファイルAPIを使用して取得する）
          const attachments = await getAttachments(item);

          resolve({
            subject: subjectValue,
            body,
            to,
            cc,
            bcc,
            attachments,
          });
        } catch (error) {
          reject(error);
        }
      });
    });
  });
}

/**
 * 受信者のメールアドレスを取得する
 */
async function getRecipients(recipients: Office.Recipients | undefined): Promise<string[]> {
  if (!recipients) return [];

  return new Promise((resolve) => {
    recipients.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emails = result.value.map((r) => r.emailAddress);
        resolve(emails);
      } else {
        resolve([]);
      }
    });
  });
}

/**
 * 添付ファイル情報を取得する
 * MessageComposeでは直接アクセスできないため、itemをanyにキャストして取得
 */
async function getAttachments(item: Office.MessageCompose): Promise<any[]> {
  try {
    // MessageComposeではattachmentsプロパティが型定義にないが、
    // 実際には存在するため、anyにキャストして取得
    const attachments = (item as any).attachments || [];
    return attachments;
  } catch (error) {
    console.error('添付ファイルの取得に失敗しました:', error);
    return [];
  }
}

/**
 * OnMessageSendイベントハンドラー
 *
 * このハンドラーは、ユーザーが送信ボタンをクリックしたときに呼び出されます。
 * 要件定義書に基づいて、以下の処理を行います:
 * 1. メールデータを取得してチェックを実行
 * 2. チェック結果をsessionDataに保存
 * 3. ダイアログウィンドウを開く
 * 4. ユーザーが確認後、送信を実行または編集に戻る
 *
 * @param event イベントオブジェクト
 */
async function onMessageSend(event: Office.AddinCommands.Event) {
  console.log('onMessageSend called');

  const item = Office.context.mailbox.item as Office.MessageCompose;
  if (!item) {
    console.error('メールアイテムが見つかりません');
    event.completed({ allowEvent: true });
    return;
  }

  try {
    // メールデータを取得
    const mailData = await getMailData(item);

    // 設定を読み込み
    const settings = SettingsStorage.loadSettings();

    // チェックを実行
    const checkResults = await CheckerManager.checkAll(mailData, settings);

    // チェック結果とメールデータをBase64エンコードしてURLパラメータとして渡す
    const dataToPass = {
      checkResults,
      mailData: {
        subject: mailData.subject,
        body: mailData.body,
        to: mailData.to,
        cc: mailData.cc,
        bcc: mailData.bcc,
        attachments: mailData.attachments.map((a: any) => ({
          id: a.id,
          name: a.name,
          size: a.size,
          attachmentType: a.attachmentType,
        })),
      },
      settings,
    };
    const dataJson = JSON.stringify(dataToPass);
    const dataBase64 = btoa(encodeURIComponent(dataJson));
    const dialogUrl = `https://localhost:3000/taskpane.html?data=${dataBase64}`;

    // ダイアログウィンドウを開く
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      {
        height: 80,
        width: 60,
        displayInIframe: false
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error('ダイアログの表示に失敗しました:', result.error);
          // エラーの場合は送信を許可
          event.completed({ allowEvent: true });
        } else {
          const dialog = result.value;

          // ダイアログからのメッセージを受信
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
            try {
              const message = JSON.parse(arg.message);

              if (message.action === 'send') {
                console.log('送信を許可します');
                dialog.close();
                event.completed({ allowEvent: true });
              } else if (message.action === 'cancel') {
                console.log('送信をキャンセルします');
                dialog.close();
                event.completed({ allowEvent: false });
              }
            } catch (error) {
              console.error('メッセージの解析に失敗しました:', error);
              dialog.close();
              event.completed({ allowEvent: false });
            }
          });

          // ダイアログが閉じられた場合
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
            console.log('DialogEventReceived:', arg);
            if (arg.error === 12006) {
              console.log('ダイアログが閉じられました。送信をキャンセルします。');
              event.completed({ allowEvent: false });
            }
          });
        }
      }
    );
  } catch (error) {
    console.error('エラーが発生しました:', error);
    // エラーの場合は送信を許可（ユーザーが送信できなくなるのを防ぐ）
    event.completed({ allowEvent: true });
  }
}

// Office.actionsに関数を登録
Office.actions.associate('onMessageSend', onMessageSend);
