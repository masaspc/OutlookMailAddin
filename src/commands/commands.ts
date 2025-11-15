/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import { MailData } from '../types';

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
 * MessageComposeではgetAttachmentsAsyncを使用
 */
async function getAttachments(item: Office.MessageCompose): Promise<any[]> {
  return new Promise((resolve) => {
    // Office.jsのgetAttachmentsAsyncメソッドを使用
    if (item.getAttachmentsAsync) {
      item.getAttachmentsAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log('添付ファイル取得成功:', result.value);
          resolve(result.value || []);
        } else {
          console.error('添付ファイルの取得に失敗しました:', result.error);
          resolve([]);
        }
      });
    } else {
      // フォールバック: anyにキャストして直接取得
      console.log('getAttachmentsAsyncが利用できません。直接取得を試みます。');
      const attachments = (item as any).attachments || [];
      console.log('直接取得した添付ファイル:', attachments);
      resolve(attachments);
    }
  });
}

/**
 * メール内容に応じて最適なダイアログサイズを計算する
 * @param mailData メールデータ
 * @returns ダイアログのheightとwidth（パーセンテージ）
 */
function calculateOptimalDialogSize(mailData: any): { height: number; width: number } {
  // 基本サイズ
  let height = 50;
  let width = 50;

  // 外部宛先の数をカウント（社内ドメインフィルタリングは後で行うため、ここでは全宛先をカウント）
  const totalRecipients = (mailData.to?.length || 0) + (mailData.cc?.length || 0) + (mailData.bcc?.length || 0);

  // 添付ファイルの数
  const attachmentCount = mailData.attachments?.length || 0;

  // 件名の有無
  const hasSubject = mailData.subject && mailData.subject.trim() !== '';

  // 本文の有無
  const hasBody = mailData.body && mailData.body.trim() !== '';

  // 宛先が多い場合は高さを追加（1件あたり3%、最大30%）
  if (totalRecipients > 0) {
    height += Math.min(totalRecipients * 3, 30);
  }

  // 添付ファイルが多い場合は高さを追加（1件あたり4%、最大20%）
  if (attachmentCount > 0) {
    height += Math.min(attachmentCount * 4, 20);
  }

  // 件名がある場合
  if (hasSubject) {
    height += 5;
  }

  // 本文がある場合
  if (hasBody) {
    height += 8;
  }

  // 最小・最大値の制限
  height = Math.max(40, Math.min(90, height));
  width = Math.max(50, Math.min(70, width));

  return { height, width };
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

    // 本文が長すぎる場合は切り詰める（URLパラメータの長さ制限対策）
    const maxBodyLength = 1000; // 最大1000文字
    const truncatedBody = mailData.body.length > maxBodyLength
      ? mailData.body.substring(0, maxBodyLength) + '...(以降省略)'
      : mailData.body;

    // メールデータをBase64エンコードしてURLパラメータとして渡す
    // チェック処理はダイアログ内で実行する（処理時間短縮のため）
    const dataToPass = {
      mailData: {
        subject: mailData.subject,
        body: truncatedBody, // 切り詰めた本文を使用
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
    };
    const dataJson = JSON.stringify(dataToPass);
    const dataBase64 = btoa(encodeURIComponent(dataJson));

    // 現在の実行環境に応じてURLを生成
    // commands.htmlと同じディレクトリのtaskpane.htmlを開く
    const baseUrl = window.location.href.replace(/\/[^/]*$/, '');
    const dialogUrl = `${baseUrl}/taskpane.html?data=${dataBase64}`;

    console.log('Opening dialog with URL:', dialogUrl);
    console.log('Base URL:', baseUrl);

    // 内容に応じて最適なダイアログサイズを計算
    const dialogSize = calculateOptimalDialogSize(mailData);
    console.log('Calculated dialog size:', dialogSize);

    // ダイアログウィンドウを開く
    // promptBeforeOpen: false でOfficeの許可プロンプトを無効化
    // displayInIframe: false でポップアップ表示（URLバーが表示される）
    // displayInIframe: true はOnMessageSend + SoftBlockでError 9032を引き起こすため使用不可
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      {
        height: dialogSize.height,
        width: dialogSize.width,
        displayInIframe: false,
        promptBeforeOpen: false
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error('ダイアログの表示に失敗しました:', result.error);
          console.error('Error code:', result.error.code);
          console.error('Error message:', result.error.message);
          // エラーの場合は送信を許可
          event.completed({ allowEvent: true });
        } else {
          console.log('ダイアログが正常に開きました');
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
