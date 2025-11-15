/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // Office.jsが初期化された後の処理
  console.log('Commands initialized');
});

/**
 * OnMessageSendイベントハンドラー
 *
 * このハンドラーは、ユーザーが送信ボタンをクリックしたときに呼び出されます。
 * 要件定義書に基づいて、以下の処理を行います:
 * 1. 送信を一時的にキャンセル
 * 2. ダイアログウィンドウを開く
 * 3. ダイアログ内でチェックを実行し、結果を表示
 * 4. ユーザーが確認後、送信を実行または編集に戻る
 *
 * @param event イベントオブジェクト
 */
function onMessageSend(event: Office.AddinCommands.Event) {
  console.log('onMessageSend called');

  // ダイアログウィンドウを開く
  // displayInIframe: false で別ウィンドウとして開く
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/taskpane.html',
    {
      height: 70,  // ダイアログの高さ（パーセント）
      width: 50,   // ダイアログの幅（パーセント）
      displayInIframe: false  // 別ウィンドウとして開く
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('ダイアログの表示に失敗しました:', result.error);
        // エラーの場合は送信を許可（ユーザーが送信できなくなるのを防ぐ）
        event.completed({ allowEvent: true });
      } else {
        // ダイアログが開かれた
        const dialog = result.value;

        // ダイアログからのメッセージを受信
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: { message: string }) => {
          try {
            const message = JSON.parse(arg.message);

            if (message.action === 'send') {
              // ユーザーが送信を選択した場合
              console.log('送信を許可します');
              dialog.close();
              event.completed({ allowEvent: true });
            } else if (message.action === 'cancel') {
              // ユーザーがキャンセルした場合
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

        // ダイアログが閉じられた場合（×ボタンなど）
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          console.log('DialogEventReceived:', arg);
          // 12006: ユーザーがダイアログを閉じた
          if (arg.error === 12006) {
            console.log('ダイアログが閉じられました。送信をキャンセルします。');
            event.completed({ allowEvent: false });
          }
        });
      }
    }
  );
}

// グローバルスコープに関数を登録
(global as any).onMessageSend = onMessageSend;

Office.actions.associate('onMessageSend', onMessageSend);
