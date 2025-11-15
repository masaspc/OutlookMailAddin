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
 * 2. Task Paneを開く
 * 3. Task Pane側でチェックを実行し、結果を表示
 * 4. ユーザーが確認後、Task Pane側から送信を実行
 *
 * @param event イベントオブジェクト
 */
function onMessageSend(event: Office.AddinCommands.Event) {
  console.log('onMessageSend called');

  // イベントオブジェクトをグローバルに保存（Task Pane側で送信を完了するため）
  // Note: Office.contextを通じてTask Paneとデータを共有
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/taskpane.html',
    { height: 60, width: 40, displayInIframe: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error('Task Paneの表示に失敗しました:', result.error);
        // エラーの場合は送信を許可（ユーザーが送信できなくなるのを防ぐ）
        event.completed({ allowEvent: true });
      } else {
        // Task Paneが開かれた
        const dialog = result.value;

        // Task Paneからのメッセージを受信
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: { message: string }) => {
          const message = JSON.parse(arg.message);

          if (message.action === 'send') {
            // ユーザーが送信を選択した場合
            dialog.close();
            event.completed({ allowEvent: true });
          } else if (message.action === 'cancel') {
            // ユーザーがキャンセルした場合
            dialog.close();
            event.completed({ allowEvent: false });
          }
        });

        // ダイアログが閉じられた場合（×ボタンなど）
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          if (arg.error === 12006) {
            // ユーザーがダイアログを閉じた場合
            event.completed({ allowEvent: false });
          }
        });
      }
    }
  );

  // 送信を一時的にキャンセル
  // Task Pane側でユーザーが確認後、送信を許可する
  // Note: displayDialogAsync内でevent.completedを呼び出すため、ここでは何もしない
}

// グローバルスコープに関数を登録
(global as any).onMessageSend = onMessageSend;

Office.actions.associate('onMessageSend', onMessageSend);
