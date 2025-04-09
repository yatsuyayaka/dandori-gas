// できたよヤッホーーい

/**
 * スプレッドシート連携 Chatwork 通知システム
 */

// 設定項目
const SPREADSHEET_ID = '1nq2CCC2YkqE2poY8n0-_9EApnOP4r_0RNoGudqAIh0g'; // スプレッドシートID
const SHEET_NAME = '定常稟議管理'; // シート名

// ※ トークンは Script Properties から取得する実装へ変更済み
function getChatworkApiToken() {
  return PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
}

const CHATWORK_ROOM_ID = '249879599'; // Chatwork ルームID

// メールアドレスとChatworkアカウントIDの対応表
const accountMap = {
  "ayatsu@iyell.jp": "4205188",
  "khatsuse@iyell.jp": "2253961",
  "tabe@iyell.jp": "3772929",
  // 必要に応じて追加
};

// 追加メッセージ (請求書)
const additionalMessage = "請求書がないよーー！\n各自取得して、下記ドライブに格納してくださいな⸜（ ´ ꒳ ` ）⸝\nhttps://drive.google.com/drive/folders/1DvyR4zS3C1TqYux0-gk_1Q7aJlLjGfVo";

// 追加メッセージ (事前申請)
const expirationMessage = "事前申請が切れちゃうよ！\n今月中に承認が取れるように対応お願いいたします(　-`ω-)✧\nhttps://drive.google.com/drive/folders/1toBeG_WqJfoFyglpzs-6F9G7JftdUsTx";

function main() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const now = new Date();
  const currentMonth = now.getMonth() + 1; // 1月が0なので+1
  const currentYear = now.getFullYear();

  // データ範囲を取得
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  let notificationCount = 0;
  let notificationMessages = []; // 通知メッセージ (請求書)
  let expirationNotificationMessages = []; // 通知メッセージ (事前申請)

  // ヘッダー行をスキップ
  for (let i = 1; i < lastRow; i++) {
    const row = values[i];
    const paymentType = row[0]; // A列: 支払いタイプ
    const message = row[1];     // B列: 通知内容
    const email = row[4];       // E列: メールアドレス
    const targetMonth = row[6]; // G列: 年払い対象月
    const cellValue = row[lastColumn - 1]; // 一番右の列のセルの値(チェックボックス)
    const endDate = row[13];    // N列: 終了日 (Date オブジェクト)

    // N列の日付が当月かどうか確認
    let isExpiring = false;
    if (endDate instanceof Date) {
      const endMonth = endDate.getMonth() + 1;
      const endYear = endDate.getFullYear();
      if (endYear === currentYear && endMonth === currentMonth) {
        isExpiring = true;
      }
    }

    // ここで事前申請通知(終了日が今月)をチェック
    if (isExpiring) {
      const accountId = accountMap[email];
      if (accountId) {
        expirationNotificationMessages.push(`[To:${accountId}] ${message}`);
        notificationCount++;
        // ▼ ログを日本語に修正 ▼
        Logger.log(`事前申請切れ通知を${email}宛（ChatworkアカウントID: ${accountId}）に追加しました。`);
      } else {
        // ▼ ログを日本語に修正 ▼
        Logger.log(`メールアドレス「${email}」に対応するChatworkアカウントIDが見つからないため、事前申請切れ通知をスキップします。`);
      }
    }

    // チェック項目がない場合はスルー
    if (typeof cellValue === 'undefined') {
      Logger.log(`チェック項目がないためスキップします。 行: ${i + 1}`);
      continue;
    }

    // チェックボックスが存在しない場合はスキップ
    if (typeof cellValue === 'boolean') {
      var isChecked = cellValue;
    } else {
      Logger.log(`チェックボックスではないためスキップします。 行: ${i + 1}`);
      continue;
    }

    // チェック対象列が未チェックの場合のみ処理 (請求書リマインド)
    if (!isChecked) {
      let shouldNotify = false;
      if (paymentType === "月払い") {
        shouldNotify = true;
      } else if (paymentType === "年払い" && targetMonth === currentMonth) {
        shouldNotify = true;
      }

      if (shouldNotify) {
        const accountId = accountMap[email];
        if (accountId) {
          notificationMessages.push(`[To:${accountId}] ${message}`);
          notificationCount++;
          Logger.log(`Added notification to ${email} (Account ID: ${accountId})`);
        } else {
          Logger.log(`Account ID not found for ${email}. Skipping notification.`);
        }
      }
    }
  }

  let combinedMessage = "";

  // 通知メッセージを結合 (請求書)
  if (notificationMessages.length > 0) {
    combinedMessage += notificationMessages.join('\n') + '\n\n' + additionalMessage + '\n\n';
  }

  // 通知メッセージを結合 (事前申請)
  if (expirationNotificationMessages.length > 0) {
    combinedMessage += expirationNotificationMessages.join('\n') + '\n\n' + expirationMessage;
  }

  // まとめて送信
  if (combinedMessage.length > 0) {
    if (accountMap && Object.keys(accountMap).length > 0) {
      sendToChatwork(combinedMessage);
      Logger.log(`Sent ${notificationCount} notifications.`);
    } else {
      Logger.log("メンション対象がいないため通知をスキップします。");
    }
  } else {
    Logger.log("No notifications to send.");
  }
}

/**
 * Chatworkへメッセージを送信する
 * @param {string} message 送信するメッセージ
 */
function sendToChatwork(message) {
  const url = `https://api.chatwork.com/v2/rooms/${CHATWORK_ROOM_ID}/messages`;
  const options = {
    "method": "post",
    "headers": {
      "X-ChatWorkToken": getChatworkApiToken(),
      "Content-Type": "application/x-www-form-urlencoded"
    },
    "payload": `body=${encodeURIComponent(message)}`
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(response.getContentText()); // レスポンスをログに出力
  } catch (error) {
    Logger.log(`Chatwork API Error: ${error}`);
  }
}