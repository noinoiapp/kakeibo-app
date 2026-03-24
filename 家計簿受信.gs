// ========================================
// 設定項目
// ========================================
const CONFIG = {
  SPREADSHEET_ID: '1sTkK-3rtt8ZYyzNDJSZ94adKqvm8BlggVXQjswwCY_A', // スプレッドシートIDを設定
  SHEET_NAME: 'MasterData', // シート名を設定
  LINE_CHANNEL_ACCESS_TOKEN: 'GQjbe1Z8eOyfcbogkWhQy4dzTRPLbR92EdFte5XBmms6SV2hf+CZvXFgq1y8NwdOpsZyQmxu8joptn/vrCfoqdbvIYFdHjTcIAN6zJk+ktdyByZqX5CVdM8/78DCHqhIRTpevTKDlR7a14LGYwuHIgdB04t89/1O/w1cDnyilFU=', // LINEのチャンネルアクセストークン
  DISCORD_WEBHOOK_URL: 'https://discord.com/api/webhooks/1485198736811364503/zLpD267IBIYaQ8ixfJnuUB_yXPSvLpXk9YUpVN4sgtgb5FnNZeX_SeVXc1wCCu1SJCPx', // Discord Webhook URL
  USER_NAMES: {
    'Ud5a76fba5bd6217ee04121169f3d2432': 'おにさん', // LINEのユーザーIDと名前を設定
    'U30e38220a8285e373232d9c9e306cefe': 'おねさん'
  }
};

// ========================================
// LINE Webhook受信処理
// ========================================
function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    const events = json.events;
    
    events.forEach(event => {
      if (event.type === 'message' && event.message.type === 'text') {
        const userId = event.source.userId;
        const messageText = event.message.text.trim();
        const replyToken = event.replyToken;
        
        // 「とりけし」コマンドの判定
        if (messageText === 'とりけし') {
          const result = cancelLastEntry();
          // LINEからの取り消しはここで返信のみ（重複通知防止）
          replyToLine(replyToken, result.message);
        } else {
          // 家計簿データを処理
          const result = processKakeiboData(messageText, userId);
           replyToLine(replyToken, result.message);
        }
      }
    });
    
    return ContentService.createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error in doPost:', error);
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========================================
// 家計簿データ処理 (LINEからの入力)
// ========================================
function processKakeiboData(messageText, userId) {
  try {
    // 読点で区切る(、または,)
    const items = messageText.split(/[、,]/).map(item => item.trim());
    
    // 各項目を取得
    const datetime = items[0] || '';
    const amount = items[1] || '';
    const paymentMethod = items[2] || '';
    const payer = items[3] || '';
    const payee = items[4] || '';
    const category = items[5] || '';
    const memo = items[6] || '';
    
    // データ整形
    const formattedData = formatKakeiboData({
      datetime,
      amount,
      paymentMethod,
      payer,
      payee,
      category,
      memo,
      userId
    });
    
    // スプレッドシートに記録
    writeToSpreadsheet(formattedData);
    
    // 通知メッセージ作成 (LINEからの入力はsource指定なしで標準メッセージ)
    const message = createNotificationMessage(formattedData);
    
    return {
      success: true,
      message: message
    };
    
  } catch (error) {
    console.error('Error in processKakeiboData:', error);
    return {
      success: false,
      message: `👻 エラーが発生しました: ${error.message}`
    };
  }
}

// ========================================
// データ整形処理
// ========================================
function formatKakeiboData(data) {
  const userName = CONFIG.USER_NAMES[data.userId] || 'unknown';
  const otherUserName = getOtherUserName(userName);
  
  // 1. 日時の処理(空欄なら現在時刻)
  let datetime;
  
  if (!data.datetime) {
    datetime = new Date();
  } else {
    datetime = new Date(data.datetime);
    // 不正な日付チェック(不正な場合は現在時刻)
    if (isNaN(datetime.getTime())) {
      datetime = new Date();
    }
  }
  
  // 2. 金額の処理
  const amount = data.amount;
  
  // 3. 支払い方法(空欄なら「クレカ」)
  const paymentMethod = data.paymentMethod || 'クレカ';
  
  // 4. 支払い元(空欄なら送信ユーザー名)
  const payer = data.payer || userName;
  
  // 5. 支払い先の処理
  let payee;
  switch(data.payee) {
    case '0':
      payee = userName;
      break;
    case '1':
      payee = otherUserName;
      break;
    case '2':
      payee = '二人';
      break;
    default:
      // 0,1,2以外はそのまま記入、空白の場合は送信者
      payee = data.payee || userName;
  }
  
  // 6. 科目
  const category = data.category;
  
  // 7. 詳細メモ
  const memo = data.memo;
  
  return {
    datetime,
    amount,
    paymentMethod,
    payer,
    payee,
    category,
    memo
  };
}

// ========================================
// もう一人のユーザー名を取得
// ========================================
function getOtherUserName(currentUserName) {
  const userNames = Object.values(CONFIG.USER_NAMES);
  return userNames.find(name => name !== currentUserName) || 'unknown';
}

// ========================================
// スプレッドシートへの書き込み
// ========================================
function writeToSpreadsheet(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`シート "${CONFIG.SHEET_NAME}" が見つかりません`);
  }
  
  // 最終行の次の行に追記
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  
  // データを配列として準備
  const rowData = [
    data.datetime,
    data.amount,
    data.paymentMethod,
    data.payer,
    data.payee,
    data.category,
    data.memo
  ];
  
  // データを書き込み
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  
  // 日付列(A列)の表示形式を設定
  sheet.getRange(newRow, 1).setNumberFormat('yyyy/MM/dd (ddd)');
  
  return {
    sheetId: sheet.getSheetId(),
    row: newRow
  };
}

// ========================================
// 最後の記録を取り消す (コア処理)
// ========================================
function cancelLastEntry() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`シート "${CONFIG.SHEET_NAME}" が見つかりません`);
    }
    
    const lastRow = sheet.getLastRow();
    
    // ヘッダー行のみの場合(データが無い)
    if (lastRow <= 1) {
      return {
        success: false,
        message: '👻 取り消すデータがありません'
      };
    }
    
    // 最終行のデータを取得(取り消し前に内容を表示)
    const lastData = sheet.getRange(lastRow, 1, 1, 7).getValues()[0];
    const datetime = lastData[0];
    const amount = lastData[1];
    const category = lastData[5];
    
    // 最終行をクリア
    sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).clear();
    
    // メッセージ作成
    const message = `🗑️ 取り消しました\n日時: ${Utilities.formatDate(datetime, 'Asia/Tokyo', 'yyyy年M月d日')}\n金額: ${amount}円\n科目: ${category}`;

    // ※ここでの pushToLineGroup 呼び出しは削除（呼び出し元で制御）
    
    return {
      success: true,
      message: message,
      sheetId: sheet.getSheetId(),
      row: lastRow
    };
    
  } catch (error) {
    console.error('Error in cancelLastEntry:', error);
    return {
      success: false,
      message: `👻 取り消しエラー: ${error.message}`
    };
  }
}

// ========================================
// Webからの取り消し (ラッパー関数)
// ========================================
function cancelLastEntryFromWeb() {
  try {
    const result = cancelLastEntry();
    
    if (result.success) {
      // Webからの取り消し専用メッセージヘッダーを付与
      const webMessage = '🗑webから取り消し📱\n' + result.message.replace('🗑️ 取り消しました\n', '');
      
      // Discordへ通知
      pushToDiscord(webMessage, 'warning', result.sheetId, result.row);
      
      return result;
    } else {
      pushToDiscord(`👻 エラー\n取り消し処理に失敗しました: ${result.message}`, 'error');
      return result;
    }
  } catch (e) {
     pushToDiscord(`👻 エラー\n取り消し処理中にシステムエラーが発生しました: ${e.message}`, 'error');
     return {
      success: false,
      message: `👻 エラー: ${e.message}`
    };
  }
}

// ========================================
// LINEへの返信
// ========================================
function replyToLine(replyToken, message) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN
  };
  
  const payload = {
    replyToken: replyToken,
    messages: [{
      type: 'text',
      text: message
    }]
  };
  
  const options = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload)
  };
  
  UrlFetchApp.fetch(url, options);
}

// ========================================
// Discordにプッシュ通知
// ========================================
function pushToDiscord(message, type = 'info', sheetId = null, row = null) {
  const url = CONFIG.DISCORD_WEBHOOK_URL;
  if (!url) return;
  
  // 1行目をタイトル、それ以降を説明文にする
  const lines = message.split('\n');
  const title = lines[0] || '通知';
  let description = lines.slice(1).join('\n').trim();
  
  // メンション処理用のフラグとテキスト置換
  const mentions = [];
  if (description.includes('おにさん')) mentions.push('@noinoiapp');
  if (description.includes('おねさん')) mentions.push('@nade322');
  
  // Discordの文章上にも @noinoiapp と表示するため置換します
  description = description.replace(/おにさん/g, '@noinoiapp').replace(/おねさん/g, '@nade322');
  
  // スプレッドシートのリンク追加
  let sheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit`;
  if (sheetId !== null) {
    sheetUrl += `#gid=${sheetId}`;
    if (row !== null) {
      sheetUrl += `&range=A${row}`;
    }
  }
  
  description += `\n\n[スプレッドシートで確認](${sheetUrl})`;
  
  // 色の設定 ("webから記録"・"定期支払い正常"=水色, "取り消し"=黄色, "エラー"=赤色)
  let color = 49151; // 水色 (info)
  if (type === 'warning') color = 16766720; // 黄色
  if (type === 'error') color = 16711680; // 赤色
  
  const payload = {
    content: mentions.length > 0 ? mentions.join(' ') : "", 
    embeds: [{
      title: title,
      description: description,
      color: color
    }]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(url, options);
  } catch (error) {
    console.error('Discord通知エラー:', error);
  }
}

// ========================================
// 支払い先の名前解決
// ========================================
function resolvePayee(payer, payeeType) {
  // payeeType: '0'=自分, '1'=相手, '2'=二人
  // 既に名前が入っている場合はそのまま返す（念のため）
  if (Object.values(CONFIG.USER_NAMES).includes(payeeType)) {
    return payeeType;
  }

  const payerName = payer;
  const otherName = getOtherUserName(payerName);

  if (String(payeeType) === '0') {
    return payerName;
  } else if (String(payeeType) === '1') {
    return otherName;
  } else if (String(payeeType) === '2') {
    return '二人';
  } else {
    // 未定義の場合はそのまま、またはpayer
    return payeeType || payerName;
  }
}

// ========================================
// 通知メッセージ作成（共通）
// ========================================
function createNotificationMessage(data, source) {
  let header = '✏️ 記録しました'; // デフォルト (LINE用など)
  if (source === 'web') {
    header = '✏️webから記録📱';
  }
  
  return `${header}\n日時: ${Utilities.formatDate(data.datetime, 'Asia/Tokyo', 'yyyy年M月d日')}\n金額: ${data.amount}円\n支払い方法: ${data.paymentMethod}\n支払い元: ${data.payer}\n支払い先: ${data.payee}\n科目: ${data.category}\n詳細: ${data.memo || ''}`;
}

// ========================================
// Webアプリ: 初期表示 (doGet)
// ========================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('家計簿入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========================================
// Webアプリ: フォーム用データ取得
// ========================================
function getDataForForm() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 科目リストの取得 (シート「科目リスト」のA列)
  let categories = [];
  try {
    const sheet = ss.getSheetByName('科目リスト');
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 1) {
        // A列のデータを取得 (ヘッダーがある場合は2行目からにするなど調整可能だが、一旦A1から全て取得して空文字除去)
        const values = sheet.getRange(1, 1, lastRow, 1).getValues();
        categories = values.flat().filter(String); // 空文字を除去
      }
    } else {
      console.warn('科目リストシートが見つかりません');
      categories = ['食費', '日用品', '交通費', '娯楽費', 'その他']; // デフォルト値
    }
  } catch (e) {
    console.warn('科目リスト取得エラー:', e);
    categories = ['食費', '日用品', '交通費', '娯楽費', 'その他'];
  }

  return {
    categories: categories,
    userNames: CONFIG.USER_NAMES
  };
}

// ========================================
// Webアプリ: フォーム送信処理
// ========================================
function processForm(formData) {
  try {
    // データ整形
    // 日時はフォームから "yyyy-MM-dd" で来るのでDateオブジェクトに変換
    const datetime = new Date(formData.datetime);
    
    // データオブジェクト作成
    const data = {
      datetime: datetime,
      amount: formData.amount,
      paymentMethod: formData.paymentMethod,
      payer: formData.payer, // 名前で来る
      payee: resolvePayee(formData.payer, formData.payee), // 0,1,2 を名前に変換
      category: formData.category,
      memo: formData.memo
    };
    
    // スプレッドシートに記録
    const result = writeToSpreadsheet(data);
    
    // LINE通知用のメッセージを作成 (source='web'を指定)
    const notificationMessage = createNotificationMessage(data, 'web');
    
    // Discordへ通知
    pushToDiscord(notificationMessage, 'info', result.sheetId, result.row);

    return {
      success: true,
      message: '✅ 記録しました！'
    };
    
  } catch (error) {
    console.error('Error in processForm:', error);
    pushToDiscord(`👻 エラー\nプロセスの実行中にエラーが発生しました: ${error.message}`, 'error');
    return {
      success: false,
      message: '❌ エラー: ' + error.message
    };
  }
}
