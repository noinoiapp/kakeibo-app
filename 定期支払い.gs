// ========================================
// 定期支払い自動実行システム
// ========================================

// ========================================
// 定期支払い自動実行(トリガーで毎日実行)
// ========================================
function executeRecurringPayments() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const masterSheet = ss.getSheetByName('定期支払いマスタ');
    
    if (!masterSheet) {
      throw new Error('定期支払いマスタシートが見つかりません');
    }
    
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    const currentDay = today.getDate();
    
    const results = [];
    const errors = [];
    
    // マスタシートのデータを取得(ヘッダー行を除く)
    const lastRow = masterSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('定期支払いマスタにデータがありません');
      return;
    }
    
    const data = masterSheet.getRange(2, 1, lastRow - 1, 11).getValues(); // 11列に変更
    
    // 各行を処理
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowIndex = i + 2;
      
      const isActive = row[0]; // A列: 有効
      const paymentName = row[1]; // B列: 支払い名
      const amount = row[2]; // C列: 金額
      const paymentMethod = row[3]; // D列: 支払い方法
      const payer = row[4]; // E列: 支払い元
      const payee = row[5]; // F列: 支払い先
      const category = row[6]; // G列: 科目
      // row[7] は詳細メモ(H列) - 使用しない
      const frequency = row[8]; // I列: 頻度
      const executionDay = row[9]; // J列: 実行日
      const lastExecutionDate = row[10]; // K列: 最終実行日
      
      // 有効フラグチェック
      if (isActive !== '○' && isActive !== true && isActive !== 'TRUE') {
        continue;
      }
      
      // 今日実行すべきか判定
      if (!shouldExecuteToday(today, frequency, executionDay, lastExecutionDate)) {
        continue;
      }
      
      // 家計簿に記録
      try {
        const formattedData = {
          datetime: today,
          amount: amount,
          paymentMethod: paymentMethod || 'クレカ',
          payer: payer || 'おにさん',
          payee: String(payee) === '2' ? '二人' : (payee || 'おにさん'),
          category: category,
          memo: paymentName
        };
        
        writeToSpreadsheet(formattedData);
        
        // 最終実行日を更新
        masterSheet.getRange(rowIndex, 11).setValue(today); // K列に変更
        
        results.push(`✏️ ${paymentName}: ${amount}円`);
        
      } catch (error) {
        console.error(`Error processing ${paymentName}:`, error);
        errors.push(`👻 ${paymentName}: ${error.message}`);
      }
    }
    
    // Discordに通知(記録があった場合のみ)
    if (results.length > 0 || errors.length > 0) {
      let message = `【定期支払い自動記録 ${currentYear}年${currentMonth}月${currentDay}日】\n\n`;
      
      if (results.length > 0) {
        message += results.join('\n');
      }
      
      if (errors.length > 0) {
        message += '\n\n【エラー】\n' + errors.join('\n');
      }
      
      // エラーがあれば赤色(error)、なければ水色(info)
      const type = errors.length > 0 ? 'error' : 'info';
      const targetSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      const sheetId = targetSheet ? targetSheet.getSheetId() : null;
      
      pushToDiscord(message, type, sheetId);
    }
    
    console.log(`定期支払い実行完了: 成功${results.length}件, エラー${errors.length}件`);
    
  } catch (error) {
    console.error('Error in executeRecurringPayments:', error);
    pushToDiscord(`👻 定期支払い実行エラー\nシステムエラー: ${error.message}`, 'error');
  }
}

// ========================================
// 今日実行すべきか判定
// ========================================
function shouldExecuteToday(today, frequency, executionDay, lastExecutionDate) {
  // 最終実行日が今日なら実行済み
  if (lastExecutionDate) {
    const lastDate = new Date(lastExecutionDate);
    if (isSameDay(today, lastDate)) {
      return false;
    }
  }
  
  if (frequency === '月次') {
    return shouldExecuteMonthly(today, executionDay, lastExecutionDate);
  } else if (frequency === '年次') {
    return shouldExecuteYearly(today, executionDay, lastExecutionDate);
  }
  
  return false;
}

// ========================================
// 月次支払いの実行判定
// ========================================
function shouldExecuteMonthly(today, executionDay, lastExecutionDate) {
  const targetDay = parseInt(executionDay);
  
  if (isNaN(targetDay)) {
    return false;
  }
  
  // 今月の最終日を取得
  const lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  
  // 31日設定で30日の月は30日に実行
  const actualExecutionDay = Math.min(targetDay, lastDayOfMonth);
  
  // 今日が実行日か判定
  if (today.getDate() !== actualExecutionDay) {
    return false;
  }
  
  // 最終実行日が今月ならスキップ
  if (lastExecutionDate) {
    const lastDate = new Date(lastExecutionDate);
    if (lastDate.getFullYear() === today.getFullYear() && 
        lastDate.getMonth() === today.getMonth()) {
      return false;
    }
  }
  
  return true;
}

// ========================================
// 年次支払いの実行判定
// ========================================
function shouldExecuteYearly(today, executionDay, lastExecutionDate) {
  let targetMonth, targetDay;
  
  // executionDayがDateオブジェクトの場合
  if (executionDay instanceof Date) {
    targetMonth = executionDay.getMonth() + 1;
    targetDay = executionDay.getDate();
  } else {
    // 文字列の場合: "2/14" (月/日)
    const parts = String(executionDay).split('/');
    
    if (parts.length !== 2) {
      return false;
    }
    
    targetMonth = parseInt(parts[0]);
    targetDay = parseInt(parts[1]);
    
    if (isNaN(targetMonth) || isNaN(targetDay)) {
      return false;
    }
  }
  
  // 今日が実行日か判定
  if (today.getMonth() + 1 !== targetMonth || today.getDate() !== targetDay) {
    return false;
  }
  
  // 最終実行日が今年ならスキップ
  if (lastExecutionDate) {
    const lastDate = new Date(lastExecutionDate);
    if (lastDate.getFullYear() === today.getFullYear()) {
      return false;
    }
  }
  
  return true;
}

// ========================================
// 日付が同じか判定
// ========================================
function isSameDay(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

// LINE通知処理は家計簿受信.gsの pushToDiscord を使用するため削除

// ========================================
// テスト用関数
// ========================================
function testRecurringPayments() {
  executeRecurringPayments();
}
