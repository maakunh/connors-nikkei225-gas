// ============================================================
//  日経225 ラリー・コナーズ戦略 自動分析 (Google Apps Script)
//  現物(日中) + EWJ連動(夜間) デュアル版
//
//  【セットアップ手順】
//  1. Google スプレッドシートを新規作成
//  2. メニュー「拡張機能」→「Apps Script」を開く
//  3. このコードを全てコピーして貼り付け
//  4. 保存（Ctrl+S）後、関数 setupSheet を実行
//  5. 初回は権限の承認が必要（「許可を確認」→ Googleアカウント選択）
//  6. 関数 setupDailyTrigger を実行
//     → 毎営業日 16:30 に日中更新、翌 06:00 に夜間更新が自動実行されます
//
//  【手動実行】
//  - dailyUpdateDaytime   : 日中分(現物)のみ更新
//  - dailyUpdateOvernight : 夜間分(EWJ連動)のみ更新
//  - dailyUpdate          : 両方まとめて更新
// ============================================================

// ==================== 設定 ====================
const CONFIG = {
  SHEET_NAME: "コナーズ分析",
  TICKER: "^N225",             // Yahoo Finance の日経225ティッカー
  DATA_ROWS_NEEDED: 280,       // 1年分(約250営業日) + RSI計算バッファ
  RSI_SHORT: 2,                // RSI(2)
  RSI_MID: 4,                  // RSI(4)
  SMA_PERIOD: 5,               // 5日移動平均
  SMA_LONG: 200,               // 200日MA（トレンドフィルター）

  // ── トリガー設定 ──
  DAYTIME_HOUR: 16,            // 日中実行時刻（時）
  DAYTIME_MINUTE: 30,          // 日中実行時刻（分）
  OVERNIGHT_HOUR: 6,           // 夜間実行時刻（時）※翌朝
  OVERNIGHT_MINUTE: 0,         // 夜間実行時刻（分）

  // ── メール通知設定 ──
  NOTIFY_EMAIL: true,          // true: メール送信する / false: しない
  EMAIL_TO: "",                // 送信先メールアドレス（空欄＝自分のGmailに送信）

  // ── 夜間データ取得 ──
  // CME/SGX の日経225先物は GOOGLEFINANCE が対応していないため、
  // 米国上場の日本株ETF NYSEARCA:EWJ を代替として使用。
  // EWJ価格に日経225/EWJ比率(中央値)を掛けて日経225換算値に変換する。
};

// ==================== 初期セットアップ ====================
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  // ── 2段ヘッダー ──
  // 1行目: グループヘッダー (日付 / 日中(現物) / 夜間(先物))
  // 2行目: 項目ヘッダー
  const groupHeaders = [
    "日付",
    "日中 (現物)", "", "", "", "", "", "", "", "",
    "夜間 (EWJ連動)", "", "", "", "", ""
  ];
  const subHeaders = [
    "",
    "終値", "前日比", "前日比%", "RSI(2)", "RSI(4)", "連騰", "5日MA", "200日MA", "シグナル",
    "終値", "前日比", "前日比%", "RSI(2)", "連騰", "シグナル"
  ];

  const totalCols = subHeaders.length;

  sheet.getRange(1, 1, 1, totalCols).setValues([groupHeaders]);
  sheet.getRange(2, 1, 1, totalCols).setValues([subHeaders]);

  // 日付列を縦結合
  sheet.getRange(1, 1, 2, 1).merge();
  // 日中グループ: B1:J1 (列2〜10)
  sheet.getRange(1, 2, 1, 9).merge();
  // 夜間グループ: K1:P1 (列11〜16)
  sheet.getRange(1, 11, 1, 6).merge();

  // ── ヘッダーのスタイル ──
  const row1 = sheet.getRange(1, 1, 1, totalCols);
  row1.setFontWeight("bold").setFontSize(11)
      .setBackground("#1a1a2e").setFontColor("#ffffff")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

  const row2 = sheet.getRange(2, 1, 1, totalCols);
  row2.setFontWeight("bold").setFontSize(10)
      .setBackground("#2a2a40").setFontColor("#e0e0e0")
      .setHorizontalAlignment("center");

  // 日中グループに青系アクセント
  sheet.getRange(1, 2, 1, 9).setBackground("#1e3a5f");
  // 夜間グループに紫系アクセント
  sheet.getRange(1, 11, 1, 6).setBackground("#3b1e5f");

  // ── 列幅調整 ──
  const colWidths = [
    95,  // A 日付
    // 日中 (9列)
    100, 90, 75, 65, 65, 55, 100, 100, 90,
    // 夜間 (6列)
    100, 90, 75, 65, 55, 90
  ];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // ── 行固定 ──
  sheet.setFrozenRows(2);
  // 日付列も固定
  sheet.setFrozenColumns(1);

  // ── 初回データ取得 ──
  dailyUpdate();

  SpreadsheetApp.getUi().alert(
    "セットアップ完了！\n\n" +
    "毎日の自動実行を設定するには、\n" +
    "関数「setupDailyTrigger」を実行してください。\n\n" +
    "日中 16:30 → 現物データで更新\n" +
    "翌朝 06:00 → EWJ連動で夜間分を追記"
  );
}

// ==================== 日次トリガー設定 ====================
function setupDailyTrigger() {
  // 既存の dailyUpdate系 トリガー削除
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn === "dailyUpdate" || fn === "dailyUpdateDaytime" || fn === "dailyUpdateOvernight") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 日中: 16:30 (現物終値+リアルタイム)
  ScriptApp.newTrigger("dailyUpdateDaytime")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.DAYTIME_HOUR)
    .nearMinute(CONFIG.DAYTIME_MINUTE)
    .inTimezone("Asia/Tokyo")
    .create();

  // 夜間: 翌朝 06:00 (EWJ連動)
  ScriptApp.newTrigger("dailyUpdateOvernight")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.OVERNIGHT_HOUR)
    .nearMinute(CONFIG.OVERNIGHT_MINUTE)
    .inTimezone("Asia/Tokyo")
    .create();

  SpreadsheetApp.getUi().alert(
    "日次トリガー設定完了！\n\n" +
    `日中  ${CONFIG.DAYTIME_HOUR}:${String(CONFIG.DAYTIME_MINUTE).padStart(2, '0')} (JST) — 現物データで更新\n` +
    `夜間  ${CONFIG.OVERNIGHT_HOUR}:${String(CONFIG.OVERNIGHT_MINUTE).padStart(2, '0')} (JST) — EWJ連動で夜間分を追記\n\n` +
    "※ 土日祝日も実行されますが、データがない場合はスキップされます。"
  );
}

// ==================== メインの更新関数 ====================

// 日中分(現物)と夜間分(EWJ連動)を両方更新
function dailyUpdate() {
  dailyUpdateDaytime();
  dailyUpdateOvernight();
}

// 日中更新: 現物データを取得してシートを再構築
function dailyUpdateDaytime() {
  try {
    Logger.log("=== 日中更新開始 (現物) ===");
    const prices = fetchNikkeiData_();
    if (!prices || prices.length < CONFIG.SMA_PERIOD + 1) {
      Logger.log("現物データ不足: " + (prices ? prices.length : 0) + "件");
      return;
    }

    const daytimeAnalysis = analyzeConnors_(prices);
    writeDaytimeToSheet_(daytimeAnalysis);
    applyFormatting_();

    Logger.log("日中更新完了: " + new Date().toLocaleString("ja-JP"));

    if (CONFIG.NOTIFY_EMAIL) {
      sendNotificationEmail_(daytimeAnalysis, "daytime");
    }
  } catch (e) {
    Logger.log("日中更新エラー: " + e.message);
    if (CONFIG.NOTIFY_EMAIL) {
      sendErrorEmail_("[日中] " + e.message);
    }
  }
}

// 夜間更新: EWJ連動データを取得してシートの夜間列を更新
//
// 方針:
// - 過去日: EWJ履歴を日経225換算した値で連騰・RSI を計算
// - 当日: EWJリアルタイム換算値を最新日として追加し、その上で連騰・RSI を計算
// - カレンダーは EWJ ベース (NYSE 営業日)
function dailyUpdateOvernight() {
  try {
    Logger.log("=== 夜間更新開始 (EWJ連動) ===");

    // EWJ系列を取得（履歴 + リアルタイム値）
    const ewjPrices = fetchFuturesData_();
    if (!ewjPrices || ewjPrices.length < CONFIG.SMA_PERIOD + 1) {
      Logger.log("EWJデータ不足: " + (ewjPrices ? ewjPrices.length : 0) + "件");
      return;
    }

    // EWJシリーズを構築（履歴 + リアルタイム値を最新日として追加）
    const overnightSeries = ewjPrices.map(p => ({ date: p.date, close: p.close }));

    if (ewjPrices.realtimePrice !== null && ewjPrices.realtimePrice !== undefined) {
      // リアルタイム値を「今日」として追加または上書き
      const now = new Date();
      const dateKey = (d) => Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
      const todayKey = dateKey(now);
      const lastKey = dateKey(overnightSeries[overnightSeries.length - 1].date);

      if (lastKey === todayKey) {
        // 既に今日の履歴があれば置き換え
        const orig = overnightSeries[overnightSeries.length - 1].close;
        overnightSeries[overnightSeries.length - 1] = { date: now, close: ewjPrices.realtimePrice };
        Logger.log("当日(" + todayKey + ")の終値をリアルタイム値で上書き: " + orig + " → " + ewjPrices.realtimePrice);
      } else {
        overnightSeries.push({ date: now, close: ewjPrices.realtimePrice });
        Logger.log("当日(" + todayKey + ")をリアルタイム値で追加: " + ewjPrices.realtimePrice);
      }
    } else {
      Logger.log("EWJリアルタイム値が取得できなかったため、履歴のみで分析");
    }

    // 連騰・RSI・シグナルを計算
    const overnightAnalysis = analyzeConnors_(overnightSeries);

    // シートに書き込み（既存日中データと和集合）
    writeOvernightToSheet_(overnightAnalysis);
    applyFormatting_();

    Logger.log("夜間更新完了: " + new Date().toLocaleString("ja-JP"));

    if (CONFIG.NOTIFY_EMAIL) {
      sendNotificationEmail_(overnightAnalysis, "overnight");
    }
  } catch (e) {
    Logger.log("夜間更新エラー: " + e.message);
    if (CONFIG.NOTIFY_EMAIL) {
      sendErrorEmail_("[夜間] " + e.message);
    }
  }
}

// ==================== データ取得 ====================
function fetchNikkeiData_() {
  // 方法1: Google Finance 履歴 + リアルタイム結合（メイン）
  Logger.log("Google Finance (履歴+リアルタイム結合) でデータ取得を試行...");
  try {
    const prices = fetchFromGoogleFinanceWithRealtime_();
    if (prices && prices.length >= CONFIG.SMA_PERIOD + 1) {
      const latest = Utilities.formatDate(
        prices[prices.length - 1].date,
        "Asia/Tokyo", "yyyy/MM/dd"
      );
      Logger.log("Google Finance 結合: " + prices.length + "件取得成功 (最新: " + latest + ")");
      return prices;
    }
  } catch (e) {
    Logger.log("Google Finance 結合 失敗: " + e.message);
  }

  // 方法2: Yahoo Finance API（最終フォールバック）
  Logger.log("Yahoo Finance API でデータ取得を試行...");
  try {
    const yfPrices = fetchFromYahooFinance_();
    if (yfPrices && yfPrices.length >= CONFIG.SMA_PERIOD + 1) {
      Logger.log("Yahoo Finance: " + yfPrices.length + "件取得成功");
      return yfPrices;
    }
  } catch (e) {
    Logger.log("Yahoo Finance 失敗: " + e.message);
  }

  Logger.log("全てのデータソースが失敗しました");
  return null;
}

// Google Finance の履歴データとリアルタイム現在値を結合して返す
// 履歴モード (close, start, end) は当日データの反映が翌朝まで遅れるが、
// リアルタイムモード (price) は当日の現在値を即取得できる（15〜20分遅延）
// 両者を組み合わせることで、当日分も含めたデータを確実に取得する
function fetchFromGoogleFinanceWithRealtime_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempSheet = ss.getSheetByName("_temp_data");

  if (!tempSheet) {
    tempSheet = ss.insertSheet("_temp_data");
  } else {
    tempSheet.showSheet();
  }
  tempSheet.clear();

  const days = CONFIG.DATA_ROWS_NEEDED * 2;

  // A列: 履歴データ
  tempSheet.getRange("A1").setFormula(
    `=GOOGLEFINANCE("INDEXNIKKEI:NI225","close",TODAY()-${days},TODAY())`
  );
  // E1: リアルタイム現在値 (単一の数値を返す)
  tempSheet.getRange("E1").setFormula(
    `=GOOGLEFINANCE("INDEXNIKKEI:NI225","price")`
  );
  // F1: リアルタイム取得の最終更新時刻
  tempSheet.getRange("F1").setFormula(
    `=GOOGLEFINANCE("INDEXNIKKEI:NI225","datadelay")`
  );

  SpreadsheetApp.flush();

  // 履歴データが反映されるまで待機
  let histData = [];
  const maxRetries = 5;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    Utilities.sleep(2000 * attempt);
    SpreadsheetApp.flush();

    const allData = tempSheet.getDataRange().getValues();
    histData = [];

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] instanceof Date && typeof allData[i][1] === "number") {
        histData.push({
          date: allData[i][0],
          close: Math.round(allData[i][1] * 100) / 100
        });
      }
    }

    if (histData.length > 5) {
      Logger.log("  履歴: 試行 " + attempt + " 回目で " + histData.length + " 件取得");
      break;
    }
    Logger.log("  履歴: 試行 " + attempt + " - " + histData.length + " 件");
  }

  if (histData.length === 0) {
    try { tempSheet.hideSheet(); } catch (e) {}
    throw new Error("履歴データが取得できませんでした");
  }

  // ソート
  histData.sort((a, b) => a.date.getTime() - b.date.getTime());

  // ── リアルタイム現在値を取得 ──
  const realtimePrice = tempSheet.getRange("E1").getValue();
  Logger.log("  リアルタイム値: " + realtimePrice);

  // 履歴の最新日と今日を比較
  const now = new Date();
  const todayKey = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
  const latestHistKey = Utilities.formatDate(
    histData[histData.length - 1].date,
    "Asia/Tokyo", "yyyy-MM-dd"
  );

  // 履歴に当日が含まれていない && リアルタイム値が有効な数値 && 平日 && 東京市場営業時間中or大引け後
  const dayOfWeek = now.getDay();  // 0=日, 6=土
  const isWeekday = dayOfWeek >= 1 && dayOfWeek <= 5;
  const isValidRealtime = typeof realtimePrice === "number" && realtimePrice > 0;
  const hour = parseInt(Utilities.formatDate(now, "Asia/Tokyo", "H"), 10);
  const isAfterMarketStart = hour >= 9; // 9時以降ならザラ場中or大引け後

  if (latestHistKey !== todayKey && isValidRealtime && isWeekday && isAfterMarketStart) {
    // 当日のリアルタイム値を追加
    histData.push({
      date: now,
      close: Math.round(realtimePrice * 100) / 100
    });
    Logger.log("  当日データをリアルタイム値で追加: " + todayKey + " @ " + realtimePrice);
  } else {
    Logger.log("  リアルタイム追加スキップ: 履歴最新=" + latestHistKey + " 今日=" + todayKey +
      " 平日=" + isWeekday + " 有効値=" + isValidRealtime);
  }

  try { tempSheet.hideSheet(); } catch (e) {}

  return histData;
}

// ==================== 先物データ取得 ====================
// 夜間(日本時間の大引け後～翌朝)の日経225先物を取得
// CME:NKD (円建て日経225先物) または CME:NIY (ドル建て) を試行
// GOOGLEFINANCEの履歴は先物に対応していないので、過去分は現物(履歴)を使い、
// 最新の1日分だけ先物のリアルタイム値で上書きする
// ==================== 夜間データ取得 (EWJ ベース) ====================
//
// CME/SGX の日経225先物は GOOGLEFINANCE が対応していないため、
// 米国上場の日本株ETF NYSEARCA:EWJ を代替指標として使用する。
// EWJ は米国時間に取引されるため、東京市場の大引け後〜翌朝の動きを反映する。
//
// アプローチ:
// 1. EWJ の履歴(日足)を GOOGLEFINANCE で取得
// 2. 日経225指数の履歴も取得（スケール係数算出のため）
// 3. 両方に存在する日付の比率の中央値をスケール係数として算出
// 4. 各 EWJ 価格にスケール係数を掛けて「日経225換算値」に変換
// 5. EWJ のリアルタイム値も同様に変換して当日分として末尾に追加
function fetchFuturesData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempSheet = ss.getSheetByName("_temp_ewj");

  if (!tempSheet) {
    tempSheet = ss.insertSheet("_temp_ewj");
  } else {
    try { tempSheet.showSheet(); } catch (e) {}
  }
  tempSheet.clear();

  const days = CONFIG.DATA_ROWS_NEEDED * 2;

  // A列: EWJ 履歴 / E列: 日経225指数履歴 / I1: EWJ リアルタイム値
  tempSheet.getRange("A1").setFormula(
    `=GOOGLEFINANCE("NYSEARCA:EWJ","close",TODAY()-${days},TODAY())`
  );
  tempSheet.getRange("E1").setFormula(
    `=GOOGLEFINANCE("INDEXNIKKEI:NI225","close",TODAY()-${days},TODAY())`
  );
  tempSheet.getRange("I1").setFormula(
    `=GOOGLEFINANCE("NYSEARCA:EWJ","price")`
  );

  SpreadsheetApp.flush();

  let ewjData = [];
  let indexData = [];
  const maxRetries = 5;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    Utilities.sleep(2000 * attempt);
    SpreadsheetApp.flush();

    const allData = tempSheet.getDataRange().getValues();
    ewjData = [];
    indexData = [];

    for (let i = 1; i < allData.length; i++) {
      // A列: EWJ 日付, B列: EWJ 終値
      if (allData[i][0] instanceof Date && typeof allData[i][1] === "number") {
        ewjData.push({ date: allData[i][0], close: allData[i][1] });
      }
      // E列: 指数 日付, F列: 指数 終値
      if (allData[i][4] instanceof Date && typeof allData[i][5] === "number") {
        indexData.push({ date: allData[i][4], close: allData[i][5] });
      }
    }

    if (ewjData.length > 5) {
      Logger.log("  EWJ: 試行 " + attempt + " 回目 → EWJ " + ewjData.length + "件, 指数 " + indexData.length + "件");
      break;
    }
    Logger.log("  EWJ: 試行 " + attempt + " - EWJ " + ewjData.length + "件");
  }

  if (ewjData.length === 0) {
    try { tempSheet.hideSheet(); } catch (e) {}
    throw new Error("EWJ データが取得できませんでした");
  }

  // ソート（古い順）
  ewjData.sort((a, b) => a.date.getTime() - b.date.getTime());
  indexData.sort((a, b) => a.date.getTime() - b.date.getTime());

  // ── スケール係数の算出 ──
  // 日経225指数とEWJの両方にデータがある日の比率を計算
  const dateKey = (d) => Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
  const indexMap = {};
  indexData.forEach(d => { indexMap[dateKey(d.date)] = d.close; });

  const ratios = [];
  ewjData.forEach(d => {
    const key = dateKey(d.date);
    if (indexMap[key]) {
      ratios.push({ key: key, ratio: indexMap[key] / d.close });
    }
  });

  if (ratios.length === 0) {
    try { tempSheet.hideSheet(); } catch (e) {}
    throw new Error("スケール係数が算出できませんでした（指数データ不足）");
  }

  // 直近5日の比率の中央値を採用
  // 全期間の中央値だと為替ドリフト(ドル円の長期変動)で乖離するため、
  // 直近の比率だけを使うことで為替影響を最小化する
  const SCALE_SAMPLE_DAYS = 5;
  const recentRatios = ratios.slice(-SCALE_SAMPLE_DAYS).map(r => r.ratio);
  recentRatios.sort((a, b) => a - b);
  const scale = recentRatios[Math.floor(recentRatios.length / 2)];
  Logger.log("  EWJ→日経225 スケール係数: " + scale.toFixed(4) +
    " (直近" + recentRatios.length + "日, 全マッチ" + ratios.length + "日)");

  // EWJ価格を日経225水準に変換
  const prices = ewjData.map(d => ({
    date: d.date,
    close: Math.round(d.close * scale * 100) / 100
  }));

  // ── EWJのリアルタイム値を取得して保持 ──
  // 当日として追加せず、メタ情報として返す
  // (writeOvernightToSheet_ 側で日中シートの最新行に紐付ける)
  const ewjRealtime = tempSheet.getRange("I1").getValue();
  Logger.log("  EWJ リアルタイム値: " + ewjRealtime);

  let realtimePrice = null;
  if (typeof ewjRealtime === "number" && ewjRealtime > 0) {
    realtimePrice = Math.round(ewjRealtime * scale * 100) / 100;
    Logger.log("  リアルタイム換算値: " + realtimePrice + " (EWJ生値: " + ewjRealtime + ")");
  }

  try { tempSheet.hideSheet(); } catch (e) {}

  Logger.log("  夜間データ最終: " + prices.length + "件 (最新EWJ日: " + dateKey(prices[prices.length - 1].date) + ")");

  // prices に realtimePrice をプロパティとして付与して返す
  prices.realtimePrice = realtimePrice;
  return prices;
}

// Yahoo Finance API からデータ取得
function fetchFromYahooFinance_() {
  const now = Math.floor(Date.now() / 1000);
  const daysBack = CONFIG.DATA_ROWS_NEEDED * 2;
  const from = now - daysBack * 86400;

  // ティッカーをURLエンコード（ ^ → %5E）
  const encodedTicker = encodeURIComponent(CONFIG.TICKER);

  // 複数のエンドポイントを試行
  const urls = [
    `https://query2.finance.yahoo.com/v8/finance/chart/${encodedTicker}?period1=${from}&period2=${now}&interval=1d&includePrePost=false`,
    `https://query1.finance.yahoo.com/v8/finance/chart/${encodedTicker}?period1=${from}&period2=${now}&interval=1d&includePrePost=false`,
  ];

  const options = {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    },
    followRedirects: true
  };

  for (const url of urls) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      Logger.log("Yahoo API レスポンス: " + code + " URL: " + url.substring(0, 60) + "...");

      if (code !== 200) continue;

      const json = JSON.parse(response.getContentText());

      if (!json.chart || !json.chart.result || !json.chart.result[0]) continue;

      const result = json.chart.result[0];
      if (!result.timestamp || !result.indicators || !result.indicators.quote) continue;

      const timestamps = result.timestamp;
      const closes = result.indicators.quote[0].close;

      const prices = [];
      for (let i = 0; i < timestamps.length; i++) {
        if (closes[i] !== null && closes[i] !== undefined) {
          prices.push({
            date: new Date(timestamps[i] * 1000),
            close: Math.round(closes[i] * 100) / 100
          });
        }
      }

      if (prices.length > 0) return prices;
    } catch (e) {
      Logger.log("Yahoo endpoint 失敗: " + e.message);
      continue;
    }
  }

  return null;
}

// Google Finance 代替取得
function fetchFromGoogleFinance_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempSheet = ss.getSheetByName("_temp_data");
  
  if (!tempSheet) {
    tempSheet = ss.insertSheet("_temp_data");
  } else {
    tempSheet.showSheet();
  }
  tempSheet.clear();
  
  // GOOGLEFINANCE関数で日経平均を取得
  const days = CONFIG.DATA_ROWS_NEEDED * 2;
  const formula = `=GOOGLEFINANCE("INDEXNIKKEI:NI225","close",TODAY()-${days},TODAY())`;
  tempSheet.getRange("A1").setFormula(formula);
  
  // データ反映を待つ（GOOGLEFINANCE は非同期で遅延する場合がある）
  SpreadsheetApp.flush();
  
  let data = [];
  const maxRetries = 5;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    Utilities.sleep(2000 * attempt); // 2秒, 4秒, 6秒... と待機を延長
    SpreadsheetApp.flush();
    data = tempSheet.getDataRange().getValues();
    
    // 2行以上あればデータ取得成功（1行目ヘッダー + データ行）
    if (data.length > 5) {
      Logger.log("Google Finance: 試行 " + attempt + " 回目で " + (data.length - 1) + " 行取得");
      break;
    }
    Logger.log("Google Finance: 試行 " + attempt + " - まだ " + data.length + " 行のみ");
  }
  
  const prices = [];
  
  // 1行目はヘッダーなのでスキップ
  for (let i = 1; i < data.length; i++) {
    const dateVal = data[i][0];
    const closeVal = data[i][1];
    
    // 日付と数値の両方が有効な行のみ取り込む
    if (dateVal instanceof Date && typeof closeVal === "number" && !isNaN(closeVal) && closeVal > 0) {
      prices.push({
        date: dateVal,
        close: Math.round(closeVal * 100) / 100
      });
    }
  }
  
  // 一時シートを非表示に
  try { tempSheet.hideSheet(); } catch(e) { /* シートが1枚だけの場合は非表示にできない */ }
  
  return prices;
}

// ==================== コナーズ分析 ====================
function analyzeConnors_(prices) {
  const n = prices.length;
  
  // 前日比を計算
  const changes = [0];
  for (let i = 1; i < n; i++) {
    changes.push(prices[i].close - prices[i - 1].close);
  }
  
  // RSI計算 (Wilder's Smoothing)
  const rsi2 = calcRSI_(changes, CONFIG.RSI_SHORT);
  const rsi4 = calcRSI_(changes, CONFIG.RSI_MID);
  
  // 連騰日数
  const streaks = calcStreak_(changes);
  
  // 5日移動平均
  const sma5 = calcSMA_(prices, CONFIG.SMA_PERIOD);
  
  // 200日移動平均
  const sma200 = calcSMA_(prices, Math.min(CONFIG.SMA_LONG, n));
  
  // 直近の分析対象期間（約1年分 = 250営業日）
  const displayCount = Math.min(250, n - CONFIG.SMA_PERIOD);
  const startIdx = n - displayCount;
  
  const results = [];
  for (let i = startIdx; i < n; i++) {
    const changePct = i > 0 ? (changes[i] / prices[i - 1].close) * 100 : 0;
    
    results.push({
      date: prices[i].date,
      close: prices[i].close,
      change: changes[i],
      changePct: changePct,
      rsi2: rsi2[i],
      rsi4: rsi4[i],
      streak: streaks[i],
      sma5: sma5[i],
      sma200: sma200[i],
      signal: getSignal_(rsi2[i], rsi4[i], streaks[i], prices[i].close, sma200[i])
    });
  }
  
  return results;
}

// RSI計算 (Wilder's Smoothing Method)
function calcRSI_(changes, period) {
  const n = changes.length;
  const result = new Array(n).fill(null);
  
  if (n < period + 1) return result;
  
  let gainSum = 0, lossSum = 0;
  for (let i = 1; i <= period; i++) {
    if (changes[i] > 0) gainSum += changes[i];
    else lossSum += Math.abs(changes[i]);
  }
  
  let avgGain = gainSum / period;
  let avgLoss = lossSum / period;
  
  result[period] = avgLoss === 0 ? 100 : 100 - 100 / (1 + avgGain / avgLoss);
  
  for (let i = period + 1; i < n; i++) {
    const gain = changes[i] > 0 ? changes[i] : 0;
    const loss = changes[i] < 0 ? Math.abs(changes[i]) : 0;
    avgGain = (avgGain * (period - 1) + gain) / period;
    avgLoss = (avgLoss * (period - 1) + loss) / period;
    result[i] = avgLoss === 0 ? 100 : 100 - 100 / (1 + avgGain / avgLoss);
  }
  
  return result;
}

// 連騰日数
function calcStreak_(changes) {
  const n = changes.length;
  const result = new Array(n).fill(0);
  
  for (let i = 1; i < n; i++) {
    if (changes[i] > 0) {
      result[i] = result[i - 1] > 0 ? result[i - 1] + 1 : 1;
    } else if (changes[i] < 0) {
      result[i] = result[i - 1] < 0 ? result[i - 1] - 1 : -1;
    }
  }
  
  return result;
}

// 単純移動平均
function calcSMA_(prices, period) {
  const n = prices.length;
  const result = new Array(n).fill(null);
  
  for (let i = period - 1; i < n; i++) {
    let sum = 0;
    for (let j = i - period + 1; j <= i; j++) {
      sum += prices[j].close;
    }
    result[i] = Math.round(sum / period * 100) / 100;
  }
  
  return result;
}

// コナーズ戦略シグナル判定
function getSignal_(rsi2, rsi4, streak, close, sma200) {
  if (rsi2 === null) return "—";
  
  // 200日MAの上にいるかどうか（トレンドフィルター）
  const aboveTrend = sma200 === null || close > sma200;
  
  // ── 買いシグナル（上昇トレンド中のみ） ──
  if (aboveTrend) {
    if (rsi2 < 5 && streak <= -3)  return "★強い買い";
    if (rsi2 < 5 && streak <= -2)  return "◎買い";
    if (rsi2 < 10 && streak <= -2) return "○買い";
    if (rsi2 < 25 && streak <= -2) return "△やや買い";
  }
  
  // ── 売りシグナル（利確・エグジット） ──
  if (rsi2 > 95 && streak >= 3)  return "★強い売り";
  if (rsi2 > 90 && streak >= 2)  return "◎売り";
  if (rsi2 > 75 && streak >= 2)  return "△やや売り";
  
  // ── 警告（下降トレンド中の買いシグナル相当） ──
  if (!aboveTrend && rsi2 < 10 && streak <= -2) return "⚠トレンド注意";
  
  return "−中立";
}

// ==================== シートへの書き込み ====================
//
// 列構成:
//  A(1)  : 日付
//  日中 (現物) — B〜J (列2〜10)
//  B(2)  : 終値   C(3): 前日比   D(4): 前日比%
//  E(5)  : RSI(2) F(6): RSI(4)   G(7): 連騰
//  H(8)  : 5日MA  I(9): 200日MA  J(10): シグナル
//  夜間 (EWJ連動) — K〜P (列11〜16)
//  K(11) : 終値   L(12): 前日比  M(13): 前日比%
//  N(14) : RSI(2) O(15): 連騰   P(16): シグナル

const TOTAL_COLS = 16;
const DATA_START_ROW = 3; // 1,2行目はヘッダー

function writeDaytimeToSheet_(results) {
  // 日中データを更新する場合: 既存の夜間データを読み出して保持し、和集合で全書き直し
  const existing = readExistingData_();
  rewriteSheetUnion_(results, existing.overnight);
}

function writeOvernightToSheet_(results) {
  // 夜間データを更新する場合: 既存の日中データを読み出して保持し、和集合で全書き直し
  const existing = readExistingData_();
  rewriteSheetUnion_(existing.daytime, results);
}

// シートから既存のデータを読み出し、日中と夜間それぞれの配列に分解
function readExistingData_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    return { daytime: [], overnight: [] };
  }

  const rowCount = lastRow - DATA_START_ROW + 1;
  const allData = sheet.getRange(DATA_START_ROW, 1, rowCount, TOTAL_COLS).getValues();

  const daytime = [];
  const overnight = [];

  allData.forEach(row => {
    const date = row[0];
    if (!(date instanceof Date)) return;

    // 日中(列B〜J = idx 1〜9): 終値が数値ならデータ有り
    if (typeof row[1] === "number") {
      daytime.push({
        date: date,
        close: row[1],
        change: row[2],
        changePct: typeof row[3] === "number" ? row[3] * 100 : 0,
        rsi2: typeof row[4] === "number" ? row[4] : null,
        rsi4: typeof row[5] === "number" ? row[5] : null,
        streak: row[6],
        sma5: typeof row[7] === "number" ? row[7] : null,
        sma200: typeof row[8] === "number" ? row[8] : null,
        signal: row[9] || ""
      });
    }

    // 夜間(列K〜P = idx 10〜15): 終値が数値ならデータ有り
    if (typeof row[10] === "number") {
      overnight.push({
        date: date,
        close: row[10],
        change: row[11],
        changePct: typeof row[12] === "number" ? row[12] * 100 : 0,
        rsi2: typeof row[13] === "number" ? row[13] : null,
        rsi4: null,
        streak: row[14],
        sma5: null,
        sma200: null,
        signal: row[15] || ""
      });
    }
  });

  return { daytime: daytime, overnight: overnight };
}

// 日中と夜間の両配列から日付の和集合を作り、シートを全書き直し
function rewriteSheetUnion_(daytimeResults, overnightResults) {
  const sheet = getSheet_();
  const dateKey = (d) => Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");

  // 日付→データ のマッピング
  const daytimeMap = {};
  (daytimeResults || []).forEach(r => { daytimeMap[dateKey(r.date)] = r; });
  const overnightMap = {};
  (overnightResults || []).forEach(r => { overnightMap[dateKey(r.date)] = r; });

  // 和集合の日付セット
  const allKeys = {};
  Object.keys(daytimeMap).forEach(k => { allKeys[k] = true; });
  Object.keys(overnightMap).forEach(k => { allKeys[k] = true; });

  const sortedKeys = Object.keys(allKeys).sort();

  // 既存データ部分をクリア
  if (sheet.getLastRow() >= DATA_START_ROW) {
    sheet.getRange(DATA_START_ROW, 1, sheet.getLastRow() - DATA_START_ROW + 1, TOTAL_COLS).clear();
  }

  if (sortedKeys.length === 0) {
    Logger.log("書き込みデータなし");
    return;
  }

  // 行を構築
  const rows = sortedKeys.map(key => {
    const d = daytimeMap[key];
    const n = overnightMap[key];

    // 日付は日中・夜間どちらかの Date を使う
    const dateObj = (d && d.date) || (n && n.date);

    return [
      dateObj,
      // 日中 9列
      d ? d.close : "",
      d ? d.change : "",
      d ? (d.changePct / 100) : "",
      d && d.rsi2 !== null ? Math.round(d.rsi2 * 10) / 10 : "",
      d && d.rsi4 !== null ? Math.round(d.rsi4 * 10) / 10 : "",
      d ? d.streak : "",
      d && d.sma5 ? d.sma5 : "",
      d && d.sma200 ? d.sma200 : "",
      d ? d.signal : "",
      // 夜間 6列
      n ? n.close : "",
      n ? n.change : "",
      n ? (n.changePct / 100) : "",
      n && n.rsi2 !== null ? Math.round(n.rsi2 * 10) / 10 : "",
      n ? n.streak : "",
      n ? n.signal : ""
    ];
  });

  sheet.getRange(DATA_START_ROW, 1, rows.length, TOTAL_COLS).setValues(rows);
  Logger.log("和集合書き込み: " + rows.length + "行 (日中=" +
    Object.keys(daytimeMap).length + ", 夜間=" + Object.keys(overnightMap).length + ")");
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  return sheet;
}

// ==================== 書式設定 ====================
function applyFormatting_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  const rowCount = lastRow - DATA_START_ROW + 1;

  // ── 数値フォーマット ──
  // 日付
  sheet.getRange(DATA_START_ROW, 1, rowCount, 1).setNumberFormat("yyyy/mm/dd");
  // 日中: 終値
  sheet.getRange(DATA_START_ROW, 2, rowCount, 1).setNumberFormat("#,##0.00");
  // 日中: 前日比
  sheet.getRange(DATA_START_ROW, 3, rowCount, 1).setNumberFormat("+#,##0.00;-#,##0.00;0");
  // 日中: 前日比%
  sheet.getRange(DATA_START_ROW, 4, rowCount, 1).setNumberFormat("+0.00%;-0.00%;0.00%");
  // 日中: RSI(2)(4)
  sheet.getRange(DATA_START_ROW, 5, rowCount, 2).setNumberFormat("0.0");
  // 日中: 連騰
  sheet.getRange(DATA_START_ROW, 7, rowCount, 1).setNumberFormat("+0;-0;0");
  // 日中: 5日MA, 200日MA
  sheet.getRange(DATA_START_ROW, 8, rowCount, 2).setNumberFormat("#,##0.00");

  // 夜間: 終値
  sheet.getRange(DATA_START_ROW, 11, rowCount, 1).setNumberFormat("#,##0.00");
  // 夜間: 前日比
  sheet.getRange(DATA_START_ROW, 12, rowCount, 1).setNumberFormat("+#,##0.00;-#,##0.00;0");
  // 夜間: 前日比%
  sheet.getRange(DATA_START_ROW, 13, rowCount, 1).setNumberFormat("+0.00%;-0.00%;0.00%");
  // 夜間: RSI(2)
  sheet.getRange(DATA_START_ROW, 14, rowCount, 1).setNumberFormat("0.0");
  // 夜間: 連騰
  sheet.getRange(DATA_START_ROW, 15, rowCount, 1).setNumberFormat("+0;-0;0");

  // ── アライメント ──
  sheet.getRange(DATA_START_ROW, 1, rowCount, 1).setHorizontalAlignment("center");
  sheet.getRange(DATA_START_ROW, 2, rowCount, 8).setHorizontalAlignment("right");
  sheet.getRange(DATA_START_ROW, 10, rowCount, 1).setHorizontalAlignment("center");
  sheet.getRange(DATA_START_ROW, 11, rowCount, 5).setHorizontalAlignment("right");
  sheet.getRange(DATA_START_ROW, 16, rowCount, 1).setHorizontalAlignment("center");

  // 夜間列にうっすら背景色（判別しやすく）
  sheet.getRange(DATA_START_ROW, 11, rowCount, 6).setBackground("#f5f0ff");

  // ── 条件付き書式 ──
  sheet.clearConditionalFormatRules();
  const rules = [];

  // 日中 前日比
  const dayChangeRange = sheet.getRange(DATA_START_ROW, 3, rowCount, 2);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setFontColor("#00875a")
      .setRanges([dayChangeRange]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor("#de350b")
      .setRanges([dayChangeRange]).build()
  );

  // 夜間 前日比
  const nightChangeRange = sheet.getRange(DATA_START_ROW, 12, rowCount, 2);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0).setFontColor("#00875a")
      .setRanges([nightChangeRange]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor("#de350b")
      .setRanges([nightChangeRange]).build()
  );

  // 日中 RSI(2)
  const dayRSIRange = sheet.getRange(DATA_START_ROW, 5, rowCount, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(10).setBackground("#c6efce").setFontColor("#006100")
      .setRanges([dayRSIRange]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(90).setBackground("#ffc7ce").setFontColor("#9c0006")
      .setRanges([dayRSIRange]).build()
  );

  // 夜間 RSI(2)
  const nightRSIRange = sheet.getRange(DATA_START_ROW, 14, rowCount, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(10).setBackground("#c6efce").setFontColor("#006100")
      .setRanges([nightRSIRange]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(90).setBackground("#ffc7ce").setFontColor("#9c0006")
      .setRanges([nightRSIRange]).build()
  );

  // シグナル列 (日中J=10, 夜間P=16)
  [10, 16].forEach(col => {
    const sigRange = sheet.getRange(DATA_START_ROW, col, rowCount, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("買い").setBackground("#c6efce").setFontColor("#006100").setBold(true)
        .setRanges([sigRange]).build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("売り").setBackground("#ffc7ce").setFontColor("#9c0006").setBold(true)
        .setRanges([sigRange]).build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("注意").setBackground("#fff2cc").setFontColor("#7f6003").setBold(true)
        .setRanges([sigRange]).build()
    );
  });

  sheet.setConditionalFormatRules(rules);

  // 最新行をハイライト
  sheet.getRange(lastRow, 1, 1, TOTAL_COLS).setFontWeight("bold");
}

// ==================== メール通知 ====================

// NYSE 休場日リスト（2026〜2027年）
// 出典: https://www.nyse.com/markets/hours-calendars
// 毎年12月頃に翌年分を追加してください
const NYSE_HOLIDAYS = {
  // 2026
  "2026-01-01": "元日 (New Year's Day)",
  "2026-01-19": "キング牧師記念日 (Martin Luther King Jr. Day)",
  "2026-02-16": "大統領の日 (Presidents' Day)",
  "2026-04-03": "聖金曜日 (Good Friday)",
  "2026-05-25": "戦没者追悼記念日 (Memorial Day)",
  "2026-06-19": "ジューンティーンス (Juneteenth)",
  "2026-07-03": "独立記念日 振替休日 (Independence Day observed)",
  "2026-09-07": "労働者の日 (Labor Day)",
  "2026-11-26": "感謝祭 (Thanksgiving Day)",
  "2026-12-25": "クリスマス (Christmas Day)",
  // 2027
  "2027-01-01": "元日 (New Year's Day)",
  "2027-01-18": "キング牧師記念日 (Martin Luther King Jr. Day)",
  "2027-02-15": "大統領の日 (Presidents' Day)",
  "2027-03-26": "聖金曜日 (Good Friday)",
  "2027-05-31": "戦没者追悼記念日 (Memorial Day)",
  "2027-06-18": "ジューンティーンス 振替 (Juneteenth observed)",
  "2027-07-05": "独立記念日 振替休日 (Independence Day observed)",
  "2027-09-06": "労働者の日 (Labor Day)",
  "2027-11-25": "感謝祭 (Thanksgiving Day)",
  "2027-12-24": "クリスマス 振替 (Christmas Day observed)",
};

// NYSE 短縮取引日（13:00 ET 早期クローズ）
const NYSE_EARLY_CLOSE = {
  "2026-07-02": "独立記念日前日 (13:00 ET 早期終了)",
  "2026-11-27": "感謝祭翌日 (13:00 ET 早期終了)",
  "2026-12-24": "クリスマスイブ (13:00 ET 早期終了)",
  "2027-11-26": "感謝祭翌日 (13:00 ET 早期終了)",
  "2027-12-23": "クリスマスイブ (13:00 ET 早期終了)",
};

// 日付を YYYY-MM-DD 形式に変換（NY基準）
function formatDateNY_(date) {
  return Utilities.formatDate(date, "America/New_York", "yyyy-MM-dd");
}

// 米国市場の休場/短縮チェック
// 日本の営業日に対応する米国市場日は「同日のNY時間」(日本の取引終了後の当日NY市場)
function checkUSMarketStatus_(jpDate) {
  // 日本の営業日 X日 → その夜に開く米国市場 = X日のNY日付
  // （日本の朝 = NYの前日夕方、日本の夕方 = NY同日の朝）
  // 厳密にはJSTの15:30頃にNYの23:30なのでまだ前日扱い、
  // しかしユーザーが見るのは「日本のX日終了後に開く米国市場」なので
  // X日のJSTを基準に「翌朝までに動く米国市場」を判定する
  // → JSTで日付を取り、それを米国日付として扱う（金曜JSTなら金曜のNY市場）

  const jstDateStr = Utilities.formatDate(jpDate, "Asia/Tokyo", "yyyy-MM-dd");
  const dateObj = new Date(jstDateStr + "T12:00:00Z");
  const dayOfWeek = dateObj.getUTCDay(); // 0=Sun, 6=Sat

  // 週末判定
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return {
      isClosed: true,
      reason: "週末 (Weekend)",
      type: "weekend"
    };
  }

  // 祝日判定
  if (NYSE_HOLIDAYS[jstDateStr]) {
    return {
      isClosed: true,
      reason: NYSE_HOLIDAYS[jstDateStr],
      type: "holiday"
    };
  }

  // 短縮取引判定
  if (NYSE_EARLY_CLOSE[jstDateStr]) {
    return {
      isClosed: false,
      isEarlyClose: true,
      reason: NYSE_EARLY_CLOSE[jstDateStr],
      type: "early_close"
    };
  }

  return { isClosed: false, type: "normal" };
}

// 翌日の日付を取得
function getNextDate_(date) {
  const next = new Date(date.getTime());
  next.setDate(next.getDate() + 1);
  return next;
}

function getEmailAddress_() {
  // 設定にアドレス指定があればそれを使用、なければ自分のGmail
  return CONFIG.EMAIL_TO || Session.getActiveUser().getEmail();
}

function sendNotificationEmail_(analysis, sessionType) {
  const to = getEmailAddress_();
  if (!to) {
    Logger.log("メール送信先が取得できませんでした");
    return;
  }

  // sessionType: "daytime" (現物・日中) or "overnight" (先物・夜間)
  const isDaytime = sessionType === "daytime";
  const sessionLabel = isDaytime ? "日中(現物)" : "夜間(EWJ連動)";
  const sessionEmoji = isDaytime ? "[DAY]" : "[NIGHT]";

  const latest = analysis[analysis.length - 1];
  const prev = analysis.length >= 2 ? analysis[analysis.length - 2] : null;
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const today = Utilities.formatDate(latest.date, "Asia/Tokyo", "yyyy/MM/dd (E)");

  // ── 米国市場の休場チェック ──
  const todayUS = checkUSMarketStatus_(latest.date);
  const tomorrowDate = getNextDate_(latest.date);
  const tomorrowUS = checkUSMarketStatus_(tomorrowDate);
  const tomorrowStr = Utilities.formatDate(tomorrowDate, "Asia/Tokyo", "yyyy/MM/dd (E)");

  const usMarketInfo = {
    today: todayUS,
    todayDate: today,
    tomorrow: tomorrowUS,
    tomorrowDate: tomorrowStr,
    hasNotice: todayUS.isClosed || todayUS.isEarlyClose || tomorrowUS.isClosed || tomorrowUS.isEarlyClose
  };

  // シグナルが買い・売り系かどうか
  const hasSignal = latest.signal && !latest.signal.includes("中立") && latest.signal !== "—";

  // 件名
  let subjectPrefix = `${sessionEmoji} [日次レポート]`;
  if (hasSignal) subjectPrefix = `${sessionEmoji} [SIGNAL: ${latest.signal}]`;
  if (usMarketInfo.hasNotice) subjectPrefix += " [US休場あり]";

  const subject = `${subjectPrefix} 日経225 ${sessionLabel} ${today}`;

  // ── HTML メール本文 ──
  const html = buildEmailHtml_(latest, prev, analysis, sheetUrl, today, usMarketInfo, sessionLabel);

  GmailApp.sendEmail(to, subject, "", { htmlBody: html });
  Logger.log("メール送信完了 (" + sessionLabel + "): " + to);
}

function sendErrorEmail_(errorMessage) {
  const to = getEmailAddress_();
  if (!to) return;

  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

  GmailApp.sendEmail(to,
    `[ERROR] 日経225 コナーズ分析 (${now})`,
    "",
    {
      htmlBody: `
        <div style="font-family:'Helvetica Neue',Arial,sans-serif;max-width:560px;margin:0 auto;padding:24px">
          <h2 style="color:#d32f2f;font-size:18px;margin:0 0 16px">[ERROR] 分析実行時にエラーが発生しました</h2>
          <div style="background:#fff3f3;border-left:4px solid #d32f2f;padding:14px 18px;border-radius:4px;margin-bottom:20px">
            <code style="font-size:13px;color:#333">${errorMessage}</code>
          </div>
          <p style="color:#666;font-size:13px;line-height:1.6">
            データソース（Google Finance / Yahoo Finance）に一時的な障害が発生している可能性があります。
            数時間後に自動で再試行されます。
          </p>
          <a href="${sheetUrl}" style="display:inline-block;background:#1a73e8;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-size:14px;margin-top:12px">
            スプレッドシートを確認
          </a>
        </div>`
    }
  );
}

function buildEmailHtml_(latest, prev, analysis, sheetUrl, today, usMarketInfo, sessionLabel) {
  // 前日比の色
  const changeColor = latest.change > 0 ? "#1b7d3a" : latest.change < 0 ? "#c62828" : "#666";
  const changeSign = latest.change > 0 ? "+" : "";

  // セッション情報（デフォルト: 日中）
  const isDay = !sessionLabel || sessionLabel.indexOf("日中") >= 0;
  const sessionBadgeBg = isDay ? "#1e3a5f" : "#3b1e5f";
  const sessionBadgeText = isDay ? "DAYTIME (CASH)" : "OVERNIGHT (EWJ-LINKED)";
  const headerGradient = isDay
    ? "linear-gradient(135deg,#1a1a2e 0%,#16213e 100%)"
    : "linear-gradient(135deg,#1a1a2e 0%,#2a1645 100%)";

  // シグナルのスタイル
  const sigStyles = getSignalStyle_(latest.signal);

  // RSI(2)のバー幅と色
  const rsi2Val = latest.rsi2 !== null ? latest.rsi2 : 50;
  const rsi2Color = rsi2Val < 10 ? "#1b7d3a" : rsi2Val > 90 ? "#c62828" : "#1a73e8";

  // 直近5日のミニテーブル
  const recentDays = analysis.slice(-5);
  const miniTableRows = recentDays.map(r => {
    const d = Utilities.formatDate(r.date, "Asia/Tokyo", "MM/dd");
    const chgC = r.change > 0 ? "#1b7d3a" : r.change < 0 ? "#c62828" : "#666";
    const chgS = r.change > 0 ? "+" : "";
    const sig = r.signal || "—";
    const sigSt = getSignalStyle_(sig);
    return `
      <tr>
        <td style="padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;color:#444">${d}</td>
        <td style="padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:right;font-family:monospace">${Number(r.close).toLocaleString("ja-JP", {minimumFractionDigits:2})}</td>
        <td style="padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:right;font-family:monospace;color:${chgC}">${chgS}${r.changePct.toFixed(2)}%</td>
        <td style="padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:right;font-family:monospace">${r.rsi2 !== null ? r.rsi2.toFixed(1) : "—"}</td>
        <td style="padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:center">
          <span style="background:${sigSt.bg};color:${sigSt.text};padding:2px 8px;border-radius:3px;font-size:11px;white-space:nowrap">${sig}</span>
        </td>
      </tr>`;
  }).join("");

  return `
  <div style="font-family:'Helvetica Neue','Hiragino Sans',Arial,sans-serif;max-width:580px;margin:0 auto;padding:0;background:#ffffff">

    <!-- ヘッダー -->
    <div style="background:${headerGradient};padding:24px 28px;border-radius:10px 10px 0 0">
      <div style="display:inline-block;background:${sessionBadgeBg};color:#ffffff;font-size:10px;letter-spacing:0.15em;padding:4px 10px;border-radius:3px;margin-bottom:10px;font-weight:600">${sessionBadgeText}</div>
      <div style="font-size:12px;color:#8888bb;letter-spacing:0.1em;margin-bottom:4px">CONNORS STRATEGY</div>
      <div style="font-size:22px;color:#ffffff;font-weight:700">日経225 ${sessionLabel || "コナーズ分析"}</div>
      <div style="font-size:13px;color:#aaaacc;margin-top:4px">${today}</div>
    </div>

    <!-- メインカード -->
    <div style="border:1px solid #e8e8e8;border-top:none;padding:24px 28px">

      <!-- 終値 & 前日比 -->
      <div style="display:flex;margin-bottom:20px">
        <div style="flex:1">
          <div style="font-size:11px;color:#999;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px">終値</div>
          <div style="font-size:28px;font-weight:700;color:#1a1a1a;font-family:monospace">${Number(latest.close).toLocaleString("ja-JP", {minimumFractionDigits:2})}</div>
        </div>
        <div style="flex:1;text-align:right">
          <div style="font-size:11px;color:#999;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px">前日比</div>
          <div style="font-size:22px;font-weight:700;color:${changeColor};font-family:monospace">${changeSign}${Number(latest.change).toLocaleString("ja-JP", {minimumFractionDigits:2})}</div>
          <div style="font-size:14px;color:${changeColor};font-family:monospace">(${changeSign}${latest.changePct.toFixed(2)}%)</div>
        </div>
      </div>

      <!-- シグナル -->
      <div style="background:${sigStyles.bgLight};border:1px solid ${sigStyles.border};border-radius:8px;padding:16px 20px;margin-bottom:20px;text-align:center">
        <div style="font-size:11px;color:#888;letter-spacing:0.08em;margin-bottom:6px">TODAY'S SIGNAL</div>
        <div style="font-size:24px;font-weight:700;color:${sigStyles.textDark}">${latest.signal}</div>
      </div>

      ${buildUSMarketSection_(usMarketInfo)}

      <!-- インジケーター -->
      <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0">
            <span style="font-size:12px;color:#888">RSI(2)</span>
          </td>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0;text-align:right">
            <span style="font-size:15px;font-weight:600;color:${rsi2Color};font-family:monospace">${latest.rsi2 !== null ? latest.rsi2.toFixed(1) : "—"}</span>
          </td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0">
            <span style="font-size:12px;color:#888">RSI(4)</span>
          </td>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0;text-align:right">
            <span style="font-size:15px;font-weight:600;color:#333;font-family:monospace">${latest.rsi4 !== null ? latest.rsi4.toFixed(1) : "—"}</span>
          </td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0">
            <span style="font-size:12px;color:#888">連騰日数</span>
          </td>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0;text-align:right">
            <span style="font-size:15px;font-weight:600;color:${latest.streak > 0 ? "#1b7d3a" : latest.streak < 0 ? "#c62828" : "#666"};font-family:monospace">${latest.streak > 0 ? "+" + latest.streak : latest.streak}日</span>
          </td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0">
            <span style="font-size:12px;color:#888">5日MA</span>
          </td>
          <td style="padding:8px 0;border-bottom:1px solid #f0f0f0;text-align:right">
            <span style="font-size:14px;color:#444;font-family:monospace">${latest.sma5 ? Number(latest.sma5).toLocaleString("ja-JP", {minimumFractionDigits:2}) : "—"}</span>
          </td>
        </tr>
        <tr>
          <td style="padding:8px 0">
            <span style="font-size:12px;color:#888">200日MA</span>
          </td>
          <td style="padding:8px 0;text-align:right">
            <span style="font-size:14px;color:#444;font-family:monospace">${latest.sma200 ? Number(latest.sma200).toLocaleString("ja-JP", {minimumFractionDigits:2}) : "—"}</span>
          </td>
        </tr>
      </table>

      <!-- 直近5日テーブル -->
      <div style="font-size:12px;color:#888;margin-bottom:8px;letter-spacing:0.05em">直近5営業日</div>
      <table style="width:100%;border-collapse:collapse;margin-bottom:24px">
        <tr style="background:#f8f9fa">
          <th style="padding:7px 10px;text-align:left;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0">日付</th>
          <th style="padding:7px 10px;text-align:right;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0">終値</th>
          <th style="padding:7px 10px;text-align:right;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0">前日比%</th>
          <th style="padding:7px 10px;text-align:right;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0">RSI(2)</th>
          <th style="padding:7px 10px;text-align:center;font-size:11px;color:#888;font-weight:600;border-bottom:2px solid #e0e0e0">シグナル</th>
        </tr>
        ${miniTableRows}
      </table>

      <!-- スプレッドシートリンク -->
      <div style="text-align:center">
        <a href="${sheetUrl}" style="display:inline-block;background:#1a73e8;color:#ffffff;padding:14px 36px;border-radius:8px;text-decoration:none;font-size:15px;font-weight:600;letter-spacing:0.02em">
          &gt;&gt; スプレッドシートで詳細を見る
        </a>
      </div>
    </div>

    <!-- フッター -->
    <div style="padding:16px 28px;font-size:11px;color:#aaa;text-align:center;border:1px solid #e8e8e8;border-top:none;border-radius:0 0 10px 10px;background:#fafafa">
      日経225 × ラリー・コナーズ戦略 ｜ 自動分析レポート<br>
      ※投資助言ではありません。投資判断は自己責任でお願いします。
    </div>
  </div>`;
}

function buildUSMarketSection_(info) {
  if (!info || !info.hasNotice) {
    // 休場情報がない場合はコンパクトに「平常」と表示
    return `
      <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin-bottom:20px">
        <div style="font-size:11px;color:#64748b;letter-spacing:0.08em;margin-bottom:4px">US MARKET STATUS</div>
        <div style="font-size:13px;color:#475569">本日・翌日ともに通常取引</div>
      </div>`;
  }

  const renderDayStatus = (dayInfo, label, dateStr) => {
    if (dayInfo.isClosed) {
      return `
        <div style="display:flex;align-items:flex-start;padding:10px 0;border-bottom:1px solid #fee2e2">
          <div style="min-width:90px">
            <div style="font-size:11px;color:#991b1b;font-weight:600">${label}</div>
            <div style="font-size:11px;color:#991b1b;opacity:0.7">${dateStr}</div>
          </div>
          <div style="flex:1">
            <div style="display:inline-block;background:#dc2626;color:#ffffff;padding:2px 10px;border-radius:3px;font-size:11px;font-weight:600;margin-bottom:4px">CLOSED</div>
            <div style="font-size:13px;color:#7f1d1d;margin-top:2px">${dayInfo.reason}</div>
          </div>
        </div>`;
    }
    if (dayInfo.isEarlyClose) {
      return `
        <div style="display:flex;align-items:flex-start;padding:10px 0;border-bottom:1px solid #fef3c7">
          <div style="min-width:90px">
            <div style="font-size:11px;color:#78350f;font-weight:600">${label}</div>
            <div style="font-size:11px;color:#78350f;opacity:0.7">${dateStr}</div>
          </div>
          <div style="flex:1">
            <div style="display:inline-block;background:#d97706;color:#ffffff;padding:2px 10px;border-radius:3px;font-size:11px;font-weight:600;margin-bottom:4px">EARLY CLOSE</div>
            <div style="font-size:13px;color:#78350f;margin-top:2px">${dayInfo.reason}</div>
          </div>
        </div>`;
    }
    return `
      <div style="display:flex;align-items:flex-start;padding:10px 0;border-bottom:1px solid #e2e8f0">
        <div style="min-width:90px">
          <div style="font-size:11px;color:#475569;font-weight:600">${label}</div>
          <div style="font-size:11px;color:#475569;opacity:0.7">${dateStr}</div>
        </div>
        <div style="flex:1">
          <div style="display:inline-block;background:#16a34a;color:#ffffff;padding:2px 10px;border-radius:3px;font-size:11px;font-weight:600">OPEN</div>
          <div style="font-size:13px;color:#334155;margin-top:2px">通常取引</div>
        </div>
      </div>`;
  };

  // 米国市場が今晩〜明日の朝に開かない場合、日経の翌営業日が手がかり難になる旨を補足
  const closedToday = info.today.isClosed;
  const closedTomorrow = info.tomorrow.isClosed;
  let supplementMsg = "";
  if (closedToday && closedTomorrow) {
    supplementMsg = "本日・翌日ともに米国市場が休場です。海外材料の手がかりが乏しくなる可能性があります。";
  } else if (closedToday) {
    supplementMsg = "本日の米国市場が休場のため、翌営業日の日本市場は手がかり難となる可能性があります。";
  } else if (closedTomorrow) {
    supplementMsg = "翌営業日の米国市場が休場予定です。先回り取引や様子見ムードに注意。";
  }

  return `
    <div style="background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;padding:14px 18px;margin-bottom:20px">
      <div style="display:flex;align-items:center;margin-bottom:8px">
        <div style="font-size:11px;color:#92400e;letter-spacing:0.08em;font-weight:700">US MARKET NOTICE</div>
      </div>
      <div style="margin:0">
        ${renderDayStatus(info.today, "本日", info.todayDate)}
        ${renderDayStatus(info.tomorrow, "翌日", info.tomorrowDate)}
      </div>
      ${supplementMsg ? `<div style="font-size:12px;color:#78350f;line-height:1.6;margin-top:10px;padding-top:10px;border-top:1px dashed #fcd34d">${supplementMsg}</div>` : ""}
    </div>`;
}

function getSignalStyle_(signal) {
  // bg: ミニテーブルのバッジ背景（濃い色）
  // text: シグナルカード大文字 & バッジ文字色（白）
  // bgLight: シグナルカードの背景（淡い色）
  // textDark: シグナルカード上の文字色（濃い色、bgLight上で映える）
  // border: シグナルカードの枠線
  if (!signal) return { bg: "#999", text: "#ffffff", bgLight: "#f5f5f5", textDark: "#666666", border: "#e0e0e0" };

  if (signal.includes("強い買い")) return { bg: "#15803d", text: "#ffffff", bgLight: "#dcfce7", textDark: "#14532d", border: "#86efac" };
  if (signal.includes("買い"))     return { bg: "#16a34a", text: "#ffffff", bgLight: "#dcfce7", textDark: "#166534", border: "#86efac" };
  if (signal.includes("強い売り")) return { bg: "#b91c1c", text: "#ffffff", bgLight: "#fee2e2", textDark: "#7f1d1d", border: "#fca5a5" };
  if (signal.includes("売り"))     return { bg: "#dc2626", text: "#ffffff", bgLight: "#fee2e2", textDark: "#991b1b", border: "#fca5a5" };
  if (signal.includes("注意"))     return { bg: "#d97706", text: "#ffffff", bgLight: "#fef3c7", textDark: "#78350f", border: "#fcd34d" };

  return { bg: "#64748b", text: "#ffffff", bgLight: "#f1f5f9", textDark: "#334155", border: "#cbd5e1" };
}

// ==================== ユーティリティ ====================

// 手動で全データを再計算
function forceRecalculate() {
  dailyUpdate();
  SpreadsheetApp.getUi().alert("再計算完了！");
}

// メール通知テスト送信（日中版）
function testEmailDaytime() {
  _testEmailSession("daytime");
}

// メール通知テスト送信（夜間版）
function testEmailOvernight() {
  _testEmailSession("overnight");
}

function _testEmailSession(sessionType) {
  const to = getEmailAddress_();
  if (!to) {
    SpreadsheetApp.getUi().alert("メールアドレスが取得できません。\nCONFIG.EMAIL_TO にアドレスを設定してください。");
    return;
  }

  try {
    const prices = sessionType === "overnight"
      ? fetchFuturesData_()
      : fetchNikkeiData_();

    if (!prices || prices.length < CONFIG.SMA_PERIOD + 1) {
      SpreadsheetApp.getUi().alert("データが不足しています。先に「今すぐ更新」を実行してください。");
      return;
    }
    const analysis = analyzeConnors_(prices);
    sendNotificationEmail_(analysis, sessionType);
    SpreadsheetApp.getUi().alert("テストメール送信完了！\n送信先: " + to + "\nセッション: " + sessionType);
  } catch (e) {
    SpreadsheetApp.getUi().alert("メール送信エラー:\n" + e.message);
  }
}

// トリガーを全削除
function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  SpreadsheetApp.getUi().alert("全トリガーを削除しました。");
}

// カスタムメニュー追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("コナーズ分析")
    .addItem("日中+夜間 まとめて更新", "dailyUpdate")
    .addSeparator()
    .addItem("日中のみ更新 (現物)", "dailyUpdateDaytime")
    .addItem("夜間のみ更新 (EWJ)", "dailyUpdateOvernight")
    .addSeparator()
    .addItem("日次自動実行を設定", "setupDailyTrigger")
    .addItem("初期セットアップ", "setupSheet")
    .addSeparator()
    .addItem("テストメール(日中)", "testEmailDaytime")
    .addItem("テストメール(夜間)", "testEmailOvernight")
    .addItem("全データ再計算", "forceRecalculate")
    .addItem("トリガー全削除", "removeAllTriggers")
    .addToUi();
}
