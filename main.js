function getDVSMessage(dvsRank) {
  if (dvsRank === "A") {
    return "食品摂取状況はとても良いです。毎日7点以上をこれから続けましょう。";
  } else if (dvsRank === "B") {
    return "食品摂取状況が少し心配です。食べやすい食品をプラスして栄養バランスを向上させましょう。";
  } else if (dvsRank === "C") {
    return "食品摂取状況が心配です。毎日の食生活に3ポイントプラスするように心がけましょう。";
  }
  return "DVS Not Found.";
}

function doGet() {
  // スプレッドシートとシートの取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow(); // 最新の行
  const headers = sheet.getDataRange().getValues()[0]; // ヘッダー
  const row = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0]; // 最新の行データ
  
  // MNA スコア (列 2~7 の範囲を仮定)
  const mnaScores = row.slice(1, 7);
  const dvsScores = row.slice(7, 17);

  // スコア合計の計算
  const mnaTotal = mnaScores.reduce((sum, score) => sum + (parseInt(score, 10) || 0), 0);
  const dvsTotal = dvsScores.reduce((sum, score) => sum + (parseInt(score, 10) || 0), 0);

  // MNA ランク判定
  let mnaRank = "C";
  if (mnaTotal >= 12) mnaRank = "A";
  else if (mnaTotal >= 8) mnaRank = "B";

  // DVS ランク判定
  let dvsRank = "C";
  if (dvsTotal >= 7) dvsRank = "A";
  else if (dvsTotal >= 4) dvsRank = "B";

  // メッセージ作成
  let mnaMessage = "";
  if (mnaRank === "A") {
    mnaMessage = "栄養状態は良好です。";
  } else if (mnaRank === "B") {
    mnaMessage = "低栄養の恐れがあります。";
  } else if (mnaRank === "C") {
    mnaMessage = "低栄養状態です。";
  }
  const dvsMessage = getDVSMessage(dvsRank);

  // HTML 出力 (TailwindCSS 適用)
  return HtmlService.createHtmlOutput(`
    <html>
      <head>
        <title>結果</title>
        <script src="https://cdn.tailwindcss.com"></script>
      </head>
      <body class="bg-gray-100 flex flex-col items-center justify-center min-h-screen">
        <div class="bg-white shadow-md rounded-lg p-8 max-w-md text-center">
          <h1 class="text-2xl font-bold text-green-500 mb-4">結果</h1>
          <ul class="list-disc text-left text-gray-700 space-y-2">
            <li><span class="font-semibold">MNA:<span class="text-red-600">${mnaTotal}</span></span> ${mnaMessage}</li>
            <li><span class="font-semibold">DVS:<span class="text-red-600">${dvsTotal}</span></span> ${dvsMessage}</li>
          </ul>
        </div>
      </body>
    </html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
