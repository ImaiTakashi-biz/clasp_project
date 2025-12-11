// ======================
// 設定
// ======================
const SHEET_NAME = "セット品検査記録";
const CACHE_TTL_HOURS = 24; // 送信履歴キャッシュ保持期間（時間）

// ======================
// メイン処理
// ======================
function main() {
  try {
    // スクリプトプロパティから設定値を取得
    const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    
    if (!SPREADSHEET_ID) {
      Logger.log("SPREADSHEET_IDがスクリプトプロパティに設定されていません");
      return;
    }
    
    // 1. Google Sheetsからデータを取得
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log("シートが見つかりません: " + SHEET_NAME);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // ヘッダーのインデックスを取得
    const headerIndexes = {};
    headers.forEach((header, index) => {
      headerIndexes[header] = index;
    });
    
    // データ行を処理
    const rowsToProcess = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowData = {};
      headers.forEach((header, index) => {
        rowData[header] = row[index];
      });
      rowData.row_number = i + 1; // 行番号（1ベース）
      
      // 条件チェック: 検査数が空でない、かつ送信が空
      const inspectionCount = rowData["検査数"];
      const sentFlag = rowData["送信"];
      
      if (inspectionCount && inspectionCount !== "" && (!sentFlag || sentFlag === "")) {
        rowsToProcess.push(rowData);
      }
    }
    
    if (rowsToProcess.length === 0) {
      Logger.log("処理対象のデータがありません");
      return;
    }
    
    // 2. 不良率をパーセントに変換
    const processedRows = rowsToProcess.map(row => {
      let rate = row["不良率"];
      if (rate === "" || rate === null || rate === undefined) {
        row["不良率"] = "0.0%";
      } else {
        row["不良率"] = (Number(rate) * 100).toFixed(1) + "%";
      }
      return row;
    });
    
    // 3. AIエージェント用のプロンプト生成とHTML生成
    const promptData = generatePromptForAI(processedRows);
    
    // 4. AI分析コメントを生成
    const aiCommentHtml = generateAIComment(promptData.dataForAgent);
    
    // 5. HTMLとAIコメントを結合してファイル生成
    const finalHtml = combineHtmlAndAIComment(promptData.initialHtml, aiCommentHtml, processedRows, promptData.fileName);
    
    // 6. 生成HTMLをARAICHATへ直接送信
    const htmlBytes = Utilities.newBlob(finalHtml, 'text/html', promptData.fileName).getBytes();
    const sendResult = sendFileToAraichat(htmlBytes, promptData.fileName);
    if (!sendResult) {
      Logger.log("ARAICHATへの送信に失敗しました");
    }
    
    // 7. 送信フラグを「済」に更新
    processedRows.forEach(row => {
      sheet.getRange(row.row_number, headerIndexes["送信"] + 1).setValue("済");
    });
    
    Logger.log(`処理完了: ${processedRows.length}件のデータを処理しました`);
    
  } catch (error) {
    Logger.log("エラーが発生しました: " + error.toString());
    throw error;
  }
}

// ======================
// Google DriveのHTMLファイルを削除
// ======================
function deleteHtmlFilesInDrive(DRIVE_FOLDER_ID) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    if (!folder) {
      Logger.log("フォルダが見つかりません: " + DRIVE_FOLDER_ID);
      return;
    }
    
    // HTMLファイルを検索（配列に一度格納してから処理）
    const files = folder.getFilesByType(MimeType.HTML);
    const fileList = [];
    while (files.hasNext()) {
      fileList.push(files.next());
    }
    
    if (fileList.length === 0) {
      Logger.log("削除対象のHTMLファイルはありませんでした");
      return;
    }
    
    let deletedCount = 0;
    let errorCount = 0;
    
    // 各ファイルを削除
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      const fileName = file.getName();
      const fileId = file.getId();
      
      try {
        // ファイルをゴミ箱に移動
        file.setTrashed(true);
        
        // 削除を確認（少し待ってから確認）
        Utilities.sleep(500); // 500ms待機
        
        // ファイルがまだ存在するか確認
        try {
          const checkFile = DriveApp.getFileById(fileId);
          if (checkFile.isTrashed()) {
            deletedCount++;
            Logger.log(`削除成功（ゴミ箱に移動）: ${fileName}`);
          } else {
            // ゴミ箱に移動できなかった場合は完全削除を試みる
            try {
              DriveApp.removeFile(checkFile);
              deletedCount++;
              Logger.log(`削除成功（完全削除）: ${fileName}`);
            } catch (removeError) {
              errorCount++;
              Logger.log(`削除失敗: ${fileName} - ${removeError.toString()}`);
            }
          }
        } catch (notFoundError) {
          // ファイルが見つからない = 削除成功
          deletedCount++;
          Logger.log(`削除成功: ${fileName}`);
        }
      } catch (trashError) {
        errorCount++;
        Logger.log(`削除失敗: ${fileName} - ${trashError.toString()}`);
      }
    }
    
    if (deletedCount > 0) {
      Logger.log(`${deletedCount}件のHTMLファイルを削除しました`);
    }
    if (errorCount > 0) {
      Logger.log(`${errorCount}件のファイルで削除エラーが発生しました`);
    }
  } catch (error) {
    Logger.log("HTMLファイル削除エラー: " + error.toString());
    Logger.log("エラー詳細: " + JSON.stringify(error));
  }
}

// ======================
// AIエージェント用のプロンプト生成
// ======================
function generatePromptForAI(dataRows) {
  // 検査日を取得（最初の行から）
  const inspectionDateRaw = dataRows.length > 0 && dataRows[0]["検査日"] 
    ? dataRows[0]["検査日"] 
    : null;
  
  const inspectionDate = formatDate(inspectionDateRaw) || "日付不明";
  
  // ファイル名用の日付文字列を生成（Dateオブジェクトにも対応）
  let fileDate = "日付不明";
  if (inspectionDateRaw) {
    if (inspectionDateRaw instanceof Date) {
      // Dateオブジェクトの場合
      const year = inspectionDateRaw.getFullYear();
      const month = String(inspectionDateRaw.getMonth() + 1).padStart(2, '0');
      const day = String(inspectionDateRaw.getDate()).padStart(2, '0');
      fileDate = `${year}-${month}-${day}`;
    } else if (typeof inspectionDateRaw === 'string') {
      // 文字列の場合
      fileDate = inspectionDateRaw.substring(0, 10).replace(/\//g, '-');
    } else {
      // その他の場合（数値のタイムスタンプなど）
      const d = new Date(inspectionDateRaw);
      if (!isNaN(d.getTime())) {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        fileDate = `${year}-${month}-${day}`;
      }
    }
  }
  
  const fileName = `セット品検査結果報告_${fileDate}.html`;
  
  // HTML初期化
  let initialHtml = `<h1>${inspectionDate}付のセット品検査結果報告</h1>`;
  
  // テーブルヘッダー
  initialHtml += `
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: メイリオ, Meiryo, ＭＳ Ｐゴシック, MS PGothic, Arial, sans-serif; font-size: 14px; width: 100%;">
  <thead style="background-color: #f2f2f2;">
    <tr>
      <th>セット日</th>
      <th>セット者</th>
      <th>機番</th>
      <th>客先</th>
      <th>品番</th>
      <th>品名</th>
      <th>検査数</th>
      <th>不良数</th>
      <th>不良率</th>
      <th>不具合内容</th>
    </tr>
  </thead>
  <tbody>
`;
  
  // 各行を追加
  for (const row of dataRows) {
    const rate = row["不良率"] !== undefined && row["不良率"] !== null ? row["不良率"] : "";
    const comment = row["コメント"] && row["コメント"].trim() !== "" ? row["コメント"] : "";
    
    initialHtml += `
    <tr>
      <td>${formatDateSlash(row["セット日"])}</td>
      <td>${row["セット者"] || ""}</td>
      <td>${row["機番"] || ""}</td>
      <td>${row["客先"] || ""}</td>
      <td>${row["品番"] || ""}</td>
      <td>${row["品名"] || ""}</td>
      <td align="right">${Number(row["検査数"] || 0).toLocaleString()}</td>
      <td align="right">${Number(row["不良数"] || 0).toLocaleString()}</td>
      <td align="right">${rate}</td>
      <td>${comment}</td>
    </tr>
  `;
  }
  
  // テーブル閉じタグ
  initialHtml += `
  </tbody>
</table>`;
  
  // AIエージェント用データ
  const dataForAgent = JSON.stringify(dataRows);
  
  return {
    initialHtml: initialHtml,
    dataForAgent: dataForAgent,
    fileName: fileName
  };
}

// ======================
// 日付整形関数
// ======================
function formatDate(dateString) {
  if (!dateString) return "";
  const d = new Date(dateString);
  if (isNaN(d.getTime())) return "";
  const mm = (d.getMonth() + 1).toString().padStart(2, '0');
  const dd = d.getDate().toString().padStart(2, '0');
  return `${d.getFullYear()}年${mm}月${dd}日`;
}

// ======================
// 日付をYYYY/MM/DD形式に変換する関数
// ======================
function formatDateSlash(dateValue) {
  if (!dateValue) return "";
  let d;
  if (dateValue instanceof Date) {
    d = dateValue;
  } else {
    d = new Date(dateValue);
  }
  if (isNaN(d.getTime())) return "";
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

// ======================
// AI分析コメント生成（Gemini API）
// ======================
function generateAIComment(dataForAgent) {
  // スクリプトプロパティからGEMINI_API_KEYを取得
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!GEMINI_API_KEY) {
    Logger.log("Gemini APIキーがスクリプトプロパティに設定されていません");
    return "";
  }
  
  const prompt = `あなたは、**製造業の精密部品加工の品質管理担当者であり、熟練された上級位のプロのCNC自動旋盤オペレーター**です。
（シチズンマシナリー株式会社のCincom製マシンを使用）

以下の入力データ（JSON形式の検査結果データ）を元に、**検査結果のAI分析コメントのみ**を生成してください。

---

<h3>目的</h3>
入力データを分析し、**不良の根本原因特定と、その低減のための最も効果的かつ実践的な改善策を、簡潔かつ要点を絞って提供する**ため、熟練された上級位のプロのCNC自動旋盤オペレーター視点と統計品質管理の知見に基づいた、分かりやすい分析コメントを生成する。

---

<h3>出力ルール</h3>
<ol>
  <li>書式：<code>&lt;ul&gt;&lt;li&gt;〜&lt;/li&gt;&lt;/ul&gt;</code>の箇条書き形式で出力すること。</li>
  <li>同一の客先・品番・品名については、原因と改善策を1つの&lt;li&gt;内にまとめ、別の&lt;li&gt;やテーブルに分けないこと。</li>
  <li>「OK切粉」「切粉OK」「ok切粉」「切粉ok」など、切粉を検査員が除去してOK処理したコメントは不良計上対象外として扱い、切粉除去済みの良品として解釈すること。</li>
  <li>視点：**製造業の精密部品加工の品質管理担当者としての深い分析力と、CNC自動旋盤の熟練された上級位のプロのオペレーターとしての実践的な経験**、そして統計品質管理（SQC）に基づく多角的な品質分析の視点から、**要点を絞って**コメントすること。</li>
  <li>表現：誰もが理解しやすい、**簡潔で分かりやすい日本語**で記述し、回りくどい表現は避ける。専門用語には必要に応じ簡単な補足を付ける。</li>
  <li>敬称：客先名に「様」などは付けないこと。</li>
  <li>用語：「チャック爪」という表現は使わず、必ず「チャック」と表記すること。</li>
  <li>用語補足：「処理前」はメッキなどの表面処理を行う前の状態を指すものとして解釈すること。</li>
  <li>**1コメントあたり200文字以内に収めること。**</li>
  <li>内容例：
    <ul>
      <li>各不良項目（客先、品番、品名単位）に対し、最も可能性の高い原因と具体的な改善策を簡潔にまとめ、その効果や優先度も短く示唆する。</li>
      <li>原因として考慮すべき観点：切削条件、工具、クランプ方法、加工後の排出・搬送、段取り初日の稼働における初期流動の観点も含む。</li>
      <li>切削油に関しては、種類変更不可・高圧クーラント不可・濃度およびpH管理も現状不可。</li>
      <li>不具合内容に「落下」という文言が含まれる場合、原因は「検査員が誤って落下させたこと」とし、その旨を明記してコメントする。</li>
      <li>安定している項目は成功要因を簡潔に分析し、横展開の可能性を示唆する。</li>
      <li>予防保全やモニタリングの提案は重要時のみ簡潔に含める。</li>
    </ul>
  </li>
  <li>**出力は分析コメントのHTML（見出しと箇条書き）のみとし、他は含めない。**</li>
</ol>

---

<h3>入力データ</h3>
<code>${dataForAgent}</code>

---

<h3>出力形式</h3>
<ul>
  <li>HTML形式でそのまま使える分析コメントの本文のみを出力すること。</li>
  <li>余分な返答は不要。</li>
  <li>正しいHTML構文を保持すること。</li>
</ul>`;
  
  try {
    // Gemini APIの最新モデル名を使用（gemini-2.5-flash または gemini-2.5-pro）
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
    
    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }]
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log("Gemini API エラー: " + response.getContentText());
      return "";
    }
    
    const result = JSON.parse(response.getContentText());
    const aiComment = result.candidates && result.candidates[0] && result.candidates[0].content
      ? result.candidates[0].content.parts[0].text
      : "";
    
    return aiComment || "";
    
  } catch (error) {
    Logger.log("AI分析コメント生成エラー: " + error.toString());
    return "";
  }
}

// ======================
// HTMLとAIコメントを結合
// ======================
function combineHtmlAndAIComment(initialHtml, aiCommentHtml, dataRows, fileName) {
  const generatedTableHtml = generateTableHtml(dataRows);

  // レポート日付（ファイル名から抽出）
  const badgeDate = fileName
    .replace('セット品検査結果報告_', '')
    .replace('.html', '')
    .trim();

  // ロゴ（Base64埋め込み）
  const embeddedLogoSrc = 'data:image/png;base64,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';

  const analysisHtml = `
    <div class="analysis-list">
      ${aiCommentHtml || ""}
    </div>
    <div class="footer-note">※ コメントはAIによる自動生成であり、最終判断は担当者が行ってください。</div>
  `;

  const finalHtml = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>${fileName.replace('.html', '')}</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    * { box-sizing: border-box; }
    :root {
      --primary: #2563eb;
      --primary-soft: #e0edff;
      --bg-page: #f3f4f6;
      --text-main: #111827;
      --text-sub: #6b7280;
      --border-soft: #e5e7eb;
      --card-radius: 16px;
    }
    body {
      margin: 0;
      min-height: 100vh;
      background: radial-gradient(circle at top left, #e0edff 0, #f3f4f6 40%, #eef2ff 100%);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI","Noto Sans JP", sans-serif;
      color: var(--text-main);
    }
    .layout { max-width: 1100px; margin: 24px auto 40px; padding: 0 20px; }
    .app-header {
      background: linear-gradient(135deg, #1d4ed8, #2563eb, #38bdf8);
      border-radius: 24px;
      padding: 18px 24px;
      color: #f9fafb;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      box-shadow: 0 18px 40px rgba(15, 23, 42, 0.25);
      margin-bottom: 22px;
    }
    .app-header-left { display: flex; align-items: center; gap: 16px; min-width: 0; }
    .app-logo-wrap {
      width: 56px; height: 56px; border-radius: 18px;
      background: rgba(15, 23, 42, 0.35);
      display: flex; align-items: center; justify-content: center; flex-shrink: 0;
    }
    .app-logo-wrap img { max-width: 80%; max-height: 80%; display: block; filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.4)); }
    .app-title-block { display: flex; flex-direction: column; gap: 6px; min-width: 0; }
    .app-title-main { font-size: 20px; font-weight: 600; letter-spacing: 0.02em; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .app-title-sub { font-size: 12px; opacity: 0.9; }
    .app-header-right { display: flex; align-items: center; gap: 12px; flex-shrink: 0; font-size: 12px; }
    .badge {
      padding: 4px 10px; border-radius: 999px;
      background: rgba(15, 23, 42, 0.28);
      border: 1px solid rgba(248, 250, 252, 0.22);
      display: inline-flex; align-items: center; gap: 6px;
    }
    .badge-dot { width: 8px; height: 8px; border-radius: 999px; background: #22c55e; box-shadow: 0 0 0 4px rgba(34, 197, 94, 0.3); }
    .badge-label { opacity: 0.95; }
    .content-column { display: flex; flex-direction: column; gap: 16px; margin-top: 6px; }
    .card {
      background: #ffffffcc; backdrop-filter: blur(8px);
      border-radius: var(--card-radius);
      border: 1px solid rgba(148, 163, 184, 0.25);
      padding: 16px 18px 18px;
      box-shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
    }
    .card-header { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 10px; }
    .card-title { font-size: 14px; font-weight: 600; display: flex; align-items: center; gap: 8px; color: #111827; }
    .card-title-pill { width: 18px; height: 18px; border-radius: 999px; background: var(--primary-soft); display: inline-flex; align-items: center; justify-content: center; font-size: 11px; color: var(--primary); }
    .card-caption { font-size: 11px; color: var(--text-sub); }
    /* テーブル（PC表示優先）、カード（モバイル表示） */
    .table-block { display: none; }
    .cards-block { display: block; }
    .table-scroll { width: 100%; overflow-x: auto; }
    .table-scroll::-webkit-scrollbar { height: 6px; }
    .table-scroll::-webkit-scrollbar-thumb { background: #cbd5f5; border-radius: 999px; }
    .report-table { width: 100%; border-collapse: collapse; font-size: 12px; min-width: 720px; }
    .report-table thead { background: #f1f5f9; }
    .report-table th, .report-table td { padding: 6px 8px; border-bottom: 1px solid var(--border-soft); text-align: center; white-space: nowrap; }
    .report-table th { font-size: 11px; font-weight: 600; color: #4b5563; }
    .report-table tbody tr:hover td { background: #f9fafb; }
    .number-cell { text-align: right; font-variant-numeric: tabular-nums; }
    .highlight-ng { color: #dc2626; font-weight: 700; }

    /* レコードカード（モバイル優先） */
    .records-wrap { display: flex; flex-direction: column; gap: 12px; }
    .record-card {
      background: #fff;
      border: 1px solid #e5e7eb;
      border-radius: 12px;
      box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
      overflow: hidden;
    }
    .record-row {
      display: flex;
      justify-content: space-between;
      padding: 10px 12px;
      border-bottom: 1px solid #e5e7eb;
      font-size: 13px;
      line-height: 1.5;
    }
    .record-row:last-child { border-bottom: none; }
    .record-label { font-weight: 700; color: #111827; }
    .record-value { color: #111827; text-align: right; }
    .metric-ok { color: #16a34a; font-weight: 700; }
    .metric-ng { color: #dc2626; font-weight: 700; }
    /* デスクトップではテーブル表示（カード非表示）、カードは2カラム */
    @media (min-width: 900px) {
      .table-block { display: block; }
      .cards-block { display: none; }
      .records-wrap { display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); gap: 16px; }
      .record-row { font-size: 13px; }
    }
    .analysis-list { display: flex; flex-direction: column; gap: 10px; margin-top: 4px; font-size: 12px; list-style: none; padding: 0; }
    .analysis-list ul { list-style: none; padding: 0; margin: 0; display: flex; flex-direction: column; gap: 10px; }
    .analysis-list li {
      border-radius: 12px;
      border: 1px solid var(--primary-soft);
      background: #f8fafc;
      padding: 10px 11px 10px 18px;
      position: relative;
      line-height: 1.7;
      color: #4b5563;
    }
    .analysis-list li::before {
      content: "";
      position: absolute;
      left: 10px;
      top: 12px;
      width: 3px;
      height: 22px;
      border-radius: 999px;
      background: var(--primary);
    }
    .analysis-list ul ul { margin-top: 6px; }
    .analysis-list ul ul li { background: #eef2ff; border-color: #d8e3ff; }
    .analysis-header { display: flex; align-items: center; gap: 6px; margin-left: 4px; margin-bottom: 4px; color: #0f172a; font-weight: 600; font-size: 12px; }
    .analysis-chip { padding: 1px 6px; border-radius: 999px; background: var(--primary-soft); color: var(--primary); font-size: 10px; }
    .analysis-body { margin-left: 4px; color: #4b5563; line-height: 1.7; }
    .footer-note { margin-top: 14px; font-size: 10px; color: var(--text-sub); text-align: right; }
    @media (max-width: 960px) {
      .layout { margin: 16px auto 28px; padding: 0 14px; }
      .app-header { padding: 14px 16px; border-radius: 18px; }
      .card { padding: 14px 14px 16px; border-radius: 14px; }
    }
    @media (max-width: 640px) {
      .layout { margin: 12px auto 24px; padding: 0 10px; }
      .app-header { flex-direction: column; align-items: flex-start; gap: 10px; }
      .app-header-left { width: 100%; justify-content: flex-start; }
      .app-logo-wrap { width: 48px; height: 48px; }
      .app-title-main { font-size: 17px; white-space: normal; }
      .app-title-sub { font-size: 11px; }
      .app-header-right { width: 100%; justify-content: flex-start; font-size: 11px; }
      .card { padding: 12px 12px 14px; box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08); }
      .card-header { flex-direction: column; align-items: flex-start; gap: 4px; }
      .card-title { font-size: 13px; }
      .card-caption { font-size: 10px; }
      .report-table { font-size: 11px; }
      .report-table th { font-size: 10px; }
      .analysis-list { font-size: 11px; }
      .analysis-header { font-size: 11px; }
    }
    @media print {
      body { background: #ffffff; }
      .layout { max-width: none; margin: 0; padding: 0; }
      .app-header { box-shadow: none; }
      .card { box-shadow: none; background: #fff; }
    }
  </style>
</head>
<body>
  <div class="layout">
    <header class="app-header">
      <div class="app-header-left">
        <div class="app-logo-wrap">
          <img src="${embeddedLogoSrc}" alt="ARAI ロゴ">
        </div>
        <div class="app-title-block">
          <div class="app-title-main">セット品検査結果ダッシュボード</div>
          <div class="app-title-sub">株式会社新井精密 | ARAI Quality Insight</div>
        </div>
      </div>
      <div class="app-header-right">
        <div class="badge">
          <span class="badge-dot"></span>
          <span class="badge-label">${badgeDate || "レポート"}</span>
        </div>
      </div>
    </header>

    <main class="content-column">
      <section class="card">
        <div class="card-header">
          <div class="card-title">
            <span class="card-title-pill">T</span>
            <span>セット品検査結果</span>
          </div>
          <div class="card-caption">検査結果サマリー</div>
        </div>
        ${generatedTableHtml}
      </section>

      <section class="card">
        <div class="card-header">
          <div class="card-title">
            <span class="card-title-pill">AI</span>
            <span>AI分析コメント</span>
          </div>
          <div class="card-caption">検査結果からの自動インサイト</div>
        </div>
        ${analysisHtml}
      </section>
    </main>
  </div>
</body>
</html>`;

  return finalHtml;
}

// ======================
// 色付けテーブル生成
// ======================
function generateTableHtml(dataRows) {
  let tableRowsHtml = '';
  let cardsHtml = '';

  dataRows.forEach(item => {
    const defectCount = parseInt(item['不良数'], 10) || 0;
    const hasComment = item['コメント'] && item['コメント'].trim() !== '';
    const rate = item['不良率'] || '';
    const countClass = defectCount > 0 ? 'highlight-ng' : '';
    const rateClass = defectCount > 0 ? 'highlight-ng' : '';
    const commentClass = defectCount > 0 || hasComment ? 'highlight-ng' : '';

    tableRowsHtml += `
      <tr>
        <td>${formatDateSlash(item['セット日'])}</td>
        <td>${item['セット者'] || ''}</td>
        <td>${item['機番'] || ''}</td>
        <td>${item['客先'] || ''}</td>
        <td>${item['品番'] || ''}</td>
        <td>${item['品名'] || ''}</td>
        <td class="number-cell ${countClass}">${(item['検査数'] || 0).toLocaleString()}</td>
        <td class="number-cell ${countClass}">${(item['不良数'] || 0).toLocaleString()}</td>
        <td class="number-cell ${rateClass}">${rate}</td>
        <td class="${commentClass}">${item['コメント'] || ''}</td>
      </tr>
    `;

    // カード表示ではOK時の緑を廃止し、NG時のみ赤にする
    const labelCountClass = defectCount > 0 ? 'metric-ng' : '';
    const labelRateClass = defectCount > 0 ? 'metric-ng' : '';
    const labelCommentClass = defectCount > 0 ? 'metric-ng' : '';
    const valueCountClass = defectCount > 0 ? 'metric-ng' : '';
    const valueRateClass = defectCount > 0 ? 'metric-ng' : '';
    const valueCommentClass = defectCount > 0 || hasComment ? 'metric-ng' : '';

    cardsHtml += `
      <div class="record-card">
        <div class="record-row"><span class="record-label">セット日</span><span class="record-value">${formatDateSlash(item['セット日'])}</span></div>
        <div class="record-row"><span class="record-label">セット者</span><span class="record-value">${item['セット者'] || ''}</span></div>
        <div class="record-row"><span class="record-label">機番</span><span class="record-value">${item['機番'] || ''}</span></div>
        <div class="record-row"><span class="record-label">客先</span><span class="record-value">${item['客先'] || ''}</span></div>
        <div class="record-row"><span class="record-label">品番</span><span class="record-value">${item['品番'] || ''}</span></div>
        <div class="record-row"><span class="record-label">品名</span><span class="record-value">${item['品名'] || ''}</span></div>
        <div class="record-row"><span class="record-label ${labelCountClass}">検査数</span><span class="record-value ${valueCountClass}">${(item['検査数'] || 0).toLocaleString()}</span></div>
        <div class="record-row"><span class="record-label ${labelCountClass}">不良数</span><span class="record-value ${valueCountClass}">${(item['不良数'] || 0).toLocaleString()}</span></div>
        <div class="record-row"><span class="record-label ${labelRateClass}">不良率</span><span class="record-value ${valueRateClass}">${rate}</span></div>
        <div class="record-row"><span class="record-label ${labelCommentClass}">不具合内容</span><span class="record-value ${valueCommentClass}">${item['コメント'] || ''}</span></div>
      </div>
    `;
  });

  const tableHtml = `
    <div class="table-block">
      <div class="table-scroll">
        <table class="report-table">
          <thead>
            <tr>
              <th>セット日</th>
              <th>セット者</th>
              <th>機番</th>
              <th>客先</th>
              <th>品番</th>
              <th>品名</th>
              <th>検査数</th>
              <th>不良数</th>
              <th>不良率</th>
              <th>不具合内容</th>
            </tr>
          </thead>
          <tbody>${tableRowsHtml}</tbody>
        </table>
      </div>
    </div>
  `;

  const cardsBlock = `<div class="cards-block"><div class="records-wrap">${cardsHtml}</div></div>`;

  return tableHtml + cardsBlock;
}

// ======================
// Google Driveにアップロード
// ======================
function uploadToDrive(fileName, htmlContent, DRIVE_FOLDER_ID) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    
    // 同じファイル名の既存ファイルを検索して削除
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      try {
        existingFile.setTrashed(true);
        Logger.log(`既存ファイルを削除: ${fileName}`);
        Utilities.sleep(300); // 削除処理の完了を待つ
      } catch (deleteError) {
        Logger.log(`既存ファイル削除エラー: ${deleteError.toString()}`);
      }
    }
    
    // 新しいファイルをアップロード
    const blob = Utilities.newBlob(htmlContent, 'text/html', fileName);
    const file = folder.createFile(blob);
    
    Logger.log(`ファイルをアップロードしました: ${fileName}`);
    return file.getId();
  } catch (error) {
    Logger.log("Google Driveアップロードエラー: " + error.toString());
    throw error;
  }
}

// ======================
// ARAICHAT送信関連関数
// ======================

/**
 * ファイルのハッシュ値を計算（重複チェック用）
 */
function calculateFileDigest(fileData, fileName) {
  const combined = Utilities.newBlob(fileData + fileName + fileData.length, 'text/plain').getBytes();
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined);
  return hash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * 送信履歴キャッシュを読み込む
 */
function loadSentCache() {
  const cache = CacheService.getScriptCache();
  const cacheData = cache.get('araichat_sent_cache');
  
  if (!cacheData) {
    // PropertiesServiceからも試す
    const props = PropertiesService.getScriptProperties();
    const propsData = props.getProperty('araichat_sent_cache');
    if (propsData) {
      try {
        return JSON.parse(propsData);
      } catch (e) {
        Logger.log("送信履歴キャッシュ読み込みエラー: " + e.toString());
        return {};
      }
    }
    return {};
  }
  
  try {
    const parsed = JSON.parse(cacheData);
    const currentTime = new Date().getTime();
    const ttlMs = CACHE_TTL_HOURS * 3600 * 1000;
    
    // 古いエントリを削除
    const filtered = {};
    for (const key in parsed) {
      if (currentTime - parsed[key].sentTime < ttlMs) {
        filtered[key] = parsed[key];
      }
    }
    return filtered;
  } catch (e) {
    Logger.log("送信履歴キャッシュ読み込みエラー: " + e.toString());
    return {};
  }
}

/**
 * 送信履歴キャッシュを保存
 */
function saveSentCache(cache) {
  try {
    const cacheData = JSON.stringify(cache);
    const cacheService = CacheService.getScriptCache();
    
    // CacheServiceは最大100KB、大きい場合はPropertiesServiceを使用
    if (cacheData.length < 90000) {
      cacheService.put('araichat_sent_cache', cacheData, 21600); // 6時間
    } else {
      // 大きすぎる場合はPropertiesServiceを使用
      PropertiesService.getScriptProperties().setProperty('araichat_sent_cache', cacheData);
    }
  } catch (e) {
    Logger.log("送信履歴キャッシュ保存エラー: " + e.toString());
  }
}

/**
 * ファイルが既に送信済みかチェック
 */
function checkAlreadySent(fileDigest, fileName) {
  const cache = loadSentCache();
  
  if (cache[fileDigest]) {
    const sentInfo = cache[fileDigest];
    const sentTime = new Date(sentInfo.sentTime);
    Logger.log(`既に送信済みとしてスキップ: ${fileName}`);
    Logger.log(`前回送信日時: ${Utilities.formatDate(sentTime, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss')}`);
    return true;
  }
  return false;
}

/**
 * ファイルを送信済みとしてマーク
 */
function markAsSent(fileDigest, fileName) {
  const cache = loadSentCache();
  cache[fileDigest] = {
    fileName: fileName,
    sentTime: new Date().getTime()
  };
  saveSentCache(cache);
}

/**
 * ARAICHATにファイルを送信
 */
function sendFileToAraichat(fileData, fileName) {
  // スクリプトプロパティから設定値を取得
  const ARAICHAT_BASE_URL = PropertiesService.getScriptProperties().getProperty('ARAICHAT_BASE_URL');
  const ARAICHAT_API_KEY = PropertiesService.getScriptProperties().getProperty('ARAICHAT_API_KEY');
  const ARAICHAT_ROOM_ID = PropertiesService.getScriptProperties().getProperty('ARAICHAT_ROOM_ID');
  
  if (!ARAICHAT_BASE_URL || !ARAICHAT_API_KEY || !ARAICHAT_ROOM_ID) {
    Logger.log("ARAICHAT設定がスクリプトプロパティに設定されていません");
    return false;
  }
  
  // ファイルのハッシュ値を計算
  const fileDigest = calculateFileDigest(fileData, fileName);
  
  // 既に送信済みかチェック
  if (checkAlreadySent(fileDigest, fileName)) {
    Logger.log(`既に送信済みのためスキップ: ${fileName}`);
    return true;
  }
  
  const baseUrl = ARAICHAT_BASE_URL.replace(/\/$/, '');
  const url = `${baseUrl}/api/integrations/send/${ARAICHAT_ROOM_ID}`;
  
  const headers = {
    'Authorization': `Bearer ${ARAICHAT_API_KEY}`,
    'Idempotency-Key': `gdrive:${fileDigest.substring(0, 32)}`
  };
  
  const data = {
    'text': `セット品検査結果報告`
  };
  
  Logger.log(`ARAICHATへファイル送信開始: ${fileName}`);
  
  // リトライ設定
  const maxRetries = 3;
  const backoffSeconds = 2;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      if (attempt > 1) {
        const waitTime = backoffSeconds * Math.pow(2, attempt - 2);
        Logger.log(`リトライ ${attempt}/${maxRetries}（${waitTime}秒待機後）...`);
        Utilities.sleep(waitTime * 1000);
      }
      
      // ファイル名のエンコードを確実にするため、マルチパートを手動構築する
      const boundary = '----araichat_boundary_' + Utilities.getUuid();
      const delimiter = `--${boundary}`;
      const crlf = '\r\n';
      // filename にも UTF-8 のままの文字列を入れる（サーバーが filename* を無視する場合に備え）
      const originalName = fileName || 'file.html';
      const encodedNameStar = `utf-8''${encodeURIComponent(originalName)}`;

      const chunks = [];
      const pushString = (str) => chunks.push(Utilities.newBlob(str, 'text/plain').getBytes());
      const concatBytes = (arrays) => {
        let total = 0;
        arrays.forEach(a => total += a.length);
        const merged = new Uint8Array(total);
        let offset = 0;
        arrays.forEach(a => {
          merged.set(a, offset);
          offset += a.length;
        });
        return merged;
      };

      // textパート
      pushString([
        delimiter,
        'Content-Disposition: form-data; name="text"',
        '',
        data.text,
        ''
      ].join(crlf));

      // fileパート（filename* で UTF-8 を明示、filename で ASCII フォールバック）
      pushString([
        delimiter,
        `Content-Disposition: form-data; name="files"; filename="${originalName}"; filename*=${encodedNameStar}`,
        'Content-Type: text/html',
        '',
        ''
      ].join(crlf));
      chunks.push(fileData); // ファイル本体
      // ファイル本体の後に区切りのCRLFを入れる
      pushString(crlf);

      // 終端
      pushString(`${delimiter}--${crlf}`);

      const multipartBody = concatBytes(chunks);

      Logger.log(`送信ファイル名: ${fileName}`);
      Logger.log(`マルチパートサイズ: ${multipartBody.length} bytes`);
      
      // バイト配列をBlobに変換して送信（contentTypeにboundaryを明示）
      const multipartBlob = Utilities.newBlob(multipartBody, `multipart/form-data; boundary=${boundary}`);
      
      const options = {
        'method': 'post',
        'headers': headers,
        'payload': multipartBlob,
        'muteHttpExceptions': true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log(`レスポンスステータス: ${responseCode}`);
      
      if (responseCode >= 200 && responseCode < 300) {
        Logger.log(`ARAICHATへファイル送信成功: ${fileName}`);
        
        // 送信成功時のみ送信履歴に記録
        markAsSent(fileDigest, fileName);
        return true;
      } else if (responseCode >= 500 && responseCode < 600 && attempt < maxRetries) {
        // サーバーエラーの場合はリトライ
        Logger.log(`HTTP ${responseCode} エラー: ${responseText} - リトライします`);
        continue;
      } else {
        // クライアントエラーや最終リトライ失敗時
        Logger.log(`ARAICHAT送信HTTPエラー: ${responseCode} - ${responseText}`);
        return false;
      }
      
    } catch (e) {
      if (attempt < maxRetries) {
        Logger.log(`ネットワークエラー: ${e.toString()} - リトライします`);
        continue;
      } else {
        Logger.log(`ARAICHAT送信エラー（全リトライ試行後）: ${e.toString()}`);
        return false;
      }
    }
  }
  
  return false;
}

/**
 * Google Driveフォルダ内のHTMLファイルをARAICHATに送信
 */
function sendHtmlFilesToAraichat(folderId) {
  try {
    // スクリプトプロパティから削除設定を取得
    const deleteAfterUpload = PropertiesService.getScriptProperties().getProperty('DELETE_AFTER_UPLOAD') === 'true';
    
    Logger.log(`=== ARAICHAT送信処理開始 ===`);
    Logger.log(`フォルダID: ${folderId}`);
    Logger.log(`削除モード: ${deleteAfterUpload ? '有効' : '無効（ファイル保持）'}`);
    
    const folder = DriveApp.getFolderById(folderId);
    if (!folder) {
      Logger.log("フォルダが見つかりません: " + folderId);
      return;
    }
    
    // HTMLファイルを検索
    const files = folder.getFilesByType(MimeType.HTML);
    const fileList = [];
    while (files.hasNext()) {
      fileList.push(files.next());
    }
    
    if (fileList.length === 0) {
      Logger.log("送信対象のHTMLファイルはありませんでした");
      return;
    }
    
    Logger.log(`送信対象HTMLファイル: ${fileList.length}件`);
    
    let sentCount = 0;
    let failedCount = 0;
    let deletedCount = 0;
    
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      const fileName = file.getName();
      const fileId = file.getId();
      
      Logger.log(`[${i + 1}/${fileList.length}] 送信中: ${fileName}`);
      
      try {
        // ファイルをダウンロード
        const fileData = file.getBlob().getBytes();
        
        // ARAICHATに送信
        const result = sendFileToAraichat(fileData, fileName);
        
        if (result) {
          sentCount++;
          Logger.log(`送信成功: ${fileName}`);
          
          // 送信成功時の処理
          if (deleteAfterUpload) {
            try {
              file.setTrashed(true);
              deletedCount++;
              Logger.log(`ファイル削除成功: ${fileName}`);
            } catch (deleteError) {
              Logger.log(`ファイル削除失敗: ${fileName} - ${deleteError.toString()}`);
            }
          }
        } else {
          failedCount++;
          Logger.log(`送信失敗: ${fileName}`);
        }
        
        // 送信間隔を空ける（API制限対策）
        if (i < fileList.length - 1) {
          Utilities.sleep(2000); // 2秒待機
        }
        
      } catch (error) {
        failedCount++;
        Logger.log(`処理エラー: ${fileName} - ${error.toString()}`);
      }
    }
    
    Logger.log(`=== 送信結果 ===`);
    Logger.log(`成功: ${sentCount}件`);
    Logger.log(`失敗: ${failedCount}件`);
    Logger.log(`削除: ${deletedCount}件`);
    Logger.log(`合計: ${fileList.length}件`);
    
  } catch (error) {
    Logger.log("ARAICHAT送信処理エラー: " + error.toString());
  }
}
