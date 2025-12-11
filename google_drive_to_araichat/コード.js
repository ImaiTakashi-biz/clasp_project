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
    const DRIVE_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
    
    if (!SPREADSHEET_ID) {
      Logger.log("SPREADSHEET_IDがスクリプトプロパティに設定されていません");
      return;
    }
    if (!DRIVE_FOLDER_ID) {
      Logger.log("DRIVE_FOLDER_IDがスクリプトプロパティに設定されていません");
      return;
    }
    
    // 1. Google DriveのHTMLファイルを検索して削除
    deleteHtmlFilesInDrive(DRIVE_FOLDER_ID);
    
    // 2. Google Sheetsからデータを取得
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
    
    // 3. 不良率をパーセントに変換
    const processedRows = rowsToProcess.map(row => {
      let rate = row["不良率"];
      if (rate === "" || rate === null || rate === undefined) {
        row["不良率"] = "0.0%";
      } else {
        row["不良率"] = (Number(rate) * 100).toFixed(1) + "%";
      }
      return row;
    });
    
    // 4. AIエージェント用のプロンプト生成とHTML生成
    const promptData = generatePromptForAI(processedRows);
    
    // 5. AI分析コメントを生成
    const aiCommentHtml = generateAIComment(promptData.dataForAgent);
    
    // 6. HTMLとAIコメントを結合してファイル生成
    const finalHtml = combineHtmlAndAIComment(promptData.initialHtml, aiCommentHtml, processedRows, promptData.fileName);
    
    // 7. Google Driveにアップロード
    uploadToDrive(promptData.fileName, finalHtml, DRIVE_FOLDER_ID);
    
    // 8. 送信フラグを「済」に更新
    processedRows.forEach(row => {
      sheet.getRange(row.row_number, headerIndexes["送信"] + 1).setValue("済");
    });
    
    Logger.log(`処理完了: ${processedRows.length}件のデータを処理しました`);
    
    // 9. ARAICHATにHTMLファイルを送信
    sendHtmlFilesToAraichat(DRIVE_FOLDER_ID);
    
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

以下の入力データ（JSON形式の検査結果データ）を元に、**検査結果のAI分析コメントのみ**を生成してください。

---

<h3>目的</h3>
入力データを分析し、**不良の根本原因特定と、その低減のための最も効果的かつ実践的な改善策を、簡潔かつ要点を絞って提供する**ため、熟練された上級位のプロのCNC自動旋盤オペレーター視点と統計品質管理の知見に基づいた、分かりやすい分析コメントを生成する。

---

<h3>出力ルール</h3>
<ol>
  <li>見出し：「AI分析コメント：」とすること。</li>
  <li>書式：<code>&lt;ul&gt;&lt;li&gt;〜&lt;/li&gt;&lt;/ul&gt;</code>の箇条書き形式で出力すること。</li>
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
  // テーブルを再構築（色付け付き）
  const generatedTableHtml = generateTableHtml(dataRows);
  
  // タイトル部分を抽出
  const titleMatch = initialHtml.match(/<h1>(.*?)<\/h1>/);
  const pageTitleHtml = titleMatch && titleMatch[0] ? titleMatch[0] : `<h1>検査結果報告</h1>`;
  
  // タイトルとテーブルを結合
  let combinedHtmlBody = pageTitleHtml + generatedTableHtml;
  
  // 最終HTMLを生成
  const finalHtml = `<!DOCTYPE html>
<html lang="ja">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>${fileName.replace('.html', '')}</title>
	<link rel="preconnect" href="https://fonts.googleapis.com">
	<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
	<link href="https://fonts.googleapis.com/css2?family=Harenosora&family=Noto+Sans+JP:wght@300;400;700&display=swap" rel="stylesheet">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
	
	<style>
		/* 全体的なスタイル */
		body {
			font-family: 'Harenosora', 'Noto Sans JP', sans-serif;
			line-height: 1.6;
			color: #333;
			background-color: #fcfaf7;
			padding: 30px;
			margin: 0 auto;
			max-width: 900px;
			box-shadow: 5px 5px 15px rgba(0,0,0,0.1);
			border: 2px dashed #a0d8b4;
			border-radius: 8px;
		}
		
		/* 見出しスタイル */
		h1, h2, h3, h4, h5, h6 {
			font-family: 'Harenosora', 'Noto Sans JP', cursive;
			color: #4a90e2;
			border-bottom: 2px solid #f2c94c;
			padding-bottom: 5px;
			margin-top: 30px;
		}
		
		/* テーブルスタイル */
		table {
			width: 100%;
			border-collapse: collapse;
			margin-top: 20px;
			box-shadow: 2px 2px 8px rgba(0,0,0,0.05);
			background-color: #fff;
		}
		th, td {
			border: 1px solid #dcdcdc;
			padding: 10px;
			text-align: left;
			color: black;	
		}
		thead th {
			background-color: #e6f7ff;
			color: #333;
		}
		
		/* リスト（箇条書き）スタイル */
		ul {
			list-style: none;
			padding: 0;
			margin-top: 15px;
		}
		ul li {
			position: relative;
			margin-bottom: 10px;
			background-color: #fff;
			padding: 10px 10px 10px 40px;
			border-radius: 5px;
			box-shadow: 1px 1px 5px rgba(0,0,0,0.03);
			border-left: 5px solid #6ab04c;
		}
		ul li::before {
			content: "\\f00c";
			font-family: "Font Awesome 5 Free";
			font-weight: 900;
			position: absolute;
			left: 12px;
			color: #6ab04c;
			top: 50%;
			transform: translateY(-50%);
		}
		h3 + ul li::before {
			content: "\\f0a1";
			color: #f2c94c;
		}
		
		/* 強調テキストの色 */
		strong {
			color: #e74c3c;	
		}
		
		/* 箇条書きのサブリスト */
		ul ul li {
			padding-left: 20px;
			border-left: none;
			background-color: #f5f5f5;
		}
		ul ul li::before {
			content: "-";
			font-family: sans-serif;
			font-weight: normal;
			color: #555;
			left: 5px;
		}

        /* レスポンシブ対応のテーブルスタイル */
        @media screen and (max-width: 767px) {
            table, thead, tbody, th, td, tr {
                display: block;
            }
            thead tr {
                position: absolute;
                top: -9999px;
                left: -9999px;
            }
            tr {
                border: 1px solid #ccc;
                margin-bottom: 10px;
            }
            td {
                border: none;
                border-bottom: 1px solid #eee;
                position: relative;
                padding-left: 50%;
                text-align: right;
            }
            td:before {
                position: absolute;
                top: 0px;
                left: 6px;
                width: 45%;
                padding-right: 10px;
                white-space: nowrap;
                text-align: left;
                font-weight: bold;
            }
            td:nth-of-type(1):before { content: "セット日"; }
            td:nth-of-type(2):before { content: "セット者"; }
            td:nth-of-type(3):before { content: "機番"; }
            td:nth-of-type(4):before { content: "客先"; }
            td:nth-of-type(5):before { content: "品番"; }
            td:nth-of-type(6):before { content: "品名"; }
            td:nth-of-type(7):before { content: "検査数"; }
            td:nth-of-type(8):before { content: "不良数"; }
            td:nth-of-type(9):before { content: "不良率"; }
            td:nth-of-type(10):before { content: "不具合内容"; }
        }
	</style>
</head>
<body>
	${combinedHtmlBody} <br><br> ${aiCommentHtml}
</body>
</html>`;
  
  return finalHtml;
}

// ======================
// 色付けテーブル生成
// ======================
function generateTableHtml(dataRows) {
  let tableRowsHtml = '';
  const tableHeaderHtml = `
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
	`;

  dataRows.forEach(item => {
    const 不良数 = parseInt(item['不良数'], 10);
    
    // 不良数に応じてテキストの色を決定
    let textColor = 'black';
    let commentColor = 'black';
    
    if (不良数 > 0) {
      textColor = 'red';
    } else if (不良数 === 0) {
      textColor = 'green';
    }

    // コメント欄に値がある場合は、不良数が0でも赤色にする
    if (不良数 === 0 && item['コメント'] && item['コメント'].trim() !== '') {
      commentColor = 'red';
    } else {
      commentColor = textColor;
    }

    tableRowsHtml += `
			<tr>
				<td>${formatDateSlash(item['セット日'])}</td>
				<td>${item['セット者'] || ''}</td>
				<td>${item['機番'] || ''}</td>
				<td>${item['客先'] || ''}</td>
				<td>${item['品番'] || ''}</td>
				<td>${item['品名'] || ''}</td>
				<td align="right" style="color: ${textColor};">${(item['検査数'] || 0).toLocaleString()}</td>
				<td align="right" style="color: ${textColor};">${(item['不良数'] || 0).toLocaleString()}</td>
				<td align="right" style="color: ${textColor};">${item['不良率'] || ''}</td>
				<td style="color: ${commentColor};">${item['コメント'] || ''}</td>
			</tr>
		`;
  });

  return `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: メイリオ, Meiryo, ＭＳ Ｐゴシック, MS PGothic, Arial, sans-serif; font-size: 14px; width: 100%;">
			${tableHeaderHtml}
			<tbody>${tableRowsHtml}</tbody>
		</table>`;
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
