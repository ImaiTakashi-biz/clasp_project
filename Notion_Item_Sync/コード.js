// ======================
// 設定（ここに直接書く）
// ======================
const NOTION_VERSION = "2022-06-28";

// セキュリティ向上：トークンはプロパティストアから取得
// 初回のみ setupCredentials() を実行してトークンを保存してください
const DB1_ID = "2ae37bffefe880b7b6cacbd899c96b1f"; // 刃物依頼DB
const DB2_ID = "27837bffefe88033a039e993538cf379"; // 仕入品マスタDB

/**
 * 初回セットアップ用：トークンをプロパティストアに保存
 * 一度だけ実行し、その後はこの関数内のトークンを削除してください
 */
function setupCredentials() {
  const token = ""; // ← 保存後は削除！
  PropertiesService.getScriptProperties().setProperty("NOTION_TOKEN", token);
  Logger.log("トークンを保存しました。コード内のトークンは削除してください。");
}

/**
 * トークン取得用ヘルパー
 */
function getNotionToken() {
  const token = PropertiesService.getScriptProperties().getProperty("NOTION_TOKEN");
  if (!token) {
    throw new Error("NOTION_TOKEN が未設定です。setupCredentials() を実行してください。");
  }
  return token;
}

/**
 * メイン処理（統合版：DB①→DB② と DB②→DB① の両方向を処理）
 * 
 * 【DB① → DB②】
 * 1. DB① 同期フラグ = TRUE のページ取得
 * 2. 刃物品番・数量・依頼日を取得
 * 3. DB② の品番 = 刃物品番 のページ検索
 * 4. 一致ページの 数量・依頼日 を更新
 * 5. DB① の同期フラグを FALSE に戻す
 * 
 * 【DB② → DB①】
 * 1. DB② 同期フラグ = TRUE のページ取得
 * 2. 品番を取得
 * 3. DB① の刃物品番 = 品番 のページ検索
 * 4. 一致ページの 数量・依頼日 をクリア更新
 * 5. DB② の同期フラグを FALSE に戻す
 */
function syncDb1ToDb2() {
  const token = getNotionToken();

  Logger.log("DB1_ID=" + DB1_ID);
  Logger.log("DB2_ID=" + DB2_ID);

  if (!DB1_ID || !DB2_ID) {
    throw new Error("DB1_ID / DB2_ID が未設定です。（定数を確認してください）");
  }

  let totalSuccessCount = 0;
  let totalErrorCount = 0;

  // ======================
  // 処理1: DB① → DB② の同期
  // ======================
  Logger.log("=== DB① → DB② の同期を開始 ===");
  const pages1 = queryDb1PagesToSync(token, DB1_ID);

  if (pages1.length > 0) {
    let successCount1 = 0;
    let errorCount1 = 0;

    pages1.forEach(page => {
      const pageId = page.id;
      const p = page.properties;

      // 2. 刃物品番・数量・依頼日 を取得
      const toolNumber = getTitleText(p["刃物品番"]); // タイトル
      const qty = p["数量"] && p["数量"].number;
      const requestDate = p["依頼日"] && p["依頼日"].date; // {start: "...", end: null}

      if (!toolNumber) {
        Logger.log(`ページ ${pageId} は刃物品番が空のためスキップ`);
        return;
      }
      if (qty == null) {
        Logger.log(`ページ ${pageId} は数量が未入力のためスキップ（必要なければここは無視してもOK）`);
        // 必須でないならここはコメントアウトしてOK
        // return;
      }
      if (!requestDate) {
        Logger.log(`ページ ${pageId} は依頼日が未入力のためスキップ`);
        return;
      }

      // 3. DB② で 品番 = 刃物品番 のページを検索
      const targetPages = findDb2PagesByHinmoku(token, DB2_ID, toolNumber);
      if (!targetPages.length) {
        Logger.log(`DB②に品番 = ${toolNumber} のページが見つかりませんでした。`);
        return;
      }

      // 品番が一意と仮定して1件目だけ更新（複数あればループしてもOK）
      const targetPage = targetPages[0];

      // 4. DB②の 数量・依頼日 を更新（成功時のみ次へ進む）
      const updateSuccess = updateDb2Page(token, targetPage.id, qty, requestDate);
      
      if (updateSuccess) {
        // 5. DB①の同期フラグを FALSE に戻す（DB②更新成功時のみ）
        const resetSuccess = resetSyncFlagInDb1(token, pageId);
        if (resetSuccess) {
          successCount1++;
          Logger.log(`品番 ${toolNumber} の同期が完了しました。`);
        } else {
          errorCount1++;
          Logger.log(`品番 ${toolNumber} の同期フラグリセットに失敗しました。`);
        }
      } else {
        errorCount1++;
        Logger.log(`品番 ${toolNumber} のDB②更新に失敗しました。同期フラグはリセットしません。`);
      }
    });

    Logger.log(`DB① → DB② 同期完了: 成功=${successCount1}, 失敗=${errorCount1}`);
    totalSuccessCount += successCount1;
    totalErrorCount += errorCount1;
  } else {
    Logger.log("DB① → DB②: 同期対象ページはありませんでした。");
  }

  // ======================
  // 処理2: DB② → DB① の同期（クリア処理）
  // ======================
  Logger.log("=== DB② → DB① の同期を開始 ===");
  const pages2 = queryDb2PagesToSync(token, DB2_ID);

  if (pages2.length > 0) {
    let successCount2 = 0;
    let errorCount2 = 0;

    pages2.forEach(page => {
      const pageId = page.id;
      const p = page.properties;

      // 2. 品番を取得
      const hinmokuNumber = getRichTextValue(p["品番"]); // rich_text プロパティ

      if (!hinmokuNumber) {
        Logger.log(`ページ ${pageId} は品番が空のためスキップ`);
        return;
      }

      // 3. DB① で 刃物品番 = 品番 のページを検索
      const targetPages = findDb1PagesByToolNumber(token, DB1_ID, hinmokuNumber);
      if (!targetPages.length) {
        Logger.log(`DB①に刃物品番 = ${hinmokuNumber} のページが見つかりませんでした。`);
        return;
      }

      // 刃物品番が一意と仮定して1件目だけ更新（複数あればループしてもOK）
      const targetPage = targetPages[0];

      // 4. DB①の 数量・依頼日 をクリア更新（成功時のみ次へ進む）
      const clearSuccess = clearDb1Page(token, targetPage.id);
      
      if (clearSuccess) {
        // 5. DB②の同期フラグを FALSE に戻す（DB①クリア成功時のみ）
        const resetSuccess = resetSyncFlagInDb2(token, pageId);
        if (resetSuccess) {
          successCount2++;
          Logger.log(`品番 ${hinmokuNumber} のクリアが完了しました。`);
        } else {
          errorCount2++;
          Logger.log(`品番 ${hinmokuNumber} の同期フラグリセットに失敗しました。`);
        }
      } else {
        errorCount2++;
        Logger.log(`品番 ${hinmokuNumber} のDB①クリアに失敗しました。同期フラグはリセットしません。`);
      }
    });

    Logger.log(`DB② → DB① 同期完了: 成功=${successCount2}, 失敗=${errorCount2}`);
    totalSuccessCount += successCount2;
    totalErrorCount += errorCount2;
  } else {
    Logger.log("DB② → DB①: 同期対象ページはありませんでした。");
  }

  Logger.log(`=== 全体の同期完了: 成功=${totalSuccessCount}, 失敗=${totalErrorCount} ===`);
}

// ======================
// DB①: 同期対象ページ取得（同期フラグ = TRUE）
// ======================
function queryDb1PagesToSync(token, dbId) {
  const url = `https://api.notion.com/v1/databases/${dbId}/query`;

  const payload = {
    filter: {
      property: "同期フラグ",
      checkbox: {
        equals: true
      }
    }
  };

  const options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  const json = JSON.parse(res.getContentText());

  if (statusCode !== 200 || json.object === "error") {
    Logger.log("DB① query error: " + res.getContentText());
    throw new Error("DB① query error: " + (json.message || res.getContentText()));
  }

  return json.results || [];
}

// ======================
// DB②: 品番 = 刃物品番 で検索
// 品番 は「テキストプロパティ（rich_text）」想定
// ======================
function findDb2PagesByHinmoku(token, dbId, toolNumber) {
  const url = `https://api.notion.com/v1/databases/${dbId}/query`;

  const payload = {
    filter: {
      property: "品番",
      rich_text: {
        equals: toolNumber
      }
    }
  };

  const options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  const json = JSON.parse(res.getContentText());

  if (statusCode !== 200 || json.object === "error") {
    Logger.log("DB② query error: " + res.getContentText());
    throw new Error("DB② query error: " + (json.message || res.getContentText()));
  }

  return json.results || [];
}

// ======================
// DB②: 数量・依頼日 を更新
// 戻り値: 成功時 true, 失敗時 false
// ======================
function updateDb2Page(token, pageId, qty, requestDate) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const props = {};

  // 数量（数値プロパティ）
  if (qty != null) {
    props["数量"] = {
      number: qty
    };
  }

  // 依頼日（日付プロパティ）: DB①の date オブジェクトをそのまま渡す
  if (requestDate) {
    props["依頼日"] = {
      date: requestDate
    };
  }

  const payload = {
    properties: props
  };

  const options = {
    method: "patch",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  
  Logger.log("update DB②: " + statusCode + " / " + res.getContentText());

  if (statusCode !== 200) {
    Logger.log(`DB② 更新失敗 (pageId: ${pageId}): ステータス ${statusCode}`);
    return false;
  }

  const json = JSON.parse(res.getContentText());
  if (json.object === "error") {
    Logger.log(`DB② 更新エラー (pageId: ${pageId}): ${json.message}`);
    return false;
  }

  return true;
}

// ======================
// DB①: 同期フラグを FALSE に戻す
// 戻り値: 成功時 true, 失敗時 false
// ======================
function resetSyncFlagInDb1(token, pageId) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const payload = {
    properties: {
      "同期フラグ": {
        checkbox: false
      }
    }
  };

  const options = {
    method: "patch",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  
  Logger.log("reset 同期フラグ: " + statusCode + " / " + res.getContentText());

  if (statusCode !== 200) {
    Logger.log(`同期フラグリセット失敗 (pageId: ${pageId}): ステータス ${statusCode}`);
    return false;
  }

  const json = JSON.parse(res.getContentText());
  if (json.object === "error") {
    Logger.log(`同期フラグリセットエラー (pageId: ${pageId}): ${json.message}`);
    return false;
  }

  return true;
}

// ======================
// DB②: 同期対象ページ取得（同期フラグ = TRUE）
// ======================
function queryDb2PagesToSync(token, dbId) {
  const url = `https://api.notion.com/v1/databases/${dbId}/query`;

  const payload = {
    filter: {
      property: "同期フラグ",
      checkbox: {
        equals: true
      }
    }
  };

  const options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  const json = JSON.parse(res.getContentText());

  if (statusCode !== 200 || json.object === "error") {
    Logger.log("DB② query error: " + res.getContentText());
    throw new Error("DB② query error: " + (json.message || res.getContentText()));
  }

  return json.results || [];
}

// ======================
// DB①: 刃物品番 = 品番 で検索
// 刃物品番 は「タイトルプロパティ（title）」想定
// ======================
function findDb1PagesByToolNumber(token, dbId, hinmokuNumber) {
  const url = `https://api.notion.com/v1/databases/${dbId}/query`;

  const payload = {
    filter: {
      property: "刃物品番",
      title: {
        equals: hinmokuNumber
      }
    }
  };

  const options = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  const json = JSON.parse(res.getContentText());

  if (statusCode !== 200 || json.object === "error") {
    Logger.log("DB① query error: " + res.getContentText());
    throw new Error("DB① query error: " + (json.message || res.getContentText()));
  }

  return json.results || [];
}

// ======================
// DB①: 数量・依頼日 をクリア（nullに設定）
// 戻り値: 成功時 true, 失敗時 false
// ======================
function clearDb1Page(token, pageId) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const props = {};

  // 数量をクリア（nullを設定）
  props["数量"] = {
    number: null
  };

  // 依頼日をクリア（nullを設定）
  props["依頼日"] = {
    date: null
  };

  const payload = {
    properties: props
  };

  const options = {
    method: "patch",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  
  Logger.log("clear DB①: " + statusCode + " / " + res.getContentText());

  if (statusCode !== 200) {
    Logger.log(`DB① クリア失敗 (pageId: ${pageId}): ステータス ${statusCode}`);
    return false;
  }

  const json = JSON.parse(res.getContentText());
  if (json.object === "error") {
    Logger.log(`DB① クリアエラー (pageId: ${pageId}): ${json.message}`);
    return false;
  }

  return true;
}

// ======================
// DB②: 同期フラグを FALSE に戻す
// 戻り値: 成功時 true, 失敗時 false
// ======================
function resetSyncFlagInDb2(token, pageId) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const payload = {
    properties: {
      "同期フラグ": {
        checkbox: false
      }
    }
  };

  const options = {
    method: "patch",
    headers: {
      "Authorization": "Bearer " + token,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const statusCode = res.getResponseCode();
  
  Logger.log("reset 同期フラグ(DB②): " + statusCode + " / " + res.getContentText());

  if (statusCode !== 200) {
    Logger.log(`同期フラグリセット失敗 (pageId: ${pageId}): ステータス ${statusCode}`);
    return false;
  }

  const json = JSON.parse(res.getContentText());
  if (json.object === "error") {
    Logger.log(`同期フラグリセットエラー (pageId: ${pageId}): ${json.message}`);
    return false;
  }

  return true;
}

// ======================
// 共通: タイトルプロパティからテキストを取得
// ======================
function getTitleText(prop) {
  if (!prop || prop.type !== "title") return "";
  const arr = prop.title || [];
  return arr.map(t => t.plain_text).join("");
}

// ======================
// 共通: rich_textプロパティからテキストを取得
// ======================
function getRichTextValue(prop) {
  if (!prop || prop.type !== "rich_text") return "";
  const arr = prop.rich_text || [];
  return arr.map(t => t.plain_text).join("");
}
