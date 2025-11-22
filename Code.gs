/**
 * スプレッドシート「sheet」の投稿待ち行から内容と投稿日時を読み取り、
 * 期限が来ている未投稿のデータを 1 件だけ X（旧Twitter）へ投稿します。
 *
 * シート構成（ヘッダーは任意）:
 *   A列: 投稿本文
 *   B列: 文字数（自動計算、Twitter仕様）
 *   C列: noteリンク（任意、投稿後にリプライとして投稿される）
 *   D列: 日にち（1900/01/01 を 1 とするシリアル値）
 *   E列: 時（0〜23 の数値、空欄は 0 時扱い）
 *   F列: 分（0〜59 の数値、空欄は 0 分扱い）
 *   G列: 投稿URL（投稿完了時にツイートURLを格納）
 *   H列: スレッドグループID（同じIDを持つ行を連続投稿してスレッド化、空欄は単独投稿）
 *   I列: ステータス（pending/posting/posted/failed）
 *
 * 必須スクリプトプロパティ:
 *   X_API_KEY
 *   X_API_SECRET
 *   X_ACCESS_TOKEN
 *   X_ACCESS_TOKEN_SECRET
 *   （いずれも X API v2 ユーザーコンテキストで投稿権限を持つ値）
 *
 * 想定運用フロー:
 * 1. 上記フォーマットでシートへ投稿候補を入力し、E・F 列は空欄のままにする。
 * 2. C列にnoteのURLを入力すると、投稿後に自動的にリプライとして投稿される（任意）
 * 3. スレッド投稿したい場合は、H列に同じIDを入力する（例: thread001）
 * 4. 定期実行トリガーなどで postNextScheduledToX() を呼び出す。
 * 5. 実行ごとに期限を過ぎた未投稿行を先頭から探し、スレッドグループ全体を投稿・記録する。
 */
const CONFIG = Object.freeze({
  sheetName: 'sheet',
  dataStartRow: 2,
  columns: Object.freeze({
    content: 1,
    charCount: 2,      // 文字数（新規追加）
    noteUrl: 3,        // noteリンク（B列→C列へ移動）
    day: 4,            // 日にち（C列→D列へ移動）
    hour: 5,           // 時（D列→E列へ移動）
    minute: 6,         // 分（E列→F列へ移動）
    postedUrl: 7,      // 投稿URL（F列→G列へ移動）
    threadGroupId: 8,  // スレッドグループID（G列→H列へ移動）
    status: 9          // ステータス（H列→I列へ移動）
  }),
  api: Object.freeze({
    tweetEndpoint: 'https://api.twitter.com/2/tweets'
  }),
  serialBaseDate: new Date(1899, 11, 31),
  // 文字数制限設定
  // プラン別の文字数制限:
  //   - free: 無料アカウント（140文字）
  //   - premium: Xプレミアム（25,000文字）
  characterLimit: Object.freeze({
    enabled: true,         // 文字数チェックを有効にする場合は true
    plan: 'free',          // 'free' (140文字) または 'premium' (25,000文字)
    skipOnExceed: true,    // 超過時にスキップする場合は true、エラーにする場合は false
    urlLength: 23          // URLは自動短縮され、1リンクあたり約23文字としてカウント
  }),
  // プラン別の文字数上限
  planLimits: Object.freeze({
    free: 140,
    premium: 25000
  }),
  // 実行制御設定
  execution: Object.freeze({
    toleranceMinutes: 5,    // 投稿時刻の許容範囲（±5分）
    lockTimeoutSeconds: 30  // ロック取得のタイムアウト（秒）
  }),
  // ステータス定義
  statuses: Object.freeze({
    pending: 'pending',     // 投稿待ち
    posting: 'posting',     // 投稿中
    posted: 'posted',       // 投稿完了
    failed: 'failed'        // 投稿失敗
  })
});

/**
 * X API v2を使って単一のツイートを投稿します。
 * @param {string} content - 投稿内容
 * @param {string|null} replyToTweetId - 返信先のツイートID（スレッド用、nullなら通常投稿）
 * @param {Object} credentials - API認証情報
 * @return {Object} - {tweetId: string, tweetUrl: string}
 */
function postSingleTweet_(content, replyToTweetId, credentials) {
  const { apiKey, apiSecret, accessToken, accessTokenSecret } = credentials;

  const method = 'POST';
  const url = CONFIG.api.tweetEndpoint;

  // ペイロード構築（reply用のパラメータを含む）
  const payloadObj = { text: content };
  if (replyToTweetId) {
    payloadObj.reply = { in_reply_to_tweet_id: replyToTweetId };
  }
  const payload = JSON.stringify(payloadObj);

  const buildOAuthHeader = () => {
    const oauthParams = {
      oauth_consumer_key: apiKey,
      oauth_nonce: Utilities.getUuid().replace(/-/g, ''),
      oauth_signature_method: 'HMAC-SHA1',
      oauth_timestamp: Math.floor(Date.now() / 1000).toString(),
      oauth_token: accessToken,
      oauth_version: '1.0'
    };

    const sortedParams = Object.keys(oauthParams)
      .sort()
      .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(oauthParams[key])}`)
      .join('&');

    const signatureBaseString = [
      method.toUpperCase(),
      encodeURIComponent(url),
      encodeURIComponent(sortedParams)
    ].join('&');

    const signingKey = `${encodeURIComponent(apiSecret)}&${encodeURIComponent(
      accessTokenSecret
    )}`;
    const signatureBytes = Utilities.computeHmacSignature(
      Utilities.MacAlgorithm.HMAC_SHA_1,
      signatureBaseString,
      signingKey
    );
    const signature = Utilities.base64Encode(signatureBytes);
    oauthParams.oauth_signature = signature;

    const headerParams = Object.keys(oauthParams)
      .sort()
      .map(
        (key) => `${encodeURIComponent(key)}="${encodeURIComponent(oauthParams[key])}"`
      )
      .join(', ');

    return `OAuth ${headerParams}`;
  };

  const response = UrlFetchApp.fetch(url, {
    method,
    muteHttpExceptions: true,
    headers: {
      Authorization: buildOAuthHeader(),
      'Content-Type': 'application/json'
    },
    payload
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();
  if (statusCode < 200 || statusCode >= 300) {
    let errorMessage = `X API エラー ${statusCode}: ${body}`;
    try {
      const errorJson = JSON.parse(body);
      if (errorJson.detail) {
        errorMessage = `X API エラー ${statusCode}: ${errorJson.detail}`;
      }
      if (
        statusCode === 403 &&
        errorJson &&
        typeof errorJson.detail === 'string' &&
        /application-only/i.test(errorJson.detail)
      ) {
        errorMessage += '（ユーザーコンテキストのアクセストークンを使用してください）';
      }
    } catch (parseError) {
      console.warn(`エラーレスポンスの解析に失敗しました: ${parseError}`);
    }
    throw new Error(errorMessage);
  }

  let parsed;
  try {
    parsed = JSON.parse(body);
  } catch (error) {
    throw new Error(`API レスポンスの解析に失敗しました: ${error}`);
  }

  const tweetId = parsed && parsed.data && parsed.data.id;
  if (!tweetId) {
    throw new Error(`API レスポンスにツイートIDが含まれていません: ${body}`);
  }

  const tweetUrl = `https://x.com/i/web/status/${tweetId}`;
  return { tweetId, tweetUrl };
}

/**
 * X（旧Twitter）の仕様に準拠した文字数をカウントします。
 * - 改行は1文字（半角扱い）
 * - 絵文字は1文字
 * - URLは自動短縮され、1リンクあたり約23文字としてカウント
 * - 画像・動画は文字数に影響しない（このスクリプトでは未対応）
 *
 * @param {string} text - カウント対象のテキスト
 * @return {number} X仕様での文字数
 */
function countTwitterCharacters_(text) {
  if (!text) {
    return 0;
  }

  let processedText = text;
  let urlCount = 0;

  // URLを検出して置換（http/https両方に対応）
  const urlPattern = /https?:\/\/[^\s]+/gi;
  const urls = processedText.match(urlPattern);

  if (urls) {
    urlCount = urls.length;
    // URLを一時的に削除
    processedText = processedText.replace(urlPattern, '');
  }

  // 残りのテキストの文字数をカウント
  // JavaScriptの.lengthは基本的に正確だが、サロゲートペア（絵文字など）も正しくカウント
  const textLength = Array.from(processedText).length;

  // URLの文字数を加算（1URLあたり23文字）
  const totalLength = textLength + (urlCount * CONFIG.characterLimit.urlLength);

  return totalLength;
}

/**
 * 期限を過ぎた未投稿データを X へ投稿し、E 列（投稿URL）と G 列（ステータス）を更新します。
 * スレッドグループIDが設定されている場合は、同じIDを持つ行をすべて連続投稿します。
 *
 * LockServiceで二重実行を防止します。
 */
function postNextScheduledToX() {
  // スクリプトロックを取得（二重実行防止）
  const lock = LockService.getScriptLock();

  try {
    // ロック取得を試みる（タイムアウト: 30秒）
    const hasLock = lock.tryLock(CONFIG.execution.lockTimeoutSeconds * 1000);

    if (!hasLock) {
      console.warn('別のプロセスが実行中のため、この実行をスキップします。');
      return;
    }

    console.log('ロックを取得しました。投稿処理を開始します。');

    const sheet = getTargetSheet_();
    const target = findNextPublishableRow_(sheet);

    if (!target) {
      console.log('投稿対象の行が見つかりませんでした。');
      return;
    }

    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('X_API_KEY');
    const apiSecret = props.getProperty('X_API_SECRET');
    const accessToken = props.getProperty('X_ACCESS_TOKEN');
    const accessTokenSecret = props.getProperty('X_ACCESS_TOKEN_SECRET');

    if (!apiKey || !apiSecret || !accessToken || !accessTokenSecret) {
      throw new Error(
        'スクリプトプロパティ X_API_KEY / X_API_SECRET / X_ACCESS_TOKEN / X_ACCESS_TOKEN_SECRET を設定してください。'
      );
    }

    const credentials = { apiKey, apiSecret, accessToken, accessTokenSecret };

    // スレッドグループIDがある場合は、同じIDを持つすべての行を取得
    if (target.threadGroupId) {
      postThreadGroup_(sheet, target, credentials);
    } else {
      postSingleRow_(sheet, target, credentials);
    }

  } catch (error) {
    console.error(`投稿処理中にエラーが発生しました: ${error}`);
    throw error;
  } finally {
    // ロックを解放
    lock.releaseLock();
    console.log('ロックを解放しました。');
  }
}

/**
 * 単一行を投稿します（スレッドではない通常投稿）
 */
function postSingleRow_(sheet, target, credentials) {
  const { rowNumber, content, noteUrl } = target;

  if (!content) {
    console.warn(`行 ${rowNumber} の本文が空欄のためスキップしました。`);
    setRowStatus_(sheet, rowNumber, CONFIG.statuses.failed, '本文が空欄');
    return;
  }

  // 投稿中ステータスを設定
  setRowStatus_(sheet, rowNumber, CONFIG.statuses.posting, '');

  // 文字数チェック
  if (CONFIG.characterLimit.enabled) {
    const maxLength = CONFIG.planLimits[CONFIG.characterLimit.plan];
    if (!maxLength) {
      throw new Error(`無効なプラン設定: ${CONFIG.characterLimit.plan}`);
    }

    const charCount = countTwitterCharacters_(content);
    if (charCount > maxLength) {
      const message = `文字数超過: ${charCount}文字 (制限${maxLength})`;

      if (CONFIG.characterLimit.skipOnExceed) {
        console.warn(`行 ${rowNumber} の文字数が制限を超えています（${charCount}文字 > ${maxLength}文字）。スキップします。`);
        setRowStatus_(sheet, rowNumber, CONFIG.statuses.failed, message);
        return;
      } else {
        setRowStatus_(sheet, rowNumber, CONFIG.statuses.failed, message);
        throw new Error(`行 ${rowNumber} の${message}`);
      }
    }
    console.log(`行 ${rowNumber} の文字数チェック OK（${charCount}/${maxLength}文字、プラン: ${CONFIG.characterLimit.plan}）`);
  }

  try {
    // メイン投稿
    const result = postSingleTweet_(content, null, credentials);
    sheet.getRange(rowNumber, CONFIG.columns.postedUrl).setValue(result.tweetUrl);
    setRowStatus_(sheet, rowNumber, CONFIG.statuses.posted, '');
    console.log(`行 ${rowNumber} を投稿しました: ${result.tweetUrl}`);

    // noteURLがある場合はリプライとして投稿
    if (noteUrl) {
      try {
        console.log(`行 ${rowNumber} のnoteURLをリプライとして投稿中...`);
        const noteResult = postSingleTweet_(noteUrl, result.tweetId, credentials);
        console.log(`noteURLリプライを投稿しました: ${noteResult.tweetUrl}`);
      } catch (noteError) {
        console.warn(`noteURLリプライの投稿に失敗しましたが、メイン投稿は成功しています: ${noteError}`);
      }
    }
  } catch (error) {
    console.error(`行 ${rowNumber} の投稿に失敗しました: ${error}`);
    setRowStatus_(sheet, rowNumber, CONFIG.statuses.failed, String(error));
    throw error;
  }
}

/**
 * スレッドグループを投稿します（同じスレッドグループIDを持つすべての行を連続投稿）
 */
function postThreadGroup_(sheet, firstTarget, credentials) {
  const { threadGroupId, noteUrl } = firstTarget;

  // 同じスレッドグループIDを持つすべての未投稿行を取得
  const threadRows = findThreadGroupRows_(sheet, threadGroupId);

  if (threadRows.length === 0) {
    console.warn(`スレッドグループ "${threadGroupId}" の投稿可能な行が見つかりませんでした。`);
    return;
  }

  console.log(`スレッドグループ "${threadGroupId}" を投稿開始（${threadRows.length}件）`);

  let previousTweetId = null;
  const results = [];

  for (let i = 0; i < threadRows.length; i++) {
    const row = threadRows[i];
    const { rowNumber, content } = row;

    if (!content) {
      console.warn(`行 ${rowNumber} の本文が空欄のためスキップしました。`);
      setRowStatus_(sheet, rowNumber, CONFIG.statuses.failed, '本文が空欄');
      continue;
    }

    // 投稿中ステータスを設定
    setRowStatus_(sheet, rowNumber, CONFIG.statuses.posting, '');

    // 文字数チェック
    if (CONFIG.characterLimit.enabled) {
      const maxLength = CONFIG.planLimits[CONFIG.characterLimit.plan];
      if (!maxLength) {
        throw new Error(`無効なプラン設定: ${CONFIG.characterLimit.plan}`);
      }

      const charCount = countTwitterCharacters_(content);
      if (charCount > maxLength) {
        const message = `文字数超過: ${charCount}文字 (制限${maxLength})`;
        console.warn(`行 ${rowNumber} の文字数が制限を超えています（${charCount}文字 > ${maxLength}文字）`);
        setRowStatus_(sheet, rowNumber, CONFIG.statuses.failed, message);

        if (CONFIG.characterLimit.skipOnExceed) {
          continue;
        } else {
          throw new Error(`行 ${rowNumber} の${message}`);
        }
      }
    }

    try {
      console.log(`スレッド ${i + 1}/${threadRows.length}: 行 ${rowNumber} を投稿中...`);
      const result = postSingleTweet_(content, previousTweetId, credentials);

      sheet.getRange(rowNumber, CONFIG.columns.postedUrl).setValue(result.tweetUrl);
      setRowStatus_(sheet, rowNumber, CONFIG.statuses.posted, '');

      console.log(`行 ${rowNumber} を投稿しました: ${result.tweetUrl}`);

      previousTweetId = result.tweetId;
      results.push(result);

      // API制限対策: 投稿間隔を空ける（最後の投稿以外）
      if (i < threadRows.length - 1) {
        Utilities.sleep(1000); // 1秒待機
      }
    } catch (error) {
      console.error(`行 ${rowNumber} の投稿に失敗しました: ${error}`);
      setRowStatus_(sheet, rowNumber, CONFIG.statuses.failed, String(error));
      throw error;
    }
  }

  console.log(`スレッドグループ "${threadGroupId}" の投稿完了（${results.length}件成功）`);

  // スレッド完了後、noteURLがある場合は最後のツイートにリプライ
  if (noteUrl && previousTweetId) {
    try {
      console.log(`スレッドグループのnoteURLをリプライとして投稿中...`);
      Utilities.sleep(1000); // API制限対策
      const noteResult = postSingleTweet_(noteUrl, previousTweetId, credentials);
      console.log(`noteURLリプライを投稿しました: ${noteResult.tweetUrl}`);
    } catch (noteError) {
      console.warn(`noteURLリプライの投稿に失敗しましたが、スレッド投稿は成功しています: ${noteError}`);
    }
  }
}

/**
 * 次回投稿対象となる行の情報を返します（投稿は実行しません）。
 * @return {{rowNumber:number, content:string, scheduledAt:any}|null}
 */
function previewNextScheduledRow() {
  const sheet = getTargetSheet_();
  const nextRow = findNextPublishableRow_(sheet);

  if (!nextRow) {
    console.log('次に投稿される行は見つかりません。');
    return null;
  }

  const timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const scheduledText = nextRow.scheduledAt
    ? Utilities.formatDate(nextRow.scheduledAt, timezone, 'yyyy-MM-dd HH:mm')
    : '日時未設定';

  console.log(
    `次の投稿予定: 行${nextRow.rowNumber} 本文="${nextRow.content}" 予定日時=${scheduledText}`
  );
  return nextRow;
}

/**
 * 指定されたスレッドグループIDを持つすべての未投稿行を取得します
 * @param {Sheet} sheet - 対象シート
 * @param {string} threadGroupId - スレッドグループID
 * @return {Array} - 行情報の配列
 */
function findThreadGroupRows_(sheet, threadGroupId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.dataStartRow) {
    return [];
  }

  const rowCount = lastRow - CONFIG.dataStartRow + 1;
  const range = sheet.getRange(
    CONFIG.dataStartRow,
    1,
    rowCount,
    CONFIG.columns.threadGroupId
  );
  const values = range.getValues();

  const threadRows = [];

  for (let index = 0; index < values.length; index += 1) {
    const rowNumber = CONFIG.dataStartRow + index;
    const row = values[index];
    const content = String(row[CONFIG.columns.content - 1] || '').trim();
    const status = String(row[CONFIG.columns.status - 1] || '').trim().toLowerCase();
    const rowThreadGroupId = String(row[CONFIG.columns.threadGroupId - 1] || '').trim();

    // 同じスレッドグループIDで未投稿の行を収集
    if (rowThreadGroupId === threadGroupId) {
      // ステータスチェック（posted/postingはスキップ）
      if (status === CONFIG.statuses.posted || status === CONFIG.statuses.posting) {
        continue;
      }

      if (content) {
        threadRows.push({
          rowNumber,
          content,
          threadGroupId: rowThreadGroupId
        });
      }
    }
  }

  return threadRows;
}

/**
 * セルからハイパーリンクの実際のURLを取得します
 * @param {Sheet} sheet - シート
 * @param {number} rowNumber - 行番号
 * @param {number} columnNumber - 列番号
 * @return {string} - 実際のURL（ハイパーリンクがない場合はセルの値）
 */
function extractUrlFromCell_(sheet, rowNumber, columnNumber) {
  const cell = sheet.getRange(rowNumber, columnNumber);

  // まずリッチテキストから取得を試みる
  const richText = cell.getRichTextValue();
  if (richText) {
    const linkUrl = richText.getLinkUrl();
    if (linkUrl) {
      return linkUrl;
    }
  }

  // 数式からHYPERLINK関数を検出
  const formula = cell.getFormula();
  if (formula) {
    const match = formula.match(/=HYPERLINK\("([^"]+)"/i);
    if (match && match[1]) {
      return match[1];
    }
  }

  // どちらでもない場合は通常の値を返す
  return String(cell.getValue() || '').trim();
}

function findNextPublishableRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.dataStartRow) {
    return null;
  }

  const rowCount = lastRow - CONFIG.dataStartRow + 1;
  const range = sheet.getRange(
    CONFIG.dataStartRow,
    1,
    rowCount,
    CONFIG.columns.status
  );
  const values = range.getValues();
  const now = new Date();

  for (let index = 0; index < values.length; index += 1) {
    const rowNumber = CONFIG.dataStartRow + index;
    const row = values[index];
    const content = String(row[CONFIG.columns.content - 1] || '').trim();
    const noteUrl = extractUrlFromCell_(sheet, rowNumber, CONFIG.columns.noteUrl);
    const dayValue = row[CONFIG.columns.day - 1];
    const hourValue = row[CONFIG.columns.hour - 1];
    const minuteValue = row[CONFIG.columns.minute - 1];
    const threadGroupId = String(row[CONFIG.columns.threadGroupId - 1] || '').trim();
    const status = String(row[CONFIG.columns.status - 1] || '').trim().toLowerCase();

    const scheduledAt = (() => {
      if (dayValue === '' || dayValue === null || dayValue === undefined) {
        return null;
      }

      const baseDate = (() => {
        if (dayValue instanceof Date) {
          const copy = new Date(dayValue.getTime());
          copy.setHours(0, 0, 0, 0);
          return copy;
        }

        const numeric = Number(dayValue);
        if (Number.isFinite(numeric)) {
          if (numeric < 1) {
            console.warn(`日にちのシリアル値は 1 以上で指定してください: ${dayValue}`);
            return null;
          }
          const base = new Date(CONFIG.serialBaseDate.getTime());
          base.setDate(base.getDate() + Math.floor(numeric));
          base.setHours(0, 0, 0, 0);
          return base;
        }

        const parsed = new Date(String(dayValue));
        if (Number.isNaN(parsed.getTime())) {
          console.warn(`日にちの値を解釈できませんでした: ${dayValue}`);
          return null;
        }
        const normalized = new Date(parsed.getTime());
        normalized.setHours(0, 0, 0, 0);
        return normalized;
      })();

      if (!baseDate) {
        return null;
      }

      const parseTimeCell = (value, label, max) => {
        if (value === '' || value === null || value === undefined) {
          return 0;
        }
        const parsed = parseInt(String(value).trim(), 10);
        if (Number.isNaN(parsed) || parsed < 0 || parsed > max) {
          console.warn(`${label} の値が不正です (0-${max} で指定してください): ${value}`);
          return null;
        }
        return parsed;
      };

      const hourParsed = parseTimeCell(hourValue, '時', 23);
      const minuteParsed = parseTimeCell(minuteValue, '分', 59);
      if (hourParsed === null || minuteParsed === null) {
        return null;
      }

      const scheduled = new Date(baseDate.getTime());
      scheduled.setHours(hourParsed, minuteParsed, 0, 0);
      if (Number.isNaN(scheduled.getTime())) {
        console.warn(
          `日時の組み立てに失敗しました: day=${dayValue}, hour=${hourValue}, minute=${minuteValue}`
        );
        return null;
      }

      return scheduled;
    })();

    if (!content) {
      continue;
    }

    // ステータスチェック（posted/postingはスキップ）
    if (status === CONFIG.statuses.posted || status === CONFIG.statuses.posting) {
      continue;
    }

    if (!scheduledAt) {
      continue;
    }

    // ±5分の許容窓を適用
    const toleranceMs = CONFIG.execution.toleranceMinutes * 60 * 1000;
    const earliestTime = now.getTime() - toleranceMs; // 5分前
    const latestTime = now.getTime() + toleranceMs;   // 5分後

    const scheduledTime = scheduledAt.getTime();

    // 予定時刻が範囲内にあるかチェック
    if (scheduledTime > latestTime) {
      // まだ早い（5分後より未来）
      continue;
    }

    if (scheduledTime < earliestTime) {
      // 古すぎる（5分前より過去）- 投稿漏れの可能性があるのでログ出力
      // ただし、投稿済み・失敗済みの場合は警告不要
      if (status !== CONFIG.statuses.failed) {
        console.warn(
          `行 ${rowNumber} の投稿時刻が古すぎます（予定: ${Utilities.formatDate(
            scheduledAt,
            Session.getScriptTimeZone(),
            'yyyy-MM-dd HH:mm'
          )}、ステータス: ${status || '未設定'}）`
        );
      }
      continue;
    }

    // 許容範囲内なので投稿対象
    return {
      rowNumber,
      content,
      noteUrl: noteUrl || null,
      scheduledAt,
      threadGroupId: threadGroupId || null
    };
  }

  return null;
}

function getTargetSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    throw new Error(
      `シート "${CONFIG.sheetName}" がアクティブなスプレッドシートに見つかりません。`
    );
  }
  return sheet;
}

/**
 * 行のステータスを設定します
 * @param {Sheet} sheet - 対象シート
 * @param {number} rowNumber - 行番号
 * @param {string} status - ステータス（pending/posting/posted/failed）
 * @param {string} errorMessage - エラーメッセージ（失敗時のみ）
 */
function setRowStatus_(sheet, rowNumber, status, errorMessage = '') {
  sheet.getRange(rowNumber, CONFIG.columns.status).setValue(status);

  // failedステータスの場合、エラーメッセージをE列にも記録
  if (status === CONFIG.statuses.failed && errorMessage) {
    sheet.getRange(rowNumber, CONFIG.columns.postedUrl).setValue(errorMessage);
  }
}

/**
 * 全行の文字数をチェックして、超過している行をログ出力します。
 * 投稿前の確認用ヘルパー関数です。
 * URLは23文字換算、絵文字も正確にカウントします。
 */
function checkAllCharacterCounts() {
  if (!CONFIG.characterLimit.enabled) {
    console.log('文字数チェックが無効化されています。');
    return;
  }

  const maxLength = CONFIG.planLimits[CONFIG.characterLimit.plan];
  if (!maxLength) {
    throw new Error(`無効なプラン設定: ${CONFIG.characterLimit.plan}`);
  }

  const sheet = getTargetSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.dataStartRow) {
    console.log('データが存在しません。');
    return;
  }

  const rowCount = lastRow - CONFIG.dataStartRow + 1;
  const range = sheet.getRange(CONFIG.dataStartRow, CONFIG.columns.content, rowCount, 1);
  const values = range.getValues();

  let totalCount = 0;
  let okCount = 0;
  let ngCount = 0;
  const ngRows = [];

  for (let i = 0; i < values.length; i++) {
    const rowNumber = CONFIG.dataStartRow + i;
    const content = String(values[i][0] || '').trim();

    if (!content) {
      continue;
    }

    totalCount++;
    const charCount = countTwitterCharacters_(content);

    if (charCount > maxLength) {
      ngCount++;
      ngRows.push({
        row: rowNumber,
        count: charCount,
        excess: charCount - maxLength,
        preview: content.substring(0, 30) + (content.length > 30 ? '...' : '')
      });
    } else {
      okCount++;
    }
  }

  console.log('=== 文字数チェック結果 ===');
  console.log(`プラン: ${CONFIG.characterLimit.plan}`);
  console.log(`制限: ${maxLength}文字`);
  console.log(`総行数: ${totalCount}行`);
  console.log(`OK: ${okCount}行`);
  console.log(`NG: ${ngCount}行`);
  console.log('※URLは23文字換算、絵文字も1文字としてカウント');

  if (ngRows.length > 0) {
    console.log('\n--- 文字数超過の行 ---');
    ngRows.forEach(item => {
      console.log(
        `行${item.row}: ${item.count}文字（${item.excess}文字超過） "${item.preview}"`
      );
    });
  } else {
    console.log('\nすべての行が文字数制限内です。');
  }

  return {
    plan: CONFIG.characterLimit.plan,
    limit: maxLength,
    total: totalCount,
    ok: okCount,
    ng: ngCount,
    ngRows: ngRows
  };
}

/**
 * スプレッドシートを開いたときに自動で実行される関数
 * カスタムメニューを追加します
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('X自動投稿')
    .addItem('選択行を1つのスレッドにまとめる', 'groupSelectedRowsAsThread')
    .addItem('スレッドIDをクリア', 'clearThreadIdFromSelection')
    .addSeparator()
    .addItem('文字数を更新', 'updateCharacterCounts')
    .addSeparator()
    .addItem('【初回のみ】文字数列を挿入', 'insertCharCountColumn')
    .addToUi();
}

/**
 * 選択されている行を1つのスレッドグループにまとめます
 * 選択範囲のすべての行に同じスレッドグループIDを設定します
 */
function groupSelectedRowsAsThread() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection) {
    SpreadsheetApp.getUi().alert('行を選択してください。');
    return;
  }

  // シート名チェック
  if (sheet.getName() !== CONFIG.sheetName) {
    SpreadsheetApp.getUi().alert(`"${CONFIG.sheetName}" シートで実行してください。`);
    return;
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const endRow = startRow + numRows - 1;

  // デバッグログ追加
  console.log(`選択範囲: 行${startRow}〜${endRow}（${numRows}行）`);

  // ヘッダー行を含む場合は調整
  let actualStartRow = startRow;
  let actualNumRows = numRows;

  if (startRow < CONFIG.dataStartRow) {
    // 選択範囲がヘッダー行を含む場合、データ行のみを対象にする
    actualStartRow = CONFIG.dataStartRow;
    actualNumRows = endRow - CONFIG.dataStartRow + 1;

    if (actualNumRows <= 0) {
      SpreadsheetApp.getUi().alert(`データ行（${CONFIG.dataStartRow}行目以降）を選択してください。`);
      return;
    }
  }

  // 最低2行以上選択されているかチェック
  if (actualNumRows < 2) {
    SpreadsheetApp.getUi().alert(
      `スレッドを作成するには2行以上選択してください。\n（現在: ${actualNumRows}行のデータ行が選択されています）`
    );
    return;
  }

  // 一意のスレッドグループIDを生成
  const threadGroupId = generateThreadGroupId_();

  // 選択範囲のG列（threadGroupId列）に同じIDを設定
  const threadIdRange = sheet.getRange(actualStartRow, CONFIG.columns.threadGroupId, actualNumRows, 1);
  const values = [];
  for (let i = 0; i < actualNumRows; i++) {
    values.push([threadGroupId]);
  }
  threadIdRange.setValues(values);

  SpreadsheetApp.getUi().alert(
    `${actualNumRows}行をスレッドグループ化しました。\nスレッドID: ${threadGroupId}`
  );

  console.log(`スレッドグループ作成: ${threadGroupId}（行${actualStartRow}〜${actualStartRow + actualNumRows - 1}）`);
}

/**
 * 選択されている行のスレッドグループIDをクリアします
 */
function clearThreadIdFromSelection() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection) {
    SpreadsheetApp.getUi().alert('行を選択してください。');
    return;
  }

  // シート名チェック
  if (sheet.getName() !== CONFIG.sheetName) {
    SpreadsheetApp.getUi().alert(`"${CONFIG.sheetName}" シートで実行してください。`);
    return;
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const endRow = startRow + numRows - 1;

  // デバッグログ追加
  console.log(`選択範囲: 行${startRow}〜${endRow}（${numRows}行）`);

  // ヘッダー行を含む場合は調整
  let actualStartRow = startRow;
  let actualNumRows = numRows;

  if (startRow < CONFIG.dataStartRow) {
    // 選択範囲がヘッダー行を含む場合、データ行のみを対象にする
    actualStartRow = CONFIG.dataStartRow;
    actualNumRows = endRow - CONFIG.dataStartRow + 1;

    if (actualNumRows <= 0) {
      SpreadsheetApp.getUi().alert(`データ行（${CONFIG.dataStartRow}行目以降）を選択してください。`);
      return;
    }
  }

  // G列（threadGroupId列）をクリア
  const threadIdRange = sheet.getRange(actualStartRow, CONFIG.columns.threadGroupId, actualNumRows, 1);
  threadIdRange.clearContent();

  SpreadsheetApp.getUi().alert(`${actualNumRows}行のスレッドIDをクリアしました。`);

  console.log(`スレッドIDクリア: 行${actualStartRow}〜${actualStartRow + actualNumRows - 1}`);
}

/**
 * 一意のスレッドグループIDを生成します
 * 形式: thread-YYYYMMDD-連番（001, 002, 003...）
 * 既存のスレッドIDを確認して、次の番号を自動採番します
 * @return {string} スレッドグループID
 */
function generateThreadGroupId_() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const dateStr = `${year}${month}${day}`;
  const prefix = `thread-${dateStr}-`;

  // シート全体から既存のスレッドIDを取得
  const sheet = getTargetSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.dataStartRow) {
    // データがない場合は001から開始
    return `${prefix}001`;
  }

  const rowCount = lastRow - CONFIG.dataStartRow + 1;
  const range = sheet.getRange(CONFIG.dataStartRow, CONFIG.columns.threadGroupId, rowCount, 1);
  const values = range.getValues();

  // 同じ日付のスレッドIDから最大番号を探す
  let maxNumber = 0;
  for (let i = 0; i < values.length; i++) {
    const threadId = String(values[i][0] || '').trim();
    if (threadId.startsWith(prefix)) {
      // "thread-20250130-005" から "005" 部分を抽出
      const numberPart = threadId.substring(prefix.length);
      const num = parseInt(numberPart, 10);
      if (!isNaN(num) && num > maxNumber) {
        maxNumber = num;
      }
    }
  }

  // 次の番号を生成（3桁でゼロパディング）
  const nextNumber = (maxNumber + 1).toString().padStart(3, '0');
  return `${prefix}${nextNumber}`;
}

/**
 * シートにB列（文字数列）を挿入する初回セットアップ関数
 * 既存データがある場合は、B列以降を1列右にシフトします。
 *
 * 【重要】この関数は一度だけ実行してください。
 * 実行後は、既存のB〜H列がC〜I列に移動します。
 */
function insertCharCountColumn() {
  const sheet = getTargetSheet_();

  // B列の後ろに新しい列を挿入（既存のB列がC列になる）
  sheet.insertColumnAfter(1);

  // B1にヘッダーを設定
  sheet.getRange(1, 2).setValue('文字数');

  SpreadsheetApp.getUi().alert(
    'B列（文字数列）を挿入しました。\n' +
    '既存のデータは1列右にシフトされています。\n\n' +
    '次に「文字数を更新」を実行してください。'
  );

  console.log('B列（文字数列）を挿入しました。');
}

/**
 * 全行の文字数を計算してB列に表示します
 * A列に投稿本文がある行のみ処理します
 */
function updateCharacterCounts() {
  const sheet = getTargetSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.dataStartRow) {
    SpreadsheetApp.getUi().alert('データが存在しません。');
    return;
  }

  const rowCount = lastRow - CONFIG.dataStartRow + 1;

  // A列（投稿本文）を取得
  const contentRange = sheet.getRange(CONFIG.dataStartRow, CONFIG.columns.content, rowCount, 1);
  const contentValues = contentRange.getValues();

  // B列（文字数）に設定する値を準備
  const charCountValues = [];

  for (let i = 0; i < contentValues.length; i++) {
    const content = String(contentValues[i][0] || '').trim();

    if (content) {
      const charCount = countTwitterCharacters_(content);
      charCountValues.push([charCount]);
    } else {
      charCountValues.push(['']);
    }
  }

  // B列に一括で設定
  const charCountRange = sheet.getRange(CONFIG.dataStartRow, CONFIG.columns.charCount, rowCount, 1);
  charCountRange.setValues(charCountValues);

  // 文字数列を中央揃えに設定
  charCountRange.setHorizontalAlignment('center');

  // 制限超過している行に色を付ける（オプション）
  if (CONFIG.characterLimit.enabled) {
    const maxLength = CONFIG.planLimits[CONFIG.characterLimit.plan];

    for (let i = 0; i < charCountValues.length; i++) {
      const charCount = charCountValues[i][0];
      const rowNumber = CONFIG.dataStartRow + i;
      const cell = sheet.getRange(rowNumber, CONFIG.columns.charCount);

      if (charCount && charCount > maxLength) {
        // 文字数超過の場合は赤背景
        cell.setBackground('#ffcccc');
        cell.setFontColor('#cc0000');
      } else if (charCount) {
        // 正常な場合は背景をクリア
        cell.setBackground(null);
        cell.setFontColor(null);
      }
    }
  }

  SpreadsheetApp.getUi().alert(`${rowCount}行の文字数を更新しました。`);
  console.log(`${rowCount}行の文字数を更新しました。`);
}

/**
 * セルが編集されたときに自動実行される関数
 * A列（投稿本文）が編集されたら、自動的にB列（文字数）を更新します
 */
function onEdit(e) {
  // イベントオブジェクトがない場合は処理しない
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();

  // 対象シートでない場合は処理しない
  if (sheet.getName() !== CONFIG.sheetName) {
    return;
  }

  const editedRow = e.range.getRow();
  const editedColumn = e.range.getColumn();

  // ヘッダー行は処理しない
  if (editedRow < CONFIG.dataStartRow) {
    return;
  }

  // A列（投稿本文）が編集された場合のみ処理
  if (editedColumn === CONFIG.columns.content) {
    updateSingleRowCharCount_(sheet, editedRow);
  }
}

/**
 * 指定された行の文字数を更新します
 * @param {Sheet} sheet - 対象シート
 * @param {number} rowNumber - 行番号
 */
function updateSingleRowCharCount_(sheet, rowNumber) {
  const contentCell = sheet.getRange(rowNumber, CONFIG.columns.content);
  const content = String(contentCell.getValue() || '').trim();

  const charCountCell = sheet.getRange(rowNumber, CONFIG.columns.charCount);

  if (!content) {
    // 本文が空の場合は文字数もクリア
    charCountCell.setValue('');
    charCountCell.setBackground(null);
    charCountCell.setFontColor(null);
    return;
  }

  // 文字数を計算
  const charCount = countTwitterCharacters_(content);
  charCountCell.setValue(charCount);
  charCountCell.setHorizontalAlignment('center');

  // 文字数超過チェック
  if (CONFIG.characterLimit.enabled) {
    const maxLength = CONFIG.planLimits[CONFIG.characterLimit.plan];

    if (charCount > maxLength) {
      // 文字数超過の場合は赤背景
      charCountCell.setBackground('#ffcccc');
      charCountCell.setFontColor('#cc0000');
    } else {
      // 正常な場合は背景をクリア
      charCountCell.setBackground(null);
      charCountCell.setFontColor(null);
    }
  }
}
