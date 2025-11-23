/**
 * スプレッドシート「sheet」の投稿待ち行から内容と投稿日時を読み取り、
 * 期限が来ている未投稿のデータを 1 件だけ X（旧Twitter）へ投稿します。
 *
 * シート構成（ヘッダーは任意）:
 *   A列: 投稿本文
 *   B列: 文字数（自動計算、Twitter仕様）
 *   C列: 画像（セル内に画像を貼り付け、またはGoogle Drive URLを入力、最大4枚）
 *   D列: noteリンク（任意、投稿後にリプライとして投稿される）
 *   E列: 日にち（1900/01/01 を 1 とするシリアル値）
 *   F列: 時（0〜23 の数値、空欄は 0 時扱い）
 *   G列: 分（0〜59 の数値、空欄は 0 分扱い）
 *   H列: 投稿URL（投稿完了時にツイートURLを格納）
 *   I列: スレッドグループID（同じIDを持つ行を連続投稿してスレッド化、空欄は単独投稿）
 *   J列: ステータス（pending/posting/posted/failed）
 *   K列: いいね数（エンゲージメント指標、自動更新）
 *   L列: リポスト数（エンゲージメント指標、自動更新）
 *   M列: 返信数（エンゲージメント指標、自動更新）
 *   N列: 引用数（エンゲージメント指標、自動更新）
 *   O列: インプレッション数（有料プラン専用、自動更新）
 *   P列: ブックマーク数（有料プラン専用、自動更新）
 *   Q列: 最終更新日時（エンゲージメント更新日時）
 *
 * 必須スクリプトプロパティ:
 *   X_API_KEY
 *   X_API_SECRET
 *   X_ACCESS_TOKEN
 *   X_ACCESS_TOKEN_SECRET
 *   （いずれも X API v2 ユーザーコンテキストで投稿権限を持つ値）
 *
 * オプションスクリプトプロパティ:
 *   xPlanType - Xプラン設定（'free'|'basic'|'premium'）
 *              メニューの「⚙️ Xプラン設定」から設定可能
 *
 * 想定運用フロー:
 * 1. 上記フォーマットでシートへ投稿候補を入力し、F・G 列は空欄のままにする。
 * 2. C列に画像を貼り付けるか、Google Drive URLを入力する（任意、最大4枚）
 * 3. D列にnoteのURLを入力すると、投稿後に自動的にリプライとして投稿される（任意）
 * 4. スレッド投稿したい場合は、I列に同じIDを入力する（例: thread001）
 * 5. 定期実行トリガーなどで postNextScheduledToX() を呼び出す。
 * 6. 実行ごとに期限を過ぎた未投稿行を先頭から探し、スレッドグループ全体を投稿・記録する。
 * 7. エンゲージメント指標は dailyEngagementUpdate() で毎日自動更新される。
 */
const CONFIG = Object.freeze({
  sheetName: 'sheet',
  dataStartRow: 2,
  columns: Object.freeze({
    content: 1,
    charCount: 2,         // 文字数
    image: 3,             // 画像
    noteUrl: 4,           // noteリンク
    day: 5,               // 日にち
    hour: 6,              // 時
    minute: 7,            // 分
    postedUrl: 8,         // 投稿URL
    threadGroupId: 9,     // スレッドグループID
    status: 10,           // ステータス
    // エンゲージメント指標（新規追加）
    likeCount: 11,        // いいね数
    retweetCount: 12,     // リポスト数
    replyCount: 13,       // 返信数
    quoteCount: 14,       // 引用ツイート数
    impressionCount: 15,  // インプレッション数（有料プランのみ）
    bookmarkCount: 16,    // ブックマーク数（有料プランのみ）
    lastUpdated: 17       // 最終更新日時
  }),
  api: Object.freeze({
    tweetEndpoint: 'https://api.twitter.com/2/tweets',
    mediaUploadEndpoint: 'https://upload.twitter.com/1.1/media/upload.json',
    tweetLookupEndpoint: 'https://api.twitter.com/2/tweets'  // エンゲージメント取得用
  }),
  serialBaseDate: new Date(1899, 11, 31),

  // Xプラン設定（文字数制限とエンゲージメント取得の統合設定）
  xPlan: Object.freeze({
    type: 'free',  // 'free' | 'basic' | 'premium'

    // プラン別の機能制限
    limits: Object.freeze({
      characterLimit: 140,              // 文字数制限
      canAccessPremiumMetrics: false    // インプレッション・ブックマーク取得可否
    })
  }),

  // プラン別設定のプリセット
  planPresets: Object.freeze({
    free: {
      characterLimit: 140,
      canAccessPremiumMetrics: false
    },
    basic: {
      characterLimit: 25000,
      canAccessPremiumMetrics: true
    },
    premium: {
      characterLimit: 25000,
      canAccessPremiumMetrics: true
    }
  }),

  // 文字数カウント設定
  characterLimit: Object.freeze({
    enabled: true,         // 文字数チェックを有効にする場合は true
    skipOnExceed: true,    // 超過時にスキップする場合は true、エラーにする場合は false
    urlLength: 23          // URLは自動短縮され、1リンクあたり約23文字としてカウント
  }),

  // エンゲージメント取得設定
  engagement: Object.freeze({
    enabled: true,              // エンゲージメント取得機能全体のON/OFF
    daysBack: 7,                // 更新対象期間（日数）
    batchSize: 100,             // 一度に処理する件数（API制限対策）
    retryOnRateLimit: true,     // Rate Limit時にリトライするか
    sleepBetweenRequests: 100   // リクエスト間のスリープ時間（ミリ秒）
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
 * Script Propertiesから保存されたXプラン設定を読み込み、適用します。
 * 保存された設定がない場合はCONFIGのデフォルト値を使用します。
 *
 * @return {Object} - { type: string, limits: { characterLimit: number, canAccessPremiumMetrics: boolean } }
 */
function getXPlanConfig() {
  const savedPlan = PropertiesService.getScriptProperties().getProperty('xPlanType');
  const planType = savedPlan || CONFIG.xPlan.type;

  // プリセットから設定を取得
  const preset = CONFIG.planPresets[planType] || CONFIG.planPresets['free'];

  return {
    type: planType,
    limits: {
      characterLimit: preset.characterLimit,
      canAccessPremiumMetrics: preset.canAccessPremiumMetrics
    }
  };
}

/**
 * X API v2を使って単一のツイートを投稿します。
 * @param {string} content - 投稿内容
 * @param {string|null} replyToTweetId - 返信先のツイートID（スレッド用、nullなら通常投稿）
 * @param {Object} credentials - API認証情報
 * @param {Array<string>} mediaIds - メディアIDの配列（最大4つ、オプション）
 * @return {Object} - {tweetId: string, tweetUrl: string}
 */
function postSingleTweet_(content, replyToTweetId, credentials, mediaIds = []) {
  const { apiKey, apiSecret, accessToken, accessTokenSecret } = credentials;

  const method = 'POST';
  const url = CONFIG.api.tweetEndpoint;

  // ペイロード構築（reply用のパラメータとメディアIDを含む）
  const payloadObj = { text: content };
  if (replyToTweetId) {
    payloadObj.reply = { in_reply_to_tweet_id: replyToTweetId };
  }
  if (mediaIds && mediaIds.length > 0) {
    payloadObj.media = { media_ids: mediaIds };
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
 * セルから画像を取得します（セル内の画像またはGoogle Drive URL）
 * @param {Sheet} sheet - シート
 * @param {number} rowNumber - 行番号
 * @param {number} columnNumber - 列番号
 * @return {Array<Blob>} - 画像のBlobの配列（最大4枚）
 */
function getImagesFromCell_(sheet, rowNumber, columnNumber) {
  const images = [];

  // 1. セル内に貼り付けられた画像を取得（OverGridImage）
  const allImages = sheet.getImages();
  for (let i = 0; i < allImages.length; i++) {
    const img = allImages[i];
    const anchorRow = img.getAnchorRow();
    const anchorCol = img.getAnchorColumn();

    // 指定した行・列にある画像を収集（複数対応）
    if (anchorRow === rowNumber && anchorCol === columnNumber) {
      try {
        // 元の画像形式を保持してBlobを取得
        const blob = img.getBlob();
        if (blob) {
          images.push(blob);
          console.log(`行 ${rowNumber}: セル内画像を取得しました（${blob.getBytes().length} bytes, ${blob.getContentType()}）`);
        }
      } catch (error) {
        console.warn(`行 ${rowNumber}: セル内画像の取得に失敗しました: ${error}`);
      }

      // 最大4枚まで
      if (images.length >= 4) {
        break;
      }
    }
  }

  // 2. セルの値がURLの場合（画像がまだない場合のみ）
  if (images.length === 0) {
    const cellValue = String(sheet.getRange(rowNumber, columnNumber).getValue() || '').trim();
    if (cellValue) {
      // カンマ区切りで複数URL対応
      const urls = cellValue.split(',').map(url => url.trim()).filter(url => url);

      for (let i = 0; i < urls.length && images.length < 4; i++) {
        const url = urls[i];

        // Google Drive URLからファイルIDを抽出
        const driveIdMatch = url.match(/[-\w]{25,}/);
        if (driveIdMatch) {
          try {
            const fileId = driveIdMatch[0];
            const file = DriveApp.getFileById(fileId);
            const blob = file.getBlob();
            images.push(blob);
            console.log(`行 ${rowNumber}: Google Drive画像を取得しました（${blob.getBytes().length} bytes）`);
          } catch (error) {
            console.warn(`行 ${rowNumber}: Google Drive画像の取得に失敗しました (${url}): ${error}`);
          }
        } else if (url.startsWith('http')) {
          // 通常のURL（外部画像）の場合
          try {
            const response = UrlFetchApp.fetch(url);
            const blob = response.getBlob();
            images.push(blob);
            console.log(`行 ${rowNumber}: URL画像を取得しました（${blob.getBytes().length} bytes）`);
          } catch (error) {
            console.warn(`行 ${rowNumber}: URL画像の取得に失敗しました (${url}): ${error}`);
          }
        }
      }
    }
  }

  // 最大4枚まで
  return images.slice(0, 4);
}

/**
 * X API v1.1を使って画像をアップロードします
 * @param {Blob} imageBlob - アップロードする画像のBlob
 * @param {Object} credentials - API認証情報
 * @return {string} - メディアID
 */
function uploadMediaToTwitter_(imageBlob, credentials) {
  const { apiKey, apiSecret, accessToken, accessTokenSecret } = credentials;

  const method = 'POST';
  const url = CONFIG.api.mediaUploadEndpoint;

  // OAuth 1.0a署名を生成
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

  // マルチパートフォームデータを構築
  const boundary = '----GASBoundary' + Utilities.getUuid();

  // ヘッダー部分
  const header = '--' + boundary + '\r\n' +
    'Content-Disposition: form-data; name="media"\r\n' +
    'Content-Type: ' + imageBlob.getContentType() + '\r\n\r\n';

  // フッター部分
  const footer = '\r\n--' + boundary + '--\r\n';

  // バイト配列を結合
  const headerBytes = Utilities.newBlob(header).getBytes();
  const imageBytes = imageBlob.getBytes();
  const footerBytes = Utilities.newBlob(footer).getBytes();

  // 3つのバイト配列を結合
  const payload = [];
  for (let i = 0; i < headerBytes.length; i++) {
    payload.push(headerBytes[i]);
  }
  for (let i = 0; i < imageBytes.length; i++) {
    payload.push(imageBytes[i]);
  }
  for (let i = 0; i < footerBytes.length; i++) {
    payload.push(footerBytes[i]);
  }

  const response = UrlFetchApp.fetch(url, {
    method,
    muteHttpExceptions: true,
    headers: {
      Authorization: buildOAuthHeader(),
      'Content-Type': 'multipart/form-data; boundary=' + boundary
    },
    payload: payload
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(`メディアアップロードエラー ${statusCode}: ${body}`);
  }

  let parsed;
  try {
    parsed = JSON.parse(body);
  } catch (error) {
    throw new Error(`メディアアップロードレスポンスの解析に失敗しました: ${error}`);
  }

  const mediaId = parsed && parsed.media_id_string;
  if (!mediaId) {
    throw new Error(`メディアアップロードレスポンスにmedia_idが含まれていません: ${body}`);
  }

  console.log(`画像をアップロードしました（メディアID: ${mediaId}）`);
  return mediaId;
}

/**
 * ツイートURLからツイートIDを抽出します
 * @param {string} tweetUrl - ツイートURL
 * @return {string|null} - ツイートID（抽出できない場合はnull）
 */
function extractTweetId_(tweetUrl) {
  if (!tweetUrl) {
    return null;
  }

  // 対応URL形式:
  // - https://x.com/i/web/status/1234567890
  // - https://x.com/username/status/1234567890
  // - https://twitter.com/username/status/1234567890
  const match = tweetUrl.match(/status\/(\d+)/);
  return match ? match[1] : null;
}

/**
 * X API v2でツイートのエンゲージメントを取得します
 * @param {string} tweetId - ツイートID
 * @param {Object} credentials - API認証情報
 * @return {Object} - エンゲージメント指標 {like_count, retweet_count, reply_count, quote_count, impression_count, bookmark_count}
 */
function getTweetEngagement_(tweetId, credentials) {
  const { apiKey, apiSecret, accessToken, accessTokenSecret } = credentials;

  const method = 'GET';
  const baseUrl = `${CONFIG.api.tweetLookupEndpoint}/${tweetId}`;
  const queryParams = { 'tweet.fields': 'public_metrics' };
  const url = `${baseUrl}?tweet.fields=public_metrics`;

  // OAuth 1.0a署名を生成
  const buildOAuthHeader = () => {
    const oauthParams = {
      oauth_consumer_key: apiKey,
      oauth_nonce: Utilities.getUuid().replace(/-/g, ''),
      oauth_signature_method: 'HMAC-SHA1',
      oauth_timestamp: Math.floor(Date.now() / 1000).toString(),
      oauth_token: accessToken,
      oauth_version: '1.0'
    };

    // OAuthパラメータとクエリパラメータを結合してソート
    const allParams = { ...oauthParams, ...queryParams };
    const sortedParams = Object.keys(allParams)
      .sort()
      .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(allParams[key])}`)
      .join('&');

    const signatureBaseString = [
      method.toUpperCase(),
      encodeURIComponent(baseUrl),  // ベースURLのみをエンコード
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
      Authorization: buildOAuthHeader()
    }
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(`エンゲージメント取得エラー ${statusCode}: ${body}`);
  }

  let parsed;
  try {
    parsed = JSON.parse(body);
  } catch (error) {
    throw new Error(`エンゲージメント取得レスポンスの解析に失敗しました: ${error}`);
  }

  const metrics = parsed && parsed.data && parsed.data.public_metrics;
  if (!metrics) {
    throw new Error(`エンゲージメント取得レスポンスにpublic_metricsが含まれていません: ${body}`);
  }

  // プラン設定に応じて有料指標を含めるか決定
  const xPlan = getXPlanConfig();
  const includePremiumMetrics = xPlan.limits.canAccessPremiumMetrics;

  return {
    like_count: metrics.like_count || 0,
    retweet_count: metrics.retweet_count || 0,
    reply_count: metrics.reply_count || 0,
    quote_count: metrics.quote_count || 0,
    impression_count: includePremiumMetrics ? (metrics.impression_count || 0) : null,
    bookmark_count: includePremiumMetrics ? (metrics.bookmark_count || 0) : null
  };
}

/**
 * エンゲージメント更新処理の共通ロジック
 * @param {Sheet} sheet - スプレッドシート
 * @param {Array} targetRows - 対象ツイートの配列
 * @param {Object} credentials - API認証情報
 * @return {Object} - { successCount: number, failCount: number }
 * @private
 */
function processEngagementUpdates_(sheet, targetRows, credentials) {
  let successCount = 0;
  let failCount = 0;
  const batchSize = CONFIG.engagement.batchSize;
  const sleepMs = CONFIG.engagement.sleepBetweenRequests;

  for (let i = 0; i < targetRows.length; i++) {
    const target = targetRows[i];

    try {
      const engagement = getTweetEngagement_(target.tweetId, credentials);

      // スプレッドシートに書き込み
      const row = target.rowNumber;
      sheet.getRange(row, CONFIG.columns.likeCount).setValue(engagement.like_count);
      sheet.getRange(row, CONFIG.columns.retweetCount).setValue(engagement.retweet_count);
      sheet.getRange(row, CONFIG.columns.replyCount).setValue(engagement.reply_count);
      sheet.getRange(row, CONFIG.columns.quoteCount).setValue(engagement.quote_count);

      if (engagement.impression_count !== null) {
        sheet.getRange(row, CONFIG.columns.impressionCount).setValue(engagement.impression_count);
      }
      if (engagement.bookmark_count !== null) {
        sheet.getRange(row, CONFIG.columns.bookmarkCount).setValue(engagement.bookmark_count);
      }

      sheet.getRange(row, CONFIG.columns.lastUpdated).setValue(new Date());
      successCount++;

      if (sleepMs > 0 && i < targetRows.length - 1) {
        Utilities.sleep(sleepMs);
      }

      if ((i + 1) % batchSize === 0) {
        console.log(`進捗: ${i + 1}/${targetRows.length} 件完了`);
      }

    } catch (error) {
      console.error(`エラー（行${target.rowNumber}, URL: ${target.postedUrl}）: ${error.message}`);
      failCount++;

      if (CONFIG.engagement.retryOnRateLimit && error.message.includes('429')) {
        console.log('レート制限に達しました。60秒待機してリトライします...');
        Utilities.sleep(60000);
        i--;
        continue;
      }
    }
  }

  return { successCount, failCount };
}

/**
 * 投稿済みツイートのエンゲージメント指標を一括更新します。
 *
 * @param {number} daysBack - 何日前までのツイートを対象とするか（デフォルト: CONFIG.engagement.daysBack）
 */
function updateAllEngagementMetrics(daysBack) {
  if (!CONFIG.engagement.enabled) {
    SpreadsheetApp.getUi().alert('エンゲージメント取得機能が無効化されています。CONFIG.engagement.enabledをtrueに設定してください。');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const credentials = getCredentials_();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('データが見つかりません。');
    return;
  }

  // 対象期間の計算
  const targetDays = daysBack || CONFIG.engagement.daysBack;
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - targetDays);

  // 投稿済みツイートを収集
  const statusCol = CONFIG.columns.status;
  const postedUrlCol = CONFIG.columns.postedUrl;
  const dayCol = CONFIG.columns.day;

  const statuses = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
  const postedUrls = sheet.getRange(2, postedUrlCol, lastRow - 1, 1).getValues();
  const days = sheet.getRange(2, dayCol, lastRow - 1, 1).getValues();

  const targetRows = [];
  for (let i = 0; i < statuses.length; i++) {
    const status = String(statuses[i][0]).trim().toLowerCase();
    const postedUrl = String(postedUrls[i][0]).trim();
    const day = days[i][0];

    // ステータスがpostedで、投稿URLがあり、対象期間内のツイートを選択
    if (status === 'posted' && postedUrl) {
      // 日付フィルタ（日付がない場合は対象に含める）
      if (!day || new Date(day) >= cutoffDate) {
        const tweetId = extractTweetId_(postedUrl);
        if (tweetId) {
          targetRows.push({
            rowNumber: i + 2,  // 実際の行番号（ヘッダー分+1、配列インデックス分+1）
            tweetId: tweetId,
            postedUrl: postedUrl
          });
        }
      }
    }
  }

  if (targetRows.length === 0) {
    SpreadsheetApp.getUi().alert(`過去${targetDays}日間に投稿されたツイートが見つかりませんでした。`);
    return;
  }

  SpreadsheetApp.getUi().alert(`${targetRows.length}件のツイートのエンゲージメントを取得します。\nしばらくお待ちください...`);

  // 共通関数を使用してエンゲージメント更新
  const { successCount, failCount } = processEngagementUpdates_(sheet, targetRows, credentials);

  // 結果報告
  const message = `エンゲージメント更新が完了しました。\n\n` +
    `成功: ${successCount}件\n` +
    `失敗: ${failCount}件\n\n` +
    (failCount > 0 ? `詳細はログを確認してください（表示 > ログ）` : '');

  SpreadsheetApp.getUi().alert(message);
}

/**
 * 【トリガー設定用】毎日朝9時に実行する関数
 * 過去30日間の投稿済みツイートのエンゲージメントを自動更新します。
 *
 * 設定方法:
 * 1. Apps Script エディタで「トリガー」を開く
 * 2. 「トリガーを追加」をクリック
 * 3. 実行する関数: dailyEngagementUpdate
 * 4. イベントのソース: 時間主導型
 * 5. 時間ベースのトリガー: 日タイマー
 * 6. 時刻: 午前9時〜10時
 * 7. 保存
 */
function dailyEngagementUpdate() {
  console.log('=== 毎日のエンゲージメント自動更新を開始 ===');

  try {
    const daysBack = 30;  // 過去30日分

    // スプレッドシートとシートを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
    if (!sheet) {
      console.error(`シート「${CONFIG.sheetName}」が見つかりません。`);
      return;
    }

    const credentials = getCredentials_();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log('データが見つかりません。処理を終了します。');
      return;
    }

    // 対象期間の計算
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysBack);

    // 投稿済みツイートを収集
    const statusCol = CONFIG.columns.status;
    const postedUrlCol = CONFIG.columns.postedUrl;
    const dayCol = CONFIG.columns.day;

    const statuses = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
    const postedUrls = sheet.getRange(2, postedUrlCol, lastRow - 1, 1).getValues();
    const days = sheet.getRange(2, dayCol, lastRow - 1, 1).getValues();

    const targetRows = [];
    for (let i = 0; i < statuses.length; i++) {
      const status = String(statuses[i][0]).trim().toLowerCase();
      const postedUrl = String(postedUrls[i][0]).trim();
      const day = days[i][0];

      if (status === 'posted' && postedUrl) {
        if (!day || new Date(day) >= cutoffDate) {
          const tweetId = extractTweetId_(postedUrl);
          if (tweetId) {
            targetRows.push({
              rowNumber: i + 2,
              tweetId: tweetId,
              postedUrl: postedUrl
            });
          }
        }
      }
    }

    if (targetRows.length === 0) {
      console.log(`過去${daysBack}日間に投稿されたツイートが見つかりませんでした。`);
      return;
    }

    console.log(`${targetRows.length}件のツイートのエンゲージメントを取得します。`);

    // 共通関数を使用してエンゲージメント更新
    const { successCount, failCount } = processEngagementUpdates_(sheet, targetRows, credentials);

    console.log(`=== エンゲージメント更新完了 ===`);
    console.log(`成功: ${successCount}件`);
    console.log(`失敗: ${failCount}件`);

  } catch (error) {
    console.error(`dailyEngagementUpdate実行エラー: ${error.message}`);
    console.error(error.stack);
  }
}

/**
 * 期限を過ぎた未投稿データを X へ投稿し、H 列（投稿URL）と J 列（ステータス）を更新します。
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
    const xPlan = getXPlanConfig();
    const maxLength = xPlan.limits.characterLimit;
    if (!maxLength) {
      throw new Error(`無効なプラン設定: ${xPlan.type}`);
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
    console.log(`行 ${rowNumber} の文字数チェック OK（${charCount}/${maxLength}文字、プラン: ${xPlan.type}）`);
  }

  try {
    // 画像を取得
    const imageBlobs = getImagesFromCell_(sheet, rowNumber, CONFIG.columns.image);
    const mediaIds = [];

    // 画像がある場合はアップロード
    if (imageBlobs.length > 0) {
      console.log(`行 ${rowNumber}: ${imageBlobs.length}枚の画像をアップロード中...`);
      for (let i = 0; i < imageBlobs.length; i++) {
        try {
          const mediaId = uploadMediaToTwitter_(imageBlobs[i], credentials);
          mediaIds.push(mediaId);
          console.log(`行 ${rowNumber}: 画像 ${i + 1}/${imageBlobs.length} をアップロードしました`);
        } catch (uploadError) {
          console.warn(`行 ${rowNumber}: 画像 ${i + 1} のアップロードに失敗しました: ${uploadError}`);
          // 画像アップロード失敗時はスキップして続行
        }
      }
    }

    // メイン投稿（画像付き）
    const result = postSingleTweet_(content, null, credentials, mediaIds);
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
      const xPlan = getXPlanConfig();
      const maxLength = xPlan.limits.characterLimit;
      if (!maxLength) {
        throw new Error(`無効なプラン設定: ${xPlan.type}`);
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

      // 画像を取得
      const imageBlobs = getImagesFromCell_(sheet, rowNumber, CONFIG.columns.image);
      const mediaIds = [];

      // 画像がある場合はアップロード
      if (imageBlobs.length > 0) {
        console.log(`行 ${rowNumber}: ${imageBlobs.length}枚の画像をアップロード中...`);
        for (let j = 0; j < imageBlobs.length; j++) {
          try {
            const mediaId = uploadMediaToTwitter_(imageBlobs[j], credentials);
            mediaIds.push(mediaId);
            console.log(`行 ${rowNumber}: 画像 ${j + 1}/${imageBlobs.length} をアップロードしました`);
          } catch (uploadError) {
            console.warn(`行 ${rowNumber}: 画像 ${j + 1} のアップロードに失敗しました: ${uploadError}`);
            // 画像アップロード失敗時はスキップして続行
          }
        }
      }

      const result = postSingleTweet_(content, previousTweetId, credentials, mediaIds);

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

  const xPlan = getXPlanConfig();
  const maxLength = xPlan.limits.characterLimit;

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
  console.log(`プラン: ${xPlan.type}`);
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
    plan: xPlan.type,
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
    .addItem('エンゲージメントを更新', 'updateAllEngagementMetrics')
    .addItem('人気ツイートTOP10を表示', 'analyzeTopTweets')
    .addSeparator()
    .addItem('⚙️ Xプラン設定', 'showXPlanSettingsDialog')
    .addSeparator()
    .addItem('【初回のみ】文字数列を挿入', 'insertCharCountColumn')
    .addItem('【初回のみ】画像列を挿入', 'insertImageColumn')
    .addItem('【初回のみ】エンゲージメント列を挿入', 'insertEngagementColumns')
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
 * シートにC列（画像列）を挿入する初回セットアップ関数
 * 既存データがある場合は、C列以降を1列右にシフトします。
 *
 * 【重要】この関数は一度だけ実行してください。
 * 実行後は、既存のC〜I列がD〜J列に移動します。
 */
function insertImageColumn() {
  const sheet = getTargetSheet_();

  // C列の後ろに新しい列を挿入（既存のC列がD列になる）
  sheet.insertColumnAfter(2);

  // C1にヘッダーを設定
  sheet.getRange(1, 3).setValue('画像');

  SpreadsheetApp.getUi().alert(
    'C列（画像列）を挿入しました。\n' +
    '既存のデータは1列右にシフトされています。\n\n' +
    'C列に画像を貼り付けるか、Google Drive URLを入力してください。'
  );

  console.log('C列（画像列）を挿入しました。');
}

/**
 * エンゲージメント列（K〜Q列）をスプレッドシートに挿入します。
 * 初回のみ実行する必要があります。
 */
function insertEngagementColumns() {
  const sheet = getTargetSheet_();

  // 既にエンゲージメント列が存在するかチェック
  const existingHeader = sheet.getRange(1, CONFIG.columns.likeCount).getValue();
  if (existingHeader) {
    const result = SpreadsheetApp.getUi().alert(
      '確認',
      'K列以降に既にデータが存在します。\n上書きしますか？',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (result !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }

  // K列以降に既に列がある場合は、そのまま上書き
  // ない場合は、必要に応じて列を追加
  const lastColumn = sheet.getLastColumn();
  const needColumns = CONFIG.columns.lastUpdated; // Q列 = 17列目まで必要

  if (lastColumn < needColumns) {
    // 不足している列を追加
    const columnsToAdd = needColumns - lastColumn;
    sheet.insertColumnsAfter(lastColumn, columnsToAdd);
  }

  // ヘッダーを設定
  const headers = [
    ['いいね数', 'リポスト数', '返信数', '引用数', 'インプレッション数', 'ブックマーク数', '最終更新']
  ];
  sheet.getRange(1, CONFIG.columns.likeCount, 1, 7).setValues(headers);

  // 有料プラン専用列に色付け
  const premiumColumns = [CONFIG.columns.impressionCount, CONFIG.columns.bookmarkCount];
  for (const col of premiumColumns) {
    const cell = sheet.getRange(1, col);
    cell.setBackground('#FFF4E6');
    cell.setNote('この列は有料プラン（Basic/Premium）でのみデータが取得されます');
  }

  // ヘッダー行を太字に
  sheet.getRange(1, CONFIG.columns.likeCount, 1, 7).setFontWeight('bold');

  SpreadsheetApp.getUi().alert(
    'エンゲージメント列（K〜Q列）を挿入しました。\n\n' +
    '・K列: いいね数\n' +
    '・L列: リポスト数\n' +
    '・M列: 返信数\n' +
    '・N列: 引用数\n' +
    '・O列: インプレッション数（有料プラン専用）\n' +
    '・P列: ブックマーク数（有料プラン専用）\n' +
    '・Q列: 最終更新\n\n' +
    'メニューの「エンゲージメントを更新」から、投稿済みツイートの指標を取得できます。'
  );

  console.log('エンゲージメント列（K〜Q列）を挿入しました。');
}

/**
 * 人気ツイートTOP10を分析してダイアログ表示します。
 * エンゲージメント指標の合計スコアでランキングします。
 */
function analyzeTopTweets() {
  const sheet = getTargetSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('データが見つかりません。');
    return;
  }

  // データを取得
  const contents = sheet.getRange(2, CONFIG.columns.content, lastRow - 1, 1).getValues();
  const postedUrls = sheet.getRange(2, CONFIG.columns.postedUrl, lastRow - 1, 1).getValues();
  const likes = sheet.getRange(2, CONFIG.columns.likeCount, lastRow - 1, 1).getValues();
  const retweets = sheet.getRange(2, CONFIG.columns.retweetCount, lastRow - 1, 1).getValues();
  const replies = sheet.getRange(2, CONFIG.columns.replyCount, lastRow - 1, 1).getValues();
  const quotes = sheet.getRange(2, CONFIG.columns.quoteCount, lastRow - 1, 1).getValues();

  // エンゲージメントデータがあるツイートを集計
  const tweets = [];
  for (let i = 0; i < contents.length; i++) {
    const likeCount = Number(likes[i][0]) || 0;
    const retweetCount = Number(retweets[i][0]) || 0;
    const replyCount = Number(replies[i][0]) || 0;
    const quoteCount = Number(quotes[i][0]) || 0;

    // エンゲージメントデータがある場合のみ集計
    if (likeCount > 0 || retweetCount > 0 || replyCount > 0 || quoteCount > 0) {
      const totalEngagement = likeCount + retweetCount + replyCount + quoteCount;
      tweets.push({
        rowNumber: i + 2,
        content: String(contents[i][0]).substring(0, 50) + '...',  // 最初の50文字
        postedUrl: String(postedUrls[i][0]),
        likeCount,
        retweetCount,
        replyCount,
        quoteCount,
        totalEngagement
      });
    }
  }

  if (tweets.length === 0) {
    SpreadsheetApp.getUi().alert(
      'エンゲージメントデータが見つかりませんでした。\n\n' +
      'メニューの「エンゲージメントを更新」を実行してから、再度お試しください。'
    );
    return;
  }

  // 総エンゲージメント数でソート（降順）
  tweets.sort((a, b) => b.totalEngagement - a.totalEngagement);

  // TOP10を抽出
  const top10 = tweets.slice(0, 10);

  // 結果表示用のHTML作成
  let message = '【人気ツイートTOP10】\n\n';
  top10.forEach((tweet, index) => {
    message += `${index + 1}位: ${tweet.totalEngagement}エンゲージメント\n`;
    message += `   いいね:${tweet.likeCount} リポスト:${tweet.retweetCount} ` +
      `返信:${tweet.replyCount} 引用:${tweet.quoteCount}\n`;
    message += `   内容: ${tweet.content}\n`;
    message += `   URL: ${tweet.postedUrl}\n`;
    message += `   行: ${tweet.rowNumber}\n\n`;
  });

  // 統計情報
  const totalTweets = tweets.length;
  const avgEngagement = tweets.reduce((sum, t) => sum + t.totalEngagement, 0) / totalTweets;
  message += `─────────────────\n`;
  message += `分析対象ツイート数: ${totalTweets}件\n`;
  message += `平均エンゲージメント: ${Math.round(avgEngagement)}`;

  SpreadsheetApp.getUi().alert(message);
  console.log('人気ツイートTOP10分析を実行しました。');
}

/**
 * Xプラン設定ダイアログを表示します。
 * ユーザーが選択したプラン情報をScript Propertiesに保存します。
 */
function showXPlanSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentPlan = PropertiesService.getScriptProperties().getProperty('xPlanType') || 'free';

  const response = ui.alert(
    'Xプラン設定',
    `現在のプラン: ${currentPlan}\n\n` +
    `Xプランを選択してください:\n\n` +
    `・Free（無料）: 文字数制限140字、基本指標のみ\n` +
    `・Basic（有料）: 文字数制限25000字、全指標取得可能\n` +
    `・Premium（有料）: 文字数制限25000字、全指標取得可能\n\n` +
    `変更する場合は「はい」を押してください。`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // プラン選択
  const planResponse = ui.prompt(
    'Xプラン選択',
    '以下のいずれかを入力してください:\n\n' +
    '・free（無料プラン）\n' +
    '・basic（有料プラン）\n' +
    '・premium（有料プラン）',
    ui.ButtonSet.OK_CANCEL
  );

  if (planResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const selectedPlan = planResponse.getResponseText().trim().toLowerCase();

  // バリデーション
  if (!['free', 'basic', 'premium'].includes(selectedPlan)) {
    ui.alert('エラー', '無効なプランです。free、basic、premiumのいずれかを入力してください。', ui.ButtonSet.OK);
    return;
  }

  // Script Propertiesに保存
  PropertiesService.getScriptProperties().setProperty('xPlanType', selectedPlan);

  // 設定内容を表示
  const preset = CONFIG.planPresets[selectedPlan];
  ui.alert(
    '設定完了',
    `Xプランを「${selectedPlan}」に変更しました。\n\n` +
    `【適用される設定】\n` +
    `・文字数制限: ${preset.characterLimit}字\n` +
    `・プレミアム指標: ${preset.canAccessPremiumMetrics ? '有効' : '無効'}\n\n` +
    `次回の投稿・エンゲージメント取得から反映されます。`,
    ui.ButtonSet.OK
  );

  console.log(`Xプランを「${selectedPlan}」に変更しました。`);
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
    const xPlan = getXPlanConfig();
    const maxLength = xPlan.limits.characterLimit;

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
    const xPlan = getXPlanConfig();
    const maxLength = xPlan.limits.characterLimit;

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
