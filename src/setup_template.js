/**
 * テンプレートスプレッドシート自動セットアップスクリプト
 * 
 * 使い方:
 * 1. 新しいGoogle Sheetsを開く
 * 2. 拡張機能 > Apps Script
 * 3. このコードを貼り付け
 * 4. setupTemplate() を実行
 */

function setupTemplate() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. シート名を変更
    const sheet = ss.getActiveSheet();
    sheet.setName('sheet');

    // 2. ヘッダー行を設定
    const headers = [
        '投稿本文', '文字数', '画像', 'noteリンク', '日にち', '時', '分',
        '投稿URL', 'スレッドID', 'ステータス',
        'いいね数', 'リポスト数', '返信数', '引用数',
        'インプレッション数', 'ブックマーク数', '最終更新日時'
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setHorizontalAlignment('center');

    // 3. 列幅を調整
    sheet.setColumnWidth(1, 300);  // 投稿本文
    sheet.setColumnWidth(2, 60);   // 文字数
    sheet.setColumnWidth(3, 100);  // 画像
    sheet.setColumnWidth(4, 200);  // noteリンク
    sheet.setColumnWidth(5, 100);  // 日にち
    sheet.setColumnWidth(6, 50);   // 時
    sheet.setColumnWidth(7, 50);   // 分
    sheet.setColumnWidth(8, 250);  // 投稿URL
    sheet.setColumnWidth(9, 100);  // スレッドID
    sheet.setColumnWidth(10, 80);  // ステータス

    // 4. サンプルデータを追加
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);

    const sampleData = [
        [
            'おはようございます！今日も頑張りましょう🌟',
            '',  // 文字数は自動計算
            '',  // 画像
            '',  // noteリンク
            tomorrow,  // 日にち
            7,   // 時
            0,   // 分
            '',  // 投稿URL
            '',  // スレッドID
            'pending'  // ステータス
        ],
        [
            '新しい記事を書きました！\n\nぜひ読んでください📝',
            '',
            '',
            '',
            tomorrow,
            12,
            0,
            '',
            '',
            'pending'
        ]
    ];

    sheet.getRange(2, 1, sampleData.length, 10).setValues(sampleData);

    // 5. 日付の書式設定
    sheet.getRange(2, 5, sampleData.length, 1).setNumberFormat('yyyy/mm/dd');

    // 6. 「使い方」シートを作成
    const instructionSheet = ss.insertSheet('使い方');

    const instructions = [
        ['【X自動投稿システム - 使い方】'],
        [''],
        ['■ セットアップ手順'],
        [''],
        ['1. このスプレッドシートをコピー'],
        ['   ファイル > コピーを作成'],
        [''],
        ['2. X APIキーを取得'],
        ['   詳しくはこちら: https://note.com/konho/n/nf304497e6789'],
        [''],
        ['3. スクリプトプロパティを設定'],
        ['   拡張機能 > Apps Script > プロジェクトの設定 > スクリプトプロパティ'],
        [''],
        ['   以下の4つを追加：'],
        ['   - X_API_KEY'],
        ['   - X_API_SECRET'],
        ['   - X_ACCESS_TOKEN'],
        ['   - X_ACCESS_TOKEN_SECRET'],
        [''],
        ['4. トリガーを設定'],
        ['   Apps Script > トリガー > トリガーを追加'],
        [''],
        ['   関数: postNextScheduledToX'],
        ['   イベント: 時間主導型 > 分ベースのタイマー > 10分おき'],
        [''],
        ['■ 使い方'],
        [''],
        ['1. 「sheet」シートのA列に投稿内容を入力'],
        ['2. E〜G列に投稿日時を入力'],
        ['3. J列に「pending」と入力'],
        ['4. あとは自動で投稿されます！'],
        [''],
        ['■ 画像を投稿する場合'],
        [''],
        ['C列に以下のいずれかを設定：'],
        ['- セル内に画像を挿入（挿入 > 画像 > セル内に画像を挿入）'],
        ['- Google Drive URLを貼り付け'],
        [''],
        ['■ スレッド投稿（連続ツイート）'],
        [''],
        ['I列（スレッドID）に同じIDを入力すると連続投稿されます'],
        ['例: thread001'],
        [''],
        ['■ サポート'],
        [''],
        ['GitHub: https://github.com/euro0707/x-auto-post'],
        [''],
        ['⚠️ 注意: このスプレッドシートを他人と共有する場合は、'],
        ['Apps Scriptの編集権限は与えないでください。']
    ];

    instructionSheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
    instructionSheet.setColumnWidth(1, 600);

    // タイトル行を太字に
    instructionSheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);

    // 7. シートの順番を変更（使い方を先頭に）
    ss.setActiveSheet(instructionSheet);
    ss.moveActiveSheet(1);

    // 完了メッセージ
    SpreadsheetApp.getUi().alert(
        '✅ テンプレート作成完了！\n\n' +
        '次のステップ:\n' +
        '1. post_x.js のコードを Apps Script に追加\n' +
        '2. スクリプトプロパティに API キーを設定\n' +
        '3. トリガーを設定\n\n' +
        '詳しくは「使い方」シートをご覧ください。'
    );
}
