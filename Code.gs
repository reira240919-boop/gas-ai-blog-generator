/**
 * GAS + AI ブログ記事自動生成（土台）
 * 正: requirements.md
 *
 * 今回の範囲:
 *   - 1行目の見出しを用意する
 *   - メニュー「記事生成」を出す
 *   - 選択した1行のキーワードを読んで表示する（ステータスは変えない）
 *   - 記事を1本書く関数（generateArticle）
 *   - 選択した1行の記事を作り、シートに書く（generateSelectedRow）
 *
 * まだやらないこと: WordPress 送信 / G列の URL 記入 / 一括生成
 * 鍵（API キー等）はこのファイルにもシートにも書かない。
 * API キーはスクリプトプロパティ ANTHROPIC_API_KEY に置く。
 */

// Claude API の設定（ここだけで変える）
// 途中で切れないよう余裕を持たせつつ、1回の生成が 6 分（GAS の上限）を超えない大きさにする
const CLAUDE_MODEL = 'claude-sonnet-5-5';
const CLAUDE_MAX_TOKENS = 12000;

// 最新の法令メモ。AI が古い言い方で書かないよう、プロンプトに入れる
// 法律が変わったら、ここに1行足す（それでも公開前に人が確認する）
const LAW_NOTES = [
  '刑罰の「懲役」は 2025年6月から「拘禁刑」に変わった。罰則は「6か月以下の拘禁刑または30万円以下の罰金」のように書く',
];

// 1行目の見出し（A〜K の順）
const HEADERS = [
  'キーワード',      // A
  'サブキーワード',  // B
  '検索意図',        // C（空でもよい）
  'ステータス',      // D
  'タイトル',        // E
  '文字数',          // F
  'WordPress URL',   // G
  '公開前チェック',  // H
  'メモ',            // I
  '本文',            // J
  'メタディスクリプション', // K
];

// 列番号（1始まり）
const COL = {
  KEYWORD: 1,
  SUB_KEYWORD: 2,
  INTENT: 3,
  STATUS: 4,
  TITLE: 5,
  CHAR_COUNT: 6,
  WP_URL: 7,
  CHECK: 8,
  MEMO: 9,
  BODY: 10,
  META: 11,
};

// シートに保存したが WordPress にはまだ送っていない行のメモ
const MEMO_SAVED_TO_SHEET = 'シートに保存済み。WordPress未送信';
// WordPress に下書き保存した行のメモ
const MEMO_SAVED_TO_WP = 'WordPressに下書き保存';
// H列が空のときだけ入れる確認のお願い
const CHECK_REMINDER = '公開前に、類似・法令の数字・料金を確認';

// 状態はこの5つだけ
const STATUS = {
  NOT_GENERATED: '未生成',
  GENERATING: '生成中',
  SHEET_SAVED: 'シート保存済み',   // シートに保存。WordPress は未送信
  DRAFT_SAVED: '下書き保存済み',   // WordPress に下書きで保存
  FAILED: '失敗',
};

/**
 * スプレッドシートを開いたときにメニューを出す
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('記事生成')
    .addItem('見出しを用意する', 'setupHeaders')
    .addItem('選択した行を確認', 'checkSelectedRow')
    .addItem('選択した行の記事を作る', 'generateSelectedRow')
    .addItem('選択した行を下書き保存', 'saveSelectedRowDraft')
    .addToUi();
}

/**
 * 1行目に見出しを置く。すでに文字が入っているセルは動かさない
 */
function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const current = range.getValues()[0];

  let added = 0;
  const next = current.map(function (value, i) {
    if (String(value).trim() !== '') return value; // 入っている中身はそのまま
    added++;
    return HEADERS[i];
  });

  range.setValues([next]);
  sheet.setFrozenRows(1);

  const ui = SpreadsheetApp.getUi();
  if (added === 0) {
    ui.alert('見出しはすでに入っています。変更はありません。');
  } else {
    ui.alert('見出しを用意しました（' + added + ' か所に入れました）。');
  }
}

/**
 * 今選んでいる1行だけを読み、キーワードを表示する。ステータスは変えない
 */
function checkSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection) {
    ui.alert('キーワードの行を選んでください。');
    return;
  }

  // 一括では動かさない
  if (selection.getNumRows() > 1) {
    ui.alert('1行だけ選んでください。');
    return;
  }

  const row = selection.getRow();
  if (row === 1) {
    ui.alert('キーワードの行を選んでください。');
    return;
  }

  const values = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const keyword = String(values[COL.KEYWORD - 1]).trim();
  const status = String(values[COL.STATUS - 1]).trim();

  if (keyword === '') {
    ui.alert(row + ' 行目のキーワードが空です。キーワードを入れてから選んでください。');
    return;
  }

  let message = row + ' 行目のキーワード: ' + keyword;
  if (isDraftSaved(status)) {
    message += '\n\nこの行はすでに「' + STATUS.DRAFT_SAVED + '」です。上書きはしません。';
  }
  ui.alert(message);
}

/**
 * すでに「下書き保存済み」か（上書き防止の判定。今回は判定のみ）
 */
function isDraftSaved(status) {
  return status === STATUS.DRAFT_SAVED;
}

/**
 * 今選んでいる1行の記事を作り、シートに書く。WordPress にはまだ送らない
 */
function generateSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection || selection.getRow() === 1) {
    ui.alert('キーワードの行を選んでください。');
    return;
  }
  // 一括では動かさない
  if (selection.getNumRows() > 1) {
    ui.alert('1行だけ選んでください。');
    return;
  }

  const row = selection.getRow();
  const values = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const keyword = String(values[COL.KEYWORD - 1]).trim();
  const subKeyword = String(values[COL.SUB_KEYWORD - 1]).trim();
  const intent = String(values[COL.INTENT - 1]).trim();
  const status = String(values[COL.STATUS - 1]).trim();
  const body = String(values[COL.BODY - 1]).trim();
  const check = String(values[COL.CHECK - 1]).trim();

  if (keyword === '') {
    ui.alert(row + ' 行目のキーワードが空です。キーワードを入れてから選んでください。');
    return;
  }
  if (isDraftSaved(status)) {
    ui.alert(row + ' 行目はすでに「' + STATUS.DRAFT_SAVED + '」です。上書きはしません。');
    return;
  }
  if (body !== '') {
    ui.alert(row + ' 行目はすでに本文があります。上書きはしません。');
    return;
  }
  // ステータスが空の行は「未生成」として扱う
  if (status !== '' && status !== STATUS.NOT_GENERATED && status !== STATUS.FAILED) {
    ui.alert(row + ' 行目のステータスが「' + status + '」です。' +
      '「' + STATUS.NOT_GENERATED + '」か「' + STATUS.FAILED + '」の行だけ作れます。');
    return;
  }

  // API を呼ぶ前に「生成中」をシートに書く
  sheet.getRange(row, COL.STATUS).setValue(STATUS.GENERATING);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast('「' + keyword + '」の記事を作っています（1〜2分）', '記事生成', -1);

  let article;
  try {
    article = generateArticle(keyword, subKeyword, intent);
  } catch (e) {
    sheet.getRange(row, COL.STATUS).setValue(STATUS.FAILED);
    sheet.getRange(row, COL.MEMO).setValue(e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('', '記事生成', 1);
    ui.alert(row + ' 行目の記事を作れませんでした。\n\n理由: ' + e.message);
    return;
  }

  sheet.getRange(row, COL.TITLE).setValue(article.title);
  sheet.getRange(row, COL.CHAR_COUNT).setValue(article.charCount);
  sheet.getRange(row, COL.BODY).setValue(article.body);
  sheet.getRange(row, COL.META).setValue(article.metaDescription);
  sheet.getRange(row, COL.MEMO).setValue(MEMO_SAVED_TO_SHEET);
  if (check === '') {
    sheet.getRange(row, COL.CHECK).setValue(CHECK_REMINDER);
  }
  sheet.getRange(row, COL.STATUS).setValue(STATUS.SHEET_SAVED);

  SpreadsheetApp.getActiveSpreadsheet().toast('', '記事生成', 1);
  ui.alert(row + ' 行目の記事をシートに保存しました（' + article.charCount + ' 字）。\n' +
    'WordPress にはまだ送っていません。');
}

// ---------------------------------------------------------------
// 記事生成（Claude API）
// ---------------------------------------------------------------

const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const META_DESCRIPTION_MAX = 120;

// 文体の見本（case2-sample-article.md の本文）
const SAMPLE_ARTICLE = `# 「勤怠管理はエクセルで十分」は本当か？担当者が見落としがちな3つの落とし穴

「うちはまだ社員数もそこまで多くないし、勤怠管理はエクセルで回せている」——そう考えている労務・総務担当の方は少なくありません。確かにエクセルは手軽で、追加コストもかからない優秀なツールです。ただ、会社の成長や働き方の変化に合わせて、いつのまにか“見えないコスト”が膨らんでいることも事実です。この記事では、エクセル勤怠を続けるうえで担当者が見落としがちな3つの落とし穴を、現場目線で整理します。

## 落とし穴①：集計ミスは「気づけないまま」積み上がる

エクセル勤怠で最も多いトラブルが、関数のズレやコピペミスによる集計誤りです。1人分の残業時間が数分ずれるだけなら気づきにくいですが、それが数十人・十数か月分積み重なると、未払い残業として後から大きな金額で表面化することがあります。しかもエクセルは「誰がいつ何を変更したか」の履歴が残らないため、原因の特定にも時間がかかります。月末のたびに担当者が電卓で検算している——そんな状態は、すでにエクセルの限界サインかもしれません。

## 落とし穴②：法改正への対応が、どうしても後手に回る

働き方関連の法改正は、ここ数年だけでも「有給休暇の年5日取得義務」「36協定の残業上限規制」「割増賃金率の引き上げ」など立て続けに行われています。エクセル運用の場合、こうした改正のたびに数式やシートを自前で組み直す必要があり、対応が遅れるほど労務リスクが高まります。特に有給5日の取得状況や、残業の上限超過は「うっかり見落とした」では済まされない領域です。

## 落とし穴③：テレワークや多様な働き方に追いつけない

オフィス出社が前提だった頃は、タイムカードやエクセルでも十分でした。しかしテレワーク・時差出勤・中抜けといった柔軟な働き方が広がると、「いつ・どこで・どれだけ働いたか」を正確に記録すること自体が難しくなります。自己申告ベースのエクセルでは、打刻漏れや申告忘れが増え、結果として正確な労働時間の把握ができなくなる——これは、従業員にとっても会社にとっても不幸な状態です。

## まとめ：エクセルが悪いのではなく、「フェーズ」の問題

ここまで3つの落とし穴を挙げましたが、エクセルそのものが悪いわけではありません。大切なのは、自社の人数・働き方が「エクセルで回せるフェーズ」を超えていないかを定期的に見直すことです。まずは、月末の勤怠集計に毎月どれくらいの工数がかかっているかを可視化してみてください。その時間が「本来やるべき仕事」を圧迫しているなら、それがシステム化を検討するベストなタイミングです。`;

/**
 * 記事を1本書く。シートには書き込まない
 * @param {string} keyword キーワード
 * @param {string} subKeyword サブキーワード（カンマ区切り）
 * @param {string} intent 検索意図（空でもよい）
 * @return {{title: string, body: string, metaDescription: string, charCount: number}}
 */
function generateArticle(keyword, subKeyword, intent) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    throw new Error(
      'API キーが未設定です。Apps Script の「プロジェクトの設定」→「スクリプト プロパティ」に ' +
      'ANTHROPIC_API_KEY を追加してください。'
    );
  }
  if (!keyword || String(keyword).trim() === '') {
    throw new Error('キーワードが空です。');
  }

  const response = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: CLAUDE_MAX_TOKENS,
      system: buildSystemPrompt_(),
      messages: [{ role: 'user', content: buildUserPrompt_(keyword, subKeyword, intent) }],
    }),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  const data = JSON.parse(response.getContentText());
  if (code !== 200) {
    const reason = data.error && data.error.message ? data.error.message : 'HTTP ' + code;
    throw new Error('Claude API エラー: ' + reason);
  }
  if (data.stop_reason === 'max_tokens') {
    throw new Error('文章が長すぎて途中で切れました。もう一度やり直してください。');
  }

  const text = data.content
    .filter(function (block) { return block.type === 'text'; })
    .map(function (block) { return block.text; })
    .join('');
  const article = parseArticleJson_(text);
  // タグの前後に入った改行・空白を取る（WordPress で余分な空きが出ないように）
  article.body = article.body.replace(/>\s+/g, '>').replace(/\s+</g, '<').trim();

  return {
    title: article.title,
    body: article.body,
    metaDescription: article.metaDescription,
    charCount: countBodyChars_(article.body),
  };
}

function buildSystemPrompt_() {
  return [
    'あなたは BtoB SaaS のマーケティングライターです。',
    '中小企業の経営者・IT 担当者に向けて、自社オウンドメディアの SEO 記事を書きます。',
    '',
    '次の見本記事のトーンで書いてください。カジュアルだが専門的で、読者の悩みに寄り添い、現場目線で整理する書き方です。',
    '見本の語りかけは労務・総務担当ですが、今回の読者は中小企業の経営者・IT 担当者です。',
    '見本の文章・言い回しはそのまま使わず、オリジナルの表現で書いてください。',
    '',
    '<見本記事>',
    SAMPLE_ARTICLE,
    '</見本記事>',
    '',
    '守ること:',
    '- 構成は、導入（300字程度）→ 見出し3つ → まとめ（200字程度）',
    '- 本文全体は 2,000〜3,000 字',
    '- キーワードとサブキーワードを、不自然にならないように入れる',
    '- 製品の押し売りはしない。特定の製品名・サービス名は出さず、課題の整理で終える',
    '- 法令や料金の数字は、確かなものだけを書く',
    '- 見本の文やフレーズを借りない。特に導入の書き出しと、まとめの最初と最後の一文は、見本と違う言い方にする',
    '- 数は算用数字にそろえる（例: 1つ目、2つ目、3つ目。「三つ目」とは書かない）',
    '',
    '最新の法令メモ（あなたの知識より、こちらを優先する）:',
  ].concat(LAW_NOTES.map(function (note) { return '- ' + note; })).concat([
    '',
    '出力は次の JSON だけにしてください。前後に説明文やコードブロックの記号を付けないでください。',
    '{"title": "記事タイトル", "body": "本文の HTML", "metaDescription": "検索結果に出る説明文"}',
    '- body は <h2> と <p> だけを使う。導入は <p>、見出し3つとまとめは <h2> から始める。タイトルは body に入れない',
    '- body のタグの前後や中に改行を入れない',
    '- metaDescription は ' + META_DESCRIPTION_MAX + ' 字以内',
  ]).join('\n');
}

function buildUserPrompt_(keyword, subKeyword, intent) {
  const lines = [
    '次の条件で記事を1本書いてください。',
    '',
    'キーワード: ' + keyword,
    'サブキーワード: ' + (subKeyword || 'なし'),
  ];
  const trimmedIntent = String(intent || '').trim();
  lines.push('検索意図: ' + (trimmedIntent || '指定なし（キーワードから読者の悩みを考えてください）'));
  return lines.join('\n');
}

/**
 * Claude の返答から JSON を取り出し、中身を確かめる
 */
function parseArticleJson_(text) {
  const cleaned = text.replace(/^\s*```(?:json)?\s*/, '').replace(/\s*```\s*$/, '');
  let article;
  try {
    article = JSON.parse(cleaned);
  } catch (e) {
    throw new Error('Claude の返答を JSON として読めませんでした。');
  }

  ['title', 'body', 'metaDescription'].forEach(function (key) {
    if (!article[key] || String(article[key]).trim() === '') {
      throw new Error('Claude の返答に ' + key + ' がありません。');
    }
  });
  if (article.metaDescription.length > META_DESCRIPTION_MAX) {
    throw new Error('metaDescription が ' + META_DESCRIPTION_MAX + ' 字を超えました（' +
      article.metaDescription.length + ' 字）。');
  }
  return article;
}

/**
 * 本文の文字数（HTML タグと改行を除く）
 */
function countBodyChars_(html) {
  return html.replace(/<[^>]*>/g, '').replace(/\s/g, '').length;
}

// ---------------------------------------------------------------
// WordPress へ下書き保存（REST API）
// ---------------------------------------------------------------

// 投稿の状態は下書きだけ。公開は WordPress 上で人が行う
const WP_POST_STATUS = 'draft';
// 見本のサイトには送らない
const WP_BLOCKED_HOST = 'example-saas.co.jp';

/**
 * 記事を WordPress に下書きで保存する
 * スクリプトプロパティ WP_SITE_URL / WP_USER / WP_APP_PASSWORD を使う。
 * どれかが空なら送らず、送る予定の中身を Logger.log に出して null を返す
 * @param {string} title タイトル（E列）
 * @param {string} content 本文の HTML（J列）
 * @param {string} excerpt メタディスクリプション（K列）
 * @return {{id: number, link: string} | null}
 */
function saveDraftToWordPress(title, content, excerpt) {
  const props = PropertiesService.getScriptProperties();
  const siteUrl = String(props.getProperty('WP_SITE_URL') || '').trim().replace(/\/+$/, '');
  const user = props.getProperty('WP_USER');
  const appPassword = props.getProperty('WP_APP_PASSWORD');

  const post = {
    title: title,
    content: content,
    status: WP_POST_STATUS,
    excerpt: excerpt,
  };

  if (!siteUrl || !user || !appPassword) {
    Logger.log('WordPress の設定が空なので、送信しません。送る予定の中身:');
    Logger.log('title: ' + post.title);
    Logger.log('content: ' + post.content);
    Logger.log('status: ' + post.status);
    Logger.log('excerpt: ' + post.excerpt);
    return null;
  }
  if (siteUrl.indexOf(WP_BLOCKED_HOST) !== -1) {
    throw new Error('WP_SITE_URL が見本のサイト（' + WP_BLOCKED_HOST + '）です。送信しません。');
  }

  const response = UrlFetchApp.fetch(siteUrl + '/wp-json/wp/v2/posts', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(user + ':' + appPassword, Utilities.Charset.UTF_8),
    },
    payload: JSON.stringify(post),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  let data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (e) {
    throw new Error('WordPress の返答を読めませんでした（HTTP ' + code + '）。');
  }
  if (code !== 201) {
    const reason = data && data.message ? data.message : 'HTTP ' + code;
    throw new Error('WordPress エラー: ' + reason);
  }
  return { id: data.id, link: data.link };
}

/**
 * 今選んでいる1行を WordPress に下書き保存し、シートに URL と状態を書く
 */
function saveSelectedRowDraft() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection || selection.getRow() === 1) {
    ui.alert('キーワードの行を選んでください。');
    return;
  }
  // 一括では動かさない
  if (selection.getNumRows() > 1) {
    ui.alert('1行だけ選んでください。');
    return;
  }

  const row = selection.getRow();
  const values = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const status = String(values[COL.STATUS - 1]).trim();
  const title = String(values[COL.TITLE - 1]).trim();
  const body = String(values[COL.BODY - 1]).trim();
  const meta = String(values[COL.META - 1]).trim();

  if (isDraftSaved(status)) {
    ui.alert(row + ' 行目はすでに「' + STATUS.DRAFT_SAVED + '」です。上書きはしません。');
    return;
  }
  if (body === '') {
    ui.alert(row + ' 行目は本文が空です。先に「選択した行の記事を作る」を実行してください。');
    return;
  }
  if (status !== STATUS.SHEET_SAVED) {
    ui.alert(row + ' 行目のステータスが「' + (status || '空') + '」です。' +
      '「' + STATUS.SHEET_SAVED + '」の行だけ下書き保存できます。');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('WordPress に下書き保存しています', '記事生成', -1);

  let result;
  try {
    result = saveDraftToWordPress(title, body, meta);
  } catch (e) {
    // 本文（J列）は消さない
    sheet.getRange(row, COL.STATUS).setValue(STATUS.FAILED);
    sheet.getRange(row, COL.MEMO).setValue(e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('', '記事生成', 1);
    ui.alert(row + ' 行目を下書き保存できませんでした。\n\n理由: ' + e.message);
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('', '記事生成', 1);

  // 設定が空のときは送っていないので、シートは変えない
  if (!result) {
    ui.alert('WordPress の設定が空なので、送信していません。\n' +
      'Apps Script の「プロジェクトの設定」→「スクリプト プロパティ」に ' +
      'WP_SITE_URL、WP_USER、WP_APP_PASSWORD を追加してください。');
    return;
  }

  sheet.getRange(row, COL.WP_URL).setValue(result.link);
  sheet.getRange(row, COL.STATUS).setValue(STATUS.DRAFT_SAVED);
  sheet.getRange(row, COL.MEMO).setValue(MEMO_SAVED_TO_WP);
  ui.alert(row + ' 行目を WordPress に下書き保存しました。\n' + result.link);
}
