/*
　wordをGoogleDocに変換して指定したフォルダに保存する関数
*/
function wordTranslator(e) {
  const itemResponses = e.response.getItemResponses();
  
  //Q1 フォームからアップロードされたファイルidを取得する。
  let file = DriveApp.getFileById(itemResponses[0].getResponse());
  file.getId();
  
  //Q2 get folder id by form's
  const folder = DriveApp.getFolderById(itemResponses[1].getResponse());
  options = {
    title: file.getName(),
    mimeType: MimeType.GOOGLE_DOCS,
    parents: [{id: folder.getId()}]
  };
  
  // wordをGoogleDocに変換する。Drive APIへfileをPOSTする
  file = Drive.Files.insert(options, file.getBlob());

  //Q3 Language before translation
  const beforeLanguage = itemResponses[2].getResponse();
  //Q4 Language after translation
  const afterLanguage = itemResponses[3].getResponse();
  //Q5 translation type
  const translationType = itemResponses[4].getResponse();
  //Q6 mail
  const mailAddress = itemResponses[5].getResponse();
  //Q8 fontcolor
  const fontcolor = itemResponses[7].getResponse();

  const doc = DocumentApp.openById(file.id);
  translator(doc, beforeLanguage, afterLanguage, translationType, fontcolor);
  doc.saveAndClose();
  
 const resultDoc = gdoc2docx(file.id,folder);
 const resultPdf = gdoc2pdf(file.id, folder);

 if (mailAddress != ""){
   //Q6 mailsubject
   const mailsubject = itemResponses[5].getResponse();
   sendmail(mailAddress, resultPdf, mailsubject);
 };
}

/*
　翻訳関数
*/
function translator(doc, beforeLanguage, afterLanguage, translationType, fontcolor) {
  var body =  doc.getBody();
  var paragraphs = body.getParagraphs();
  var textBefore ; // 翻訳前テキスト
  var textTranslated ; // 翻訳後テキスト
  let mycolor = color[fontcolor];
  if (mycolor == "") { mycolor = color.lightgray; }
  
  paragraphs.forEach( p => {
    textBefore = p.getText();
    if (textBefore != "" ){

      // 翻訳テキストを原文と置換するタイプ
      if (translationType === "テキスト置換" ){
        textTranslated = LanguageApp.translate(textBefore, language[beforeLanguage], language[afterLanguage]);
        p.setText(textTranslated);

      }else{ //翻訳テキストを原文の後ろに追記タイプ
        var text = p.editAsText();
        var textLength = text.getText().length;

        textTranslated = LanguageApp.translate(textBefore, language[beforeLanguage], language[afterLanguage]);
        p.appendText(" ");
        p.appendText(textTranslated);
        text.setForegroundColor(textLength, textLength + textTranslated.length, mycolor);//gray
      };
    }  
  });
}

/*
　GoogleDocをwordに変換してdriveに保存する関数
*/
function gdoc2docx(gdocId,folder) {
  let new_file;
  const url = "https://docs.google.com/document/d/" + gdocId + "/export?format=docx";
  const options = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() == 200) {
    let doc = DocumentApp.openById(gdocId);
    let filename = doc.getName();
    let pos = filename.indexOf(" - ");
    filename = filename.substring(0, pos);
    new_file = folder.createFile(res.getBlob()).setName("翻訳済" + filename +  ".docx");
  }
  return new_file;
}

/*
　GoogleDocをpdfに変換してdriveに保存する関数
*/
function gdoc2pdf(gdocId,folder) {
  let new_file;
  let url = "https://docs.google.com/document/d/" + gdocId + "/export?exportFormat=pdf";
  
  const options = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() == 200) { // 200=status Ok
    const doc = DocumentApp.openById(gdocId);
    let filename = doc.getName();
    const pos = filename.indexOf(" - ");
    filename = filename.substring(0, pos);
    new_file = folder.createFile(res.getBlob()).setName("翻訳済" + filename +  ".pdf");
  }
  return new_file;
}

/*
　メールを送信する関数
*/
function sendmail(address,file,mailsubject){
  const bodymsg1 = "こんにちは。\nこのメールはGoogle Appsのプログラムから自動送信で送信しています。"
  const bodymsg2 = "Hello,dear.\nThis email is automatically sent from Google Apps."
  GmailApp.sendEmail(
    address , 
    mailsubject,
    bodymsg1 + "\n\n" + bodymsg2,
    {attachments: [file]}
  );
}

const language = {
    "クメール語":"km",
    "キニヤルワンダ語":"rw",
    "ノルウェー語":"no",
    "アラビア文字":"ar",
    "スンダ語":"su",
    "ミャンマー語（ビルマ語）":"my",
    "リトアニア語":"lt",
    "エストニア語":"et",
    "ベラルーシ語":"be",
    "ブルガリア語":"bg",
    "アフリカーンス語":"af",
    "マルタ語":"mt",
    "タタール語":"tt",
    "フランス語":"fr",
    "マレー語":"ms",
    "ポルトガル語（ポルトガル、ブラジル）":"pt",
    "イディッシュ語":"yi",
    "アイルランド語":"ga",
    "モンゴル語":"mn",
    "セブ語":"ceb",
    "サモア語":"sm",
    "カンナダ語":"kn",
    "ボスニア語":"bs",
    "ラテン語":"la",
    "タミル語":"ta",
    "マラヤーラム文字":"ml",
    "オリヤ語":"or",
    "アムハラ語":"am",
    "マケドニア語":"mk",
    "スペイン語":"es",
    "クロアチア語":"hr",
    "インドネシア語":"id",
    "パンジャブ語":"pa",
    "ネパール語":"ne",
    "ショナ語":"sn",
    "エスペラント語":"eo",
    "パシュト語":"ps",
    "アイスランド語":"is",
    "モン語":"hmn",
    "マラガシ語":"mg",
    "タイ語":"th",
    "ヨルバ語":"yo",
    "フィンランド語":"fi",
    "チェコ語":"cs",
    "アルメニア語":"hy",
    "マオリ語":"mi",
    "フリジア語":"fy",
    "ヒンディー語":"hi",
    "ウクライナ語":"uk",
    "トルコ語":"tr",
    "ロシア語":"ru",
    "ベトナム語":"vi",
    "シンハラ語":"si",
    "テルグ語":"te",
    "ポーランド語":"pl",
    "ペルシャ語":"fa",
    "セソト語":"st",
    "タガログ語（フィリピン語）":"tl",
    "ウルドゥー語":"ur",
    "ウイグル語":"ug",
    "アゼルバイジャン語":"az",
    "セルビア語":"sr",
    "イボ語":"ig",
    "ルーマニア語":"ro",
    "スウェーデン語":"sv",
    "ヘブライ語":"he",
    "ラトビア語":"lv",
    "カザフ語":"kk",
    "スワヒリ語":"sw",
    "日本語":"ja",
    "デンマーク語":"da",
    "コルシカ語":"co",
    "ラオ語":"lo",
    "タジク語":"tg",
    "コーサ語":"xh",
    "韓国語":"ko",
    "オランダ語":"nl",
    "ハワイ語":"haw",
    "ルクセンブルク語":"lb",
    "スコットランド ゲール語":"gd",
    "ズールー語":"zu",
    "ガリシア語":"gl",
    "ベンガル文字":"bn",
    "シンド語":"sd",
    "キルギス語":"ky",
    "ハウサ語":"ha",
    "グジャラト語":"gu",
    "英語":"en",
    "クルド語":"ku",
    "ドイツ語":"de",
    "中国語（繁体）":"zh-TW",
    "バスク語":"eu",
    "クレオール語（ハイチ）":"ht",
    "ソマリ語":"so",
    "スロベニア語":"sl",
    "トルクメン語":"tk",
    "グルジア語":"ka",
    "ジャワ語":"jv",
    "カタロニア語":"ca",
    "イタリア語":"it",
    "ウェールズ語":"cy",
    "スロバキア語":"sk",
    "ウズベク語":"uz",
    "アルバニア語":"sq",
    "中国語（簡体）":"zh-CN",
    "ハンガリー語":"hu",
    "ギリシャ語":"el",
    "マラーティー語":"mr",
    "ニャンジャ語（チェワ語）":"ny"
};

const color = {
  "変更しない": "",
  "lightgray": "#d3d3d3",
  "black": "#000000",
  "red": "#ff0000",
  "royalblue": "#4169e1",
  "white": "#ffffff"
}
