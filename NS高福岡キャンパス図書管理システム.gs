// README.md https://github.com/n-s-fukuoka-cp-book/cp-book-management#ns%E9%AB%98%E7%A6%8F%E5%B2%A1%E3%82%AD%E3%83%A3%E3%83%B3%E3%83%91%E3%82%B9%E5%9B%B3%E6%9B%B8%E7%AE%A1%E7%90%86%E3%82%B7%E3%82%B9%E3%83%86%E3%83%A0


// シートと展開時のお知らせ画面の表示とサーチシートのクリア
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#on-open%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function onOpen(){
  Browser.msgBox("いつもご利用ありがとうございます。\\n\\nこちらのシートを利用注意点\\n・検索・貸出・返却以外は扱わないようにお願いします。\\n・基本的にセルを編集する動作はないので直接数字を入力などはしないでください\\n　（ボタンを押して処理する動作のみでお願いします。）\\n・メールアドレスを入力する箇所がありますが、ニックネームなどを利用されて構いません。\\n　（メールアドレス入れていただくと、より便利に使うことができます。）\\n\\nなにか不明な点がありましたら、お近くの図書委員にお声掛けください。")
  search_sheet_clear()
}

// DBの情報更新を2行目から最終行までを更新
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#db%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function DB() {
  e_count = 0;
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var db_sheet = active_sheet.getSheetByName("DB");//指定名のシート取得
  var line = db_sheet.getLastRow() - 1//DB件数　検索
  var db_isbn = db_sheet.getRange(2, 2, line) //セルを指定
  var db_values = db_isbn.getValues();//値の取得
  var google_api_url = "https://www.googleapis.com/books/v1/volumes?q=isbn:"//googlebooksAPI-URL
  var google_api_url_end = "&country=jp"//googlebooksAPI-URL-end
  for (var i = 0; i < line; i++) {
    var request_url = google_api_url + db_values[i] + google_api_url_end //request_url文字列結合（行列分）
    var options = {
      'method': 'get',
      "muteHttpExceptions": true,
      "validateHttpsCertificates": false,
      "followRedirects": true
    }//オプション
    // request_urlがからのURLの場合処理をスキップするために中身をnullに変更
    if (request_url == "https://www.googleapis.com/books/v1/volumes?q=isbn:&country=jp") {
      request_url = null
    }
    try {
      // googlebooksにURLのリクエストを送信
      var res = UrlFetchApp.fetch(request_url, options);//URL送信
      var json = JSON.parse(res);
      var count = i + 2;
      // googlebooksの情報を取得・内容を書き込み
      var book_title = json["items"][0]["volumeInfo"]["title"];//タイトル
      var title_write_in = db_sheet.getRange(count, 7);
      var title_values = title_write_in.setValue(book_title);//タイトル書き込み
      var book_authors = json["items"][0]["volumeInfo"]["authors"][0];//著者
      var authors_write_in = db_sheet.getRange(count, 8);
      var authors_values = authors_write_in.setValue(book_authors);//著者書き込み
      var book_beginning = json["items"][0]["volumeInfo"]["description"];//出だし文章
      var beginning_write_in = db_sheet.getRange(count, 9);
      var beginning_values = beginning_write_in.setValue(book_beginning);//出だし文章書き込み
      try {
        // サムネイル画像が取得できる場合のみ取得
        var book_thumbnail = json["items"][0]["volumeInfo"]["imageLinks"]["smallThumbnail"];//サムネイル画像
        var thumbnail_write_in = db_sheet.getRange(count, 10);
        var thumbnail_values = thumbnail_write_in.setValue(book_thumbnail);//サムネイルURL書き込み
      } catch (e) {
        //例外処理でスキップ
        continue
      }

      var book_page = json["items"][0]["volumeInfo"]["pageCount"];//ページ数
      var page_write_in = db_sheet.getRange(count, 11);
      var page_values = page_write_in.setValue(book_page);//ページ数書き込み
      var book_release = json["items"][0]["volumeInfo"]["publishedDate"];//公開日
      var release_write_in = db_sheet.getRange(count, 12);
      var release_values = release_write_in.setValue(book_release);//公開日書き込み 
    } catch (e) {
      e_count = e_count + 1;
      Logger.log(e)
      continue
    }//例外処理
  }
  if (e_count > 0) {
    // エラー内容を表示
    e_code_1 = "処理の途中でエラーが";
    e_code_2 = "回発生しました。<br>DBシートの脱字を確認してください。";
    e_code_3 = "<br><br>実行処理列"
    set_e_code = e_code_1 + e_count + e_code_2 + e_code_3 + line;
    var htmlOutput = HtmlService
      .createHtmlOutput(set_e_code)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error...');
  }
  var rog_msg = "~DBの登録・更新~を実行しました。/DB 処理件数　→" + i + "error件数　→" + e_count;
  write_rog(rog_msg);
}

// 本の登録情報読込（通常）作業を行うプログラム
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#loading%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function Loading() {
  e_count = 0;
  // 登録用（通常）シートを読み込み
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var db_in_sheet = active_sheet.getSheetByName("DB登録");//指定名のシート取得
  var lastrow = db_in_sheet.getLastRow() + 1;//登録リスト　最終行を取得
  // 管理用コード・ISBN・図書コードを順番に登録
  // エラーが発生した場合は処理を停止
  var meetingDate = Browser.inputBox("管理バーコード", Browser.Buttons.OK_CANCEL)
  if (meetingDate == "cancel") {
    Browser.msgBox("登録を中断します。今までの作業分を登録するか、破棄してください。");
    return;
  } else {
    var management = meetingDate
    var management_write_in = db_in_sheet.getRange(lastrow, 1);
    var management_values = management_write_in.setValue(management);
  }//管理コード登録
  var meetingDate = Browser.inputBox("ISBN", Browser.Buttons.OK_CANCEL);
  if (meetingDate == "cancel") {
    Browser.msgBox("登録を中断します。今までの作業分を登録するか、破棄してください。");
    return;
  } else {
    var isbn = meetingDate
    var isbn_write_in = db_in_sheet.getRange(lastrow, 2);
    var isbn_values = isbn_write_in.setValue(isbn);
  }//ISBN登録
  var meetingDate = Browser.inputBox("図書分類コード", Browser.Buttons.OK_CANCEL);
  if (meetingDate == "cancel") {
    Browser.msgBox("登録を中断します。今までの作業分を登録するか、破棄してください。");
    return;
  } else {
    var bookcode = meetingDate
    var bookcode_write_in = db_in_sheet.getRange(lastrow, 3);
    var bookcode_values = bookcode_write_in.setValue(bookcode);
  }//図書コード登録
  // ISBNをもとにしたgooglebooksAPIへのリクエスト
  var google_api_url = "https://www.googleapis.com/books/v1/volumes?q=isbn:"//googlebooksAPI-URL
  var google_api_url_end = "&country=jp"//googlebooksAPI-URL-end
  var request_url = google_api_url + isbn + google_api_url_end //request_url文字列結合（行列分）
  var options = {
    'method': 'get',
    "muteHttpExceptions": true,
    "validateHttpsCertificates": false,
    "followRedirects": true
  }//オプション
  if (request_url == "https://www.googleapis.com/books/v1/volumes?q=isbn:&country=jp") {
    request_url = null
  }
  try {
    // タイトルなどの居本情報を書き込み
    var res = UrlFetchApp.fetch(request_url, options);//URL送信
    var json = JSON.parse(res);
    var book_title = json["items"][0]["volumeInfo"]["title"];//タイトル
    var title_write_in = db_in_sheet.getRange(lastrow, 7);
    var title_values = title_write_in.setValue(book_title);//タイトル書き込み
    var Lending_count_write_in = db_in_sheet.getRange(lastrow, 4);
    // デフォルト数値を書き込み
    var Lending_count = Lending_count_write_in.setValue(0);//貸出件数０件を追加
    var status_write_in = db_in_sheet.getRange(lastrow, 5);
    var status_count = status_write_in.setValue("貸出可");//貸出状況を貸出可を入力
    var day_write_in = db_in_sheet.getRange(lastrow, 6);
    var day_count = day_write_in.setValue("1899/12/30");//貸出状況を貸出可を入力
  } catch (e) {
    // エラーの場合処理をスキップ
    Logger.log("エラー")
    Logger.log(e)
  }//例外処理
  if (e_count > 0) {
    // エラー内容を記載
    e_code_1 = "処理の途中でエラーが";
    e_code_2 = "回発生しました。<br>DBシートの脱字を確認してください。";
    e_code_3 = "<br><br>実行処理列"
    set_e_code = e_code_1 + e_count + e_code_2 + e_code_3 + lastrow;
    var htmlOutput = HtmlService
      .createHtmlOutput(set_e_code)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error...');
  }
  // 正常処理の完了情報をrogに書き込み
  var rog_msg = "~データベース登録用の読込作業~を実行しました。/DB登録";
  write_rog(rog_msg);
}

// 本の登録情報書き込み（通常）作業を行うプログラム
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#register%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function Register() {
  // DB登録シートとDBシートの読込
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var db_in_sheet = active_sheet.getSheetByName("DB登録");//指定名のシート取得
  var in_lastrow = db_in_sheet.getLastRow() - 7;//DB登録リスト　最終行を取得

  var active_in_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var db_sheet = active_in_sheet.getSheetByName("DB");//指定名のシート取得
  var lastrow = db_sheet.getLastRow() + 1;//DBリスト　最終行を取得　
  try {
    // 書き込み用データを取得してDBに書き込む
    var range = db_in_sheet.getRange(8, 1, in_lastrow, 6).getValues();//登録用データ取得
    var range = db_sheet.getRange(lastrow, 1, in_lastrow, 6).setValues(range);//DBに書き込み;
    DB();
    var in_sheet = db_in_sheet.getRange("A8:G107");
    var msg = Browser.msgBox("登録作業が完了しました。データーはDBを確認してください。FORMをクリアします。");
    var in_sheet_crystal = in_sheet.clearContent()
  } catch (e) {
    Logger.log(e)
    var msg = Browser.msgBox("処理の途中でerrorが発生しました。\\n”登録DB”に値が入っていないか、DBの編集権限がない可能性があります。");
  }//例外処理
  var rog_msg = "~読み込んだ作業のDB登録~を実行しました。/DB登録";
  write_rog(rog_msg);
}

// 正規の手順以外でまとめて登録する（非推奨）
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#big_in%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function big_in() {
  var e_count = 0;
  // 注意書きを表示
  var meetingDate = Browser.msgBox("このスクリプトを実行すると、DBに影響を与える可能性があります。\\n通常の場合は”登録DB”を必ず使用してください。\\n処理を続ける場合はOKを、停止する場合はキャンセルを押してください。", Browser.Buttons.OK_CANCEL);
  if (meetingDate == "cancel") {
    Browser.msgBox("処理をキャンセルしました。");
    return;
  } else {
    // 最終確認　ここで処理の停止は不可能です
    var htmlOutput = HtmlService
      .createHtmlOutput("処理を続行します。　右上の✗を押してください。<br>この処理は中断することができません。<br>誤って実行した場合は図書委員に連絡してください。")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '重要');
    in_new_data();
    // 大型導入用のシートとDBを取得　DBの最終列を取得
    var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
    var db_in_sheet = active_sheet.getSheetByName("大型導入用");//指定名のシート取得
    var in_lastrow = db_in_sheet.getLastRow() - 1;//DB登録リスト　最終行を取得
    var active_in_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
    var db_sheet = active_in_sheet.getSheetByName("DB");//指定名のシート取得
    var lastrow = db_sheet.getLastRow() + 1;//DBリスト　最終行を取得　
    try {
      // 登録データを纏めて取得
      var range = db_in_sheet.getRange(2, 1, in_lastrow, 6);
      var values = range.getValues();//登録用データ取得
      // 登録データをDBに書き込み
      var range = db_sheet.getRange(lastrow, 1, in_lastrow, 6);
      var values = range.setValues(values);//DBに書き込み
      DB();
      var in_sheet = db_in_sheet.getRange("A2:F107");
      var msg = Browser.msgBox("登録作業が完了しました。データーはDBを確認してください。FORMをクリアします。");
      var in_sheet_crystal = in_sheet.clearContent()
      // シートをクリア
    } catch (e) {
      // エラー処理　エラー内容を表示する
      var msg = Browser.msgBox("処理の途中でerrorが発生しました。\\n”大型導入用”に値が入っていないか、DBの編集権限がない可能性があります。");
      if (e_count > 0) {
        e_code_1 = "処理の途中でエラーが";
        e_code_2 = "回発生しました。<br>DBシートの脱字を確認してください。";
        e_code_3 = "<br><br>実行処理列"
        set_e_code = e_code_1 + e_count + e_code_2 + e_code_3 + line;
        var htmlOutput = HtmlService
          .createHtmlOutput(set_e_code)
          .setSandboxMode(HtmlService.SandboxMode.IFRAME)
          .setWidth(400)
          .setHeight(100);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error...');
      }
    }//例外処理
  }
  // rogを書き込み
  var rog_msg = "~大型導入用のDB登録~を実行しました。/大型導入用";
  write_rog(rog_msg);
}

// 基本データを記入
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#in_new_data%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function in_new_data() {
  var count = 2;
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var db_in_sheet = active_sheet.getSheetByName("大型導入用");//指定名のシート取得
  var in_lastrow = db_in_sheet.getLastRow() - 1;//DB登録リスト　最終行を取得
  for (var i = 0; i < in_lastrow; i++) {
    var Lending_count_write_in = db_in_sheet.getRange(count, 4);
    var Lending_count = Lending_count_write_in.setValue(0);//貸出件数０件を追加
    var status_write_in = db_in_sheet.getRange(count, 5);
    var status_count = status_write_in.setValue("貸出可");//貸出状況を貸出可を入力
    var day_write_in = db_in_sheet.getRange(count, 6);
    var day_count = day_write_in.setValue("1899/12/30");
    var count = count + 1
    var rog_msg = "~大型導入用の内部処理~を実行しました。/Null";
    write_rog(rog_msg);
  }
}

// ほんのタイトルを検索する　正規一致のもののみ検索可能
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#word_search%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function word_search() {
  search_sheet_clear();
  // inputBoxで検索ワードを入力
  var search_word = Browser.inputBox("キワードを入力してください。", Browser.Buttons.OK_CANCEL)
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var db_sheet = active_sheet.getSheetByName("DB");//指定名のシート取得
  var lastrow = db_sheet.getLastRow() - 1;//DBリスト　最終行を取得
  var in_data = [];
  var word_Cell = db_sheet.getRange(2, 7, lastrow);
  var word_Cell_data = word_Cell.getValues().flat();//ブックタイトル取得
  var word_Cell_data_lan = word_Cell_data.length;
  var authors_Cell = db_sheet.getRange(2, 8, lastrow);
  var authors_Cell_data = authors_Cell.getValues().flat()//著者取得
  var management_Cell = db_sheet.getRange(2, 1, lastrow);
  var management_Cell_data = management_Cell.getValues().flat();//管理用バーコード取得
  var status_Cell = db_sheet.getRange(2, 5, lastrow);
  var status_Cell_data = status_Cell.getValues().flat();//ステータスを取得
  var thumbnail_Cell = db_sheet.getRange(2, 10, lastrow);
  var thumbnail_Cell_data = thumbnail_Cell.getValues().flat();//サムネイル取得
  var page_Cell = db_sheet.getRange(2, 11, lastrow);
  var page_Cell_data = page_Cell.getValues().flat();//ページ数取得
  var rental_day_Cell = db_sheet.getRange(2, 6, lastrow);
  var rental_day_Cell_data = rental_day_Cell.getValues().flat()//貸出日取得
  var thumbnail_Cell = db_sheet.getRange(2, 10, lastrow);
  var thumbnail_Cell_data = thumbnail_Cell.getValues().flat()//URL取得
  var Cell_count = 0;
  var img_url_count = 0;
  try {
    for (var i = 0; i < word_Cell_data_lan; i++) {
      // サーチワードと一致するものをすべてのリストから検索する
      var isExisted = word_Cell_data[i].indexOf(search_word);
      if (isExisted != -1) {
        var rental_day_Cell_data_JST = rental_day_Cell_data[i]
        var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
        var search_sheet = active_sheet.getSheetByName("検索");//指定名のシート取得
        var search_sheet_lastrow = search_sheet.getLastRow();//最終行を取得
        var write_in_data = [];
        if (status_Cell_data[i] == "貸出可") {
          rental_day_Cell_data_JST = "キャンパス在中"
        } else {
          // 貸出中の場合は返却日を取得
          rental_day_Cell_data_JST.setDate(rental_day_Cell_data_JST.getDate() + 14);
          var rental_day_Cell_data_JST = Utilities.formatDate(rental_day_Cell_data[i], "JST", "M/d") + "返却予定";//形式変換
        }
        write_in_data.push(
          [word_Cell_data[i]],
          [authors_Cell_data[i]],
          [status_Cell_data[i]],
          [rental_day_Cell_data_JST],
          [page_Cell_data[i] + "ページ"],
          [thumbnail_Cell_data[i]]
        )//書き込み用配列を作成
        // 書き込みをする
        in_data.push(write_in_data);
        var Cell_count = Cell_count + 1
      }
    }
    // 検索結果の一致件数を入力
    var range = search_sheet.getRange(7, 1, Cell_count, 6);
    var values = range.setValues(in_data);
    var msg = search_sheet.getRange(4, 1)
    var in_msgA = "「"
    var in_msgB = "」の検索結果　"
    var in_msgC = "件ヒットしました。"
    var in_msg = in_msgA + search_word + in_msgB + Cell_count + in_msgC
    var in_msg_ = msg.setValue(in_msg)
  } catch (e) {
    // 一致件数が0の場合はエラーを表示　また何かエラーが起きた場合も処理停止
    set_e_code = "入力いただいたキーワードでは一致する本は見つけられませんでした。<br><br>本のタイトルと同じキーワードを使用してください。<br>タイトル→ITパスポート<br>○→パスポート<br>☓→ぱすぽーと<br><br>もしかして？　こちらのURLに本があった場合はタブレットで読めます！<br>https://bookwalker.jp/search/?word=" + search_word + "&order=score&qsub=1<br><br>NS高福岡キャンパスにない本をお探しの場合は次のリンクから本を探してください。<br>https://books.google.co.jp/";
    // 図書委員会への本の購入リクエストは次のフォームから送信してください。<br>https://forms.gle/Ho4ZfuBXZp3XnRLr7<br><br>
    var htmlOutput = HtmlService
      .createHtmlOutput(set_e_code)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(650)
      .setHeight(410);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'お探しのものを見つけられませんでした。');
  }
  // 検索結果をrogに書き込む
  var rog_msg = "~指定ワードの検索~を実行しました。/検索　検索ワード→" + search_word + "/  ヒット件数→" + Cell_count;
  write_rog(rog_msg);
}

// 検索用シートをクリアする
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#search_sheet_clear%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function search_sheet_clear() {
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var search_sheet = active_sheet.getSheetByName("検索");//指定名のシート取得
  var search_sheet_cell = search_sheet.getRange("A7:F100");//検索シートの削除範囲を指定
  var in_sheet_crystal = search_sheet_cell.clearContent();
  var search_sheet_cell = search_sheet.getRange("A4");//A4のデータを削除
  var in_sheet_crystal = search_sheet_cell.clearContent();
  var rog_msg = "~検索結果をクリア~を実行しました。/検索";
  write_rog(rog_msg);
}

// テスト用ISBNを書き込み
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#search_sheet_clear%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function test_isbn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var sheet = sheet.getSheetByName("DB");//指定名のシート取得
  var lastrow = sheet.getLastRow() + 1;
  var isbn = 9784094030690;
  for (var i = 1; i < 300; i++) {
    var cell = sheet.getRange(lastrow, 2);
    var data = cell.setValue(isbn);
    var isbn = isbn + 24;
    var lastrow = sheet.getLastRow() + 1;
  }
  // DBの更新とrogの書き込み
  DB();
  var rog_msg = "~テスト用データをDBに書き込み~を実行しました。/Null/DB";
  write_rog(rog_msg);
}

// 貸出を登録する
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#rental_start%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function rental_start() {
  var mail_text = []
  var rental_sheet_list = []
  var data_count = []
  var management = []
  // レンタルシートをクリア
  rental_sheet_clear();
  // 貸し出しする本の冊数から繰り返し行う回数を決める
  var book_count = Browser.inputBox("貸出をする本の冊数を入力してください\\n(10冊以上の場合は複数回に分けてください。)", Browser.Buttons.OK_CANCEL)
  if (book_count == "cancel") {
    Browser.msgBox("登録を中断します。今までの作業分を登録するか、破棄してください。");
    var rog_msg = "~貸出処理を中断~を実行しました。/貸出";
    write_rog(rog_msg);
    return;
  } else if (0 >= book_count) {
    Browser.msgBox("入力する数値は１冊以上１０冊以内にしてください。")
    var rog_msg = "~貸出処理を中断~を実行しました。/貸出";
    write_rog(rog_msg);
    return;
  } else if (10 < book_count) {
    Browser.msgBox("入力する数値は１冊以上１０冊以内にしてください。")
    var rog_msg = "~貸出処理を中断~を実行しました。/貸出";
    write_rog(rog_msg);
    return;
  } else {
    var user_name = Browser.inputBox("ユーザーネームかメールアドレスを入力してください。\\n---------------------------------------------------------\\nメールアドレスを入力した場合は返却日付の前に\\nメールが送信されたりなどの便利機能が使うことができます。\\nメールアドレス以外が入力されるとメールは送信されません。\\n---------------------------------------------------------\\n使えるドメインは@nnn.ed.jpのみです。", Browser.Buttons.OK_CANCEL)
    Browser.msgBox("["+user_name+"]で処理を開始します。")
    if (user_name == "cancel") {
      Browser.msgBox("登録を中断します。初めからやり直してください。");
      var rog_msg = "~貸出処理を中断~を実行しました。/貸出";
      write_rog(rog_msg);
      return;
    } else {
      for (var i = 0; i < book_count; i++) {
        // バーコードを読み取り作業を開始
        var management_code = Browser.inputBox("貸出する本のNS高が貼ったバーコードを入力して下さい", Browser.Buttons.OK_CANCEL)
        var management_code_len = management_code.length//追加処理　　
        if (management_code == "cancel") {
          Browser.msgBox("登録を中断します。今までの作業分は自動で登録されています。\\n未登録分を再度処理してください");
          var rog_msg = "~貸出処理を中断~を実行しました。/貸出";
          write_rog(rog_msg);
          return;
        } else {
          var management_code=Number(management_code)
          if (management_code_len != 8) {
            rental_sheet_clear();
            var error_msg = Browser.msgBox("管理用バーコードを入力してください。　\\n本の裏に貼ってるA10......から始まるバーコードです。\\n今までの作業分は自動で登録されています。\\n未登録分を再度処理してください\\n今、問題が発生した本から再度吸い直してください。");
            var rog_msg = "~貸出処理を中断/管理コードの桁数が正しくありません~を実行しました。/貸出";
            write_rog(rog_msg);
            return;
          } else {
            //正のときの処理を記入。

            var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
            var sheet = active_sheet.getSheetByName("貸出");//指定名のシート取得
            var sheet_lastrow = sheet.getLastRow() + 1;
            var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
            var db_sheet = active_sheet.getSheetByName("DB");//指定名のシート取得
            var lastrow = sheet.getLastRow() + 1;//貸し出し用シートの最終列取得
            var db_lastrow = db_sheet.getLastRow();
            var db_management_code = db_sheet.getRange(2, 1, db_lastrow);
            var db_management_code_data = db_management_code.getValues().flat();//管理コード取得
            var db_data_list = []
            var word_Cell = db_sheet.getRange(2, 7, db_lastrow);
            var word_Cell_data = word_Cell.getValues().flat();//ブックタイトル取得
            var isExisted = db_management_code_data.indexOf(management_code);
            // 一致する管理コードを検索する
            if (isExisted != -1) {
              db_data_list.push([
                [management_code],
                [word_Cell_data[isExisted]]
              ])
              mail_text.push(word_Cell_data[isExisted])//本タイトル
              data_count.push([isExisted + 2])//index
              management.push(management_code)//管理コード
              rental_sheet_list.push([
                [management_code],
                [word_Cell_data[isExisted]]
              ])
              var in_data = sheet.getRange(lastrow, 1);
              var write_data1 = in_data.setValue(db_data_list[0][0]);
              var in_data2 = sheet.getRange(lastrow, 2);
              var write_data2 = in_data2.setValue(db_data_list[0][1]);
            }
            Logger.log(isExisted)
            //+2で実行列に相当　関数で配列を作ったあと、for文で繰り返し処理を記述　配列文繰り返して貸出と返却のステータスを書き換える。
          }
          double_delete();
          var index = isExisted + 2
          var rental_count_Cell = db_sheet.getRange(index, 4);
          var status_Cell = db_sheet.getRange(index, 5);
          var rental_start_day_Cell = db_sheet.getRange(index, 6);
          var user_Cell = db_sheet.getRange(index, 13);
          var db_count_cell = rental_count_Cell.getValue();
          var count = db_count_cell + 1;
          var write = rental_count_Cell.setValue(count);
          var write = status_Cell.setValue("貸出中");
          var write = user_Cell.setValue(user_name)
          var date = new Date()
          var write = rental_start_day_Cell.setValue(date);
          var rood = sheet.getRange(sheet_lastrow, 3);
          // 貸出処理が完了したら●を追加する
          var write = rood.setValue("🔴");
          if (i == 0) {
            var msgBox = Browser.msgBox("🔴がついている分は登録が完了しています。\\n続きがある場合は場合は処理を実行してください")
          }
        }
      }
    }
    var sheet_lastrow = sheet.getLastRow() + 1
    var rood = sheet.getRange(sheet_lastrow, 3);
    var write = rood.setValue("🔴");
    Browser.msgBox("処理を完了しました。フォームクリアを実行します。")
    rental_sheet_clear();
    var rog_msg = "~貸出~を実行しました。/貸出/" + book_count + "冊の本の貸出を" + user_name + "が行いました。";
    write_rog(rog_msg);
    var mail_address = user_name
    if (user_name.length >= 10) {
      // 10桁以上の処理
      var check = (user_name.substring(user_name.length - 10))
      if (check == "@nnn.ed.jp") {
        var date = new Date();
        date.setDate(date.getDate() + 14);
        var date = Utilities.formatDate(date, "JST", "Y/M/d")
        try {
          var rog_msg = "~貸出~を実行しました。/貸出" + mail_address + "に貸出案内のメールを送信しました。Gmail残り回数→" + MailApp.getRemainingDailyQuota();
          //メール処理用
          var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
          var mail_sheet = active_sheet.getSheetByName("メール処理用");//指定名のシート取
          var get_rental_code = mail_sheet.getRange(2, 1);
          var rental_end_code = get_rental_code.getDisplayValue();
          var rental_end_code = Number(rental_end_code) + 1;
          mail_sheet.insertRows(2, 1);
          var write = get_rental_code.setValue(rental_end_code);
          var mail_address_cell = mail_sheet.getRange(2, 2);
          var date_cell = mail_sheet.getRange(2, 3);
          var mail_url_cell = mail_sheet.getRange(2, 4);

          var management_code_cell = mail_sheet.getRange(2, 5);
          var book_title_cell = mail_sheet.getRange(2, 6)
          var index_cell = mail_sheet.getRange(2, 7)
          // 返却用コードを作成
          var mail_url = "https://www.webarcode.com/barcode/image.php?code=" + rental_end_code + "&type=C128B&xres=1&width=206&output=png&style=197"
          
          var write = mail_address_cell.setValue(mail_address);
          var write = date_cell.setValue(date);
          var write = mail_url_cell.setValue(mail_url);
          var write = management_code_cell.setValue(management + "");
          var write = book_title_cell.setValue(mail_text.toString());
          var write = index_cell.setValue(data_count.toString());
          var write = mail_sheet.getRange(2, 8).setValue(0);
          //メール書利用
          var recipient = mail_address;//送信先のメールアドレス
          var subject = '本の貸出を受け付けました'; 　　     　 　//件名
          // メール本文を作成
          var body = ("いつもご利用ありがとうございます。\n\nNS高等学校福岡キャンパス図書管理システムです。\n\n本の貸出を確認いたしましたのでお知らせいたします。\n貸出の確認ができた本は[" + book_count + "]冊です。\n\n本日貸し出した本の返却日は[" + date + "]までになります。\n\n詳細は以下を確認してください。\n\n貸出した本\n---------------------------\n" + mail_text.join("\n\n") + "\n---------------------------\n返却時に以下のバーコードを使うとすぐに返却することができます！\n\n" + mail_url+"\n\n　一部本のタイトルが表示されない場合がありますが、異常ではありません。")
          const options = { name: 'NS高福岡キャンパス図書委員会:図書管理システム【自動送信】' };  //送信者の名前
          GmailApp.sendEmail(recipient, subject, body, options);//メール送信処理
          write_rog(rog_msg);
        } catch (e) {
          // 送信に失敗した場合の処理
          Browser.msgBox("メールの送信に失敗しました。\\n本のレンタルは完了しています。\\nお手数ですが、一度図書委員会にお声掛けください。\\nまたは以下のアドレスへ連絡ください\\n n.s.fukuoka.cp.book@gmail.com")
          var rog_msg = "~貸出~でエラーが発生しました。/貸出　メール送信中にエラーが発生しました。　GASmail制限回数→" + MailApp.getRemainingDailyQuota() + "でした。";
          write_rog(rog_msg);
        }
      }
    }
  }
}

// 返却処理を実行
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#rental_end%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function rental_end() {
  rental_sheet_clear();
  var book_count = Browser.inputBox("返却をする本の冊数を入力してください\\n貸出のときにメールに届いたバーコードを使用する方\\nはここで読み込んでください。\\n(10冊以上の場合は複数回に分けてください。)", Browser.Buttons.OK_CANCEL)
  if (book_count == "cancel") {
    Browser.msgBox("登録を中断します。今までの作業分を登録するか、破棄してください。");
    var rog_msg = "~返却処理を中断~を実行しました。/返却";
    write_rog(rog_msg);
    return;
  } else if (book_count.length === 12) {
    try {
      // 返却用コードを利用する場合の処理
      var rental_book_list = [];
      var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
      var mail_sheet = active_sheet.getSheetByName("メール処理用");//指定名のシート取得
      var book_count = (Number(book_count))
      var lastrow = mail_sheet.getLastRow();
      var search_cell = mail_sheet.getRange(1, 1, lastrow).getValues().flat();
      var isExisted = search_cell.indexOf(book_count);
      if (isExisted != -1) {
        var getcell = isExisted + 1;
        //check
        var check = mail_sheet.getRange(getcell, 8).getValue();
        //checkが3のときに注意書きを書く処理を追加
        if (check == "0" || check == "1" || check == "3") {
          Logger.log("正解の処理")
          //check
          // 管理コードとほんのタイトル　インデックスを取得
          var wirte = mail_sheet.getRange(getcell, 8).setValue(2);
          var management_code = mail_sheet.getRange(getcell, 5).getValue();
          var book_title = mail_sheet.getRange(getcell, 6).getValue();
          var index = mail_sheet.getRange(getcell, 7).getValue().toString();
          // 取得データから,で文字を分割
          var write_management = management_code.split(",");
          var write_book_title = book_title.split(",");
          var write_index = index.split(",");
          var rental_end_list_count = write_management.length;

          var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
          var sheet = active_sheet.getSheetByName("返却");//指定名のシート取得

          var text_change = sheet.getRange("C3");
          var write = text_change.setValue("参照列番");
          var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
          var db_sheet = active_sheet.getSheetByName("DB");//指定名のシート取得
          var book_title_list = [];
          for (var i = 0; i < rental_end_list_count; i++) {
            // 処理コードを入力して処理状況を変更
            rental_book_list.push([
              [write_management[i]],
              [write_book_title[i]],
              [write_index[i]]
            ])
            book_title_list.push(write_book_title[i])

          }
          var rental_list_cell = sheet.getRange(4, 1, rental_end_list_count, 3);
          var rental_list_write = rental_list_cell.setValues(rental_book_list);

          for (var i = 0; i < rental_end_list_count; i++) {
            var status_Cell = db_sheet.getRange(write_index[i], 5);
            var write = status_Cell.setValue("貸出可");
          }
          var text_change = sheet.getRange("C3");
          var write = text_change.setValue("処理状況");
          // 返却処理完了のお知らせを表示
          Browser.msgBox("返却処理完了しました。\\n図書委員一同、またのご利用を待ちしております。\\n\\n\\n処理した本\\n------------------------------------------------------------\\n" + book_title_list.join("\\n") + "\\n------------------------------------------------------------\\n返却番号[" + book_count + "]");
          var rog_msg = "~レンタル番号バーコード返却~を実行しました。/返却  返却用処理番号→" + book_count;
          write_rog(rog_msg);
          rental_sheet_clear();
        } else {
          // 処理が完了しているコードが入力された場合は処理を停止
          Browser.msgBox("このバーコードは処理が完了しています。\\n間違いだと思われる場合は図書委員へお知らせください。")
        }
      } else {
        // 利用できないコードが合った場合は処理を停止
        Browser.msgBox("この処理番号は無効です。\\n間違いだと思われる場合は図書委員へお知らせください。")
      }
    } catch (e) {
      // エラーが発生した場合の処理
      Browser.msgBox("大変申し訳ございません。\\n返却の手続きの途中でエラーが発生しました。\\nお手数ですが、図書委員へお知らせください。");
      var rog_msg = "~まとめて返却の途中でエラーが発生しました。~を実行しました。/返却　エラー内容　→" + e;
      write_rog(rog_msg);
      Logger.log(e)
      return;
    }
  } else if (0 >= book_count) {
    // 管理用コード以外のコードが入力された場合に処理停止
    Browser.msgBox("入力する数値は１冊以上１０冊以内にしてください。");
    var rog_msg = "~返却処理を中断~を実行しました。/返却";
    write_rog(rog_msg);
    return;
  } else if (10 < book_count) {
    // 管理用コード以外のコードが入力された場合に処理停止
    Browser.msgBox("入力する数値は１冊以上１０冊以内にしてください。");
    var rog_msg = "~返却処理を中断~を実行しました。/返却";
    write_rog(rog_msg);
    return;
  } else {
    // 冊数指定で入力するときの処理
    for (var i = 0; i < book_count; i++) {
      var management_code = Browser.inputBox("貸出する本のN高が貼ったバーコードを入力して下さい", Browser.Buttons.OK_CANCEL);
      var management_code_len = management_code.length
      if (management_code == "cancel") {
        Browser.msgBox("登録を中断します。今までの作業分は自動で登録されています。\\n未登録分を再度処理してください");
        var rog_msg = "~返却処理を中断~を実行しました。/返却";
        write_rog(rog_msg);
        return;
      } else {
        if (management_code_len != 8) {
          rental_sheet_clear();
          var error_msg = Browser.msgBox("管理用バーコードを入力してください。　\\n本の裏に貼ってるA10......から始まるバーコードです。\\n今までの作業分は自動で登録されています。\\n未登録分を再度処理してください");
          var rog_msg = "~返却処理を中断/管理コードの桁数が正しくありません~を実行しました。/返却";
          write_rog(rog_msg);
          return;
        } else {
          //正のときの処理を記入。
          var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
          var sheet = active_sheet.getSheetByName("返却");//指定名のシート取得
          var sheet_lastrow = sheet.getLastRow() + 1;
          var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
          var db_sheet = active_sheet.getSheetByName("DB");//指定名のシート取得
          var lastrow = sheet.getLastRow() + 1;//貸し出し用シートの最終列取得
          var db_lastrow = db_sheet.getLastRow();
          var db_management_code = db_sheet.getRange(2, 1, db_lastrow);
          var db_management_code_data = db_management_code.getValues().flat();//管理コード取得
          db_management_code_data.toString();
          var db_data_list = []
          management_code=(Number(management_code))
          var word_Cell = db_sheet.getRange(2, 7, db_lastrow);
          var word_Cell_data = word_Cell.getValues().flat();//ブックタイトル取得
          var isExisted = db_management_code_data.indexOf(management_code);
          if (isExisted != -1) {
            db_data_list.push([
              [management_code],
              [word_Cell_data[isExisted]]
            ])
          }
          Logger.log (management_code)
          Logger.log(db_management_code_data)

          var in_data = sheet.getRange(lastrow, 1);
          var write_data1 = in_data.setValue(db_data_list[0]);
          var in_data2 = sheet.getRange(lastrow, 2);
          var write_data2 = in_data2.setValue(db_data_list[1]);
        }
        Logger.log(isExisted)
        //+2で実行列に相当　関数で配列を作ったあと、for文で繰り返し処理を記述　配列文繰り返して貸出と返却のステータスを書き換える。
      }
      double_delete();
      // 返却処理が完了した本に🔷を記入
      var index = isExisted + 2
      var rental_count_Cell = db_sheet.getRange(index, 4);
      var status_Cell = db_sheet.getRange(index, 5);
      var write = status_Cell.setValue("貸出可");
      var rood = sheet.getRange(sheet_lastrow, 3);
      var write = rood.setValue("🔷");
      if (i == 0) {
        var msgBox = Browser.msgBox("🔷がついている分は登録が完了しています。\\n続きがある場合は場合は処理を実行してください")
      }
    }
    var sheet_lastrow = sheet.getLastRow() + 1
    var rood = sheet.getRange(sheet_lastrow, 3);
    var write = rood.setValue("🔷");
    Browser.msgBox("返却を完了しました。\\n図書委員一同、またのご利用を待ちしております。")
    rental_sheet_clear();
  }
}

// レンタル用シートをクリア
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#rental_sheet_clear%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function rental_sheet_clear() {
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var sheet = active_sheet.getSheetByName("貸出");//指定名のシート取得
  var sheet_cell = sheet.getRange("A4:C13");
  var sheet_crystal = sheet_cell.clearContent()
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var sheet = active_sheet.getSheetByName("返却");//指定名のシート取得
  var sheet_cell = sheet.getRange("A4:C13");
  var sheet_crystal = sheet_cell.clearContent()
  var rog_msg = "~貸出・返却シートをクリア~を実行しました。/貸出/返却";
  write_rog(rog_msg);
}

// かぶっているデータを削除
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#double_delete%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function double_delete() {
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var sheet = active_sheet.getSheetByName("貸出");//指定名のシート取得
  var double_delete_cell = sheet.getRange("A4:B13");
  double_delete_cell.removeDuplicates([1]);
  double_delete_cell.setBorder(true, true, true, true, true, true);
  var rog_msg = "~データかぶりを削除~を実行しました。/貸出/返却";
  write_rog(rog_msg);
}

// 本の返却期限のお知らせメール送信
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#return_notice%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function return_notice() {
  var ture_count = 0;
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var mail_sheet = active_sheet.getSheetByName("メール処理用");//指定名のシート取得
  var lastrow = mail_sheet.getLastRow() - 1;
  var status = mail_sheet.getRange(2, 8, lastrow).getValues();
  Logger.log(status)
  var date_list = mail_sheet.getRange(2, 3, lastrow).getValues().flat();//Sat Oct 22 00:00:00 GMT+09:00 2022
  Logger.log(date_list)
  var date_list_str = []
  for (var i = 0; i < date_list.length; i++) {
    var dates = Utilities.formatDate(date_list[i], 'JST', 'yyyy-MM-dd').toString();
    date_list_str.push(dates)
  }
  Logger.log(date_list_str)
  var today = new Date();
  today.setDate(today.getDate() + 1);
  var today = Utilities.formatDate(today, 'JST', 'yyyy-MM-dd').toString();
  Logger.log(today);
  var count = status.length;
  Logger.log(count)
  var zero_list = [];
  for (var i = 0; i < count; i++) {
    var isExisted = status[i].indexOf(0)
    if (isExisted != -1) {
      zero_list.push(i)
    }
  }
  Logger.log(zero_list)
  //if文をかいて　翌日の日付と比較して　turuの場合のみメール送信　それ以外はspkip処理
  try {
    for (var i = 0; i < zero_list.length; i++) {
      var get_lastrow = zero_list[i] + 2
      var mail_address = mail_sheet.getRange(get_lastrow, 2).getValue();
      var mail_url = mail_sheet.getRange(get_lastrow, 4).getValue();
      var book_title = mail_sheet.getRange(get_lastrow, 6).getValue().split(",");
      var count = mail_sheet.getRange(get_lastrow, 5).getValue().split(",");
      var recipient = mail_address;//送信先のメールアドレス
      var day = mail_sheet.getRange(get_lastrow, 3).getValue();
      var day_str = Utilities.formatDate(day, 'JST', 'yyyy-MM-dd').toString();

      var subject = '【重要】本の返却期限が迫っています'; 　　     　 　//件名
      var body = ("いつもご利用ありがとうございます。\n\nNS高等学校福岡キャンパス図書管理システムです。\n\n返却期限が明日に迫っている本がありますので、お知らせいたします。\n\n返却期限が迫ってる本が[" + count.length + "]冊あります。\n\n詳細は以下を確認してください。\n\返却期限が迫っている本\n---------------------------\n" + book_title.join("\n\n") + "\n---------------------------\n返却時に以下のバーコードを使うとすぐに返却することができます！\n\n" + mail_url+"\n\n　一部本のタイトルが表示されない場合がありますが、異常ではありません。")
      var options = { name: 'NS高福岡キャンパス図書委員会:図書管理システム【自動送信】' };  //送信者
      if (day_str === today) {
        GmailApp.sendEmail(recipient, subject, body, options);//メール送信処理
        var write = mail_sheet.getRange(get_lastrow, 8).setValue(1)
        var ture_count = ture_count + 1;
      } else {
        continue;
      }
    }
    var rog_msg = "~返却お知らせ~を実行しました。メール処理用　処理件数→　" + ture_count + "    Gmail残り回数→" + MailApp.getRemainingDailyQuota();
    write_rog(rog_msg);
  } catch (e) {
    var rog_msg = "~返却前日の処理が失敗しました。~/メール処理用";
    write_rog(rog_msg);
    return;
  }

}

//返却期限が過ぎている本のリストアップとメール送信
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#expired_return_date%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function Expired_return_date() {
  var ture_count = 0;
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var mail_sheet = active_sheet.getSheetByName("メール処理用");//指定名のシート取得
  var lastrow = mail_sheet.getLastRow() - 1;
  var status = mail_sheet.getRange(2, 8, lastrow).getValues();
  Logger.log(status)
  var date_list = mail_sheet.getRange(2, 3, lastrow).getValues().flat();//Sat Oct 22 00:00:00 GMT+09:00 2022
  Logger.log(date_list)
  var date_list_str = []
  for (var i = 0; i < date_list.length; i++) {
    var dates = Utilities.formatDate(date_list[i], 'JST', 'yyyy-MM-dd').toString();
    date_list_str.push(dates)
  }
  Logger.log(date_list_str)
  var today = new Date();
  today.setDate(today.getDate() - 1);
  var today = Utilities.formatDate(today, 'JST', 'yyyy-MM-dd').toString();

  Logger.log(today);
  var count = status.length;
  Logger.log(count)
  var zero_list = [];
  for (var i = 0; i < count; i++) {
    var isExisted = status[i].indexOf(1)
    if (isExisted != -1) {
      zero_list.push(i)
    }
  }
  Logger.log(zero_list)//if文をかいて　翌日の日付と比較して　turuの場合のみメール送信　それ以外はspkip処理　(return処理はNG)2022/10/15
  try {
    for (var i = 0; i < zero_list.length; i++) {
      var get_lastrow = zero_list[i] + 2
      var mail_address = mail_sheet.getRange(get_lastrow, 2).getValue();
      var mail_url = mail_sheet.getRange(get_lastrow, 4).getValue();
      var book_title = mail_sheet.getRange(get_lastrow, 6).getValue().split(",");
      var count = mail_sheet.getRange(get_lastrow, 5).getValue().split(",");
      var recipient = mail_address;//送信先のメールアドレス
      var day = mail_sheet.getRange(get_lastrow, 3).getValue();
      var day_str = Utilities.formatDate(day, 'JST', 'yyyy-MM-dd').toString();

      var subject = '【超重要】本の返却期限が過ぎています！！'; 　　     　 　//件名
      var body = ("いつもご利用ありがとうございます。\n\nNS高等学校福岡キャンパス図書管理システムです。\n\n返却期限過ぎている本がありますので、お知らせいたします。\n\n返却期限を過ぎている本が[" + count.length + "]冊あります。\n\n詳細は以下を確認してください。\n速やかに返却をお願いします。\n場合によっては図書委員からお声掛けさせて頂く場合がありますのでご了承ください。\n\n返却期限を過ぎている本\n---------------------------\n" + book_title.join("\n\n") + "\n---------------------------\n返却時に以下のバーコードを使うとすぐに返却することができます！\n\n" + mail_url+"\n\n　一部本のタイトルが表示されない場合がありますが、異常ではありません。")
      const options = { name: 'NS高福岡キャンパス図書委員会:図書管理システム【自動送信】' };  //送信者
      if (day_str === today) {
        GmailApp.sendEmail(recipient, subject, body, options);//メール送信処理
        var write = mail_sheet.getRange(get_lastrow, 8).setValue(3)
        Logger.log("メール送信済み")
        var ture_count = ture_count+1
      } else {
        continue;
      }
    }
    var rog_msg = "~返却期限超過のお知らせ~を実行しました。メール処理用　処理件数→　" + ture_count + "    Gmail残り回数→" + MailApp.getRemainingDailyQuota();
    write_rog(rog_msg);
  } catch (e) {
    var rog_msg = "~返却翌日の最終メール送信処理が失敗しました。~/メール処理用";
    write_rog(rog_msg);
    return;
  }
}
// 様々なログを書き込む関数
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#write_rog%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6
function write_rog(rog_msg) {
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet(); //現在のシート取得
  var rog_sheet_get = active_sheet.getSheetByName("履歴");//指定名のシート取得
  rog_sheet_get.insertRows(2, 1);
  var mode = "運用処理モード"
  // var mode = "テスト処理モード"

  var date = new Date()
  var user_name = Session.getActiveUser();
  if (user_name == "") {
    user_name = "onOpen関数"
  }
  var sheet = rog_sheet_get.getRange("A2");
  var write = sheet.setValue(date);
  var sheet = rog_sheet_get.getRange("B2");
  var write = sheet.setValue(mode + " : " + rog_msg);
  var sheet = rog_sheet_get.getRange("C2");
  var write = sheet.setValue(user_name);
}

// 補足情報
// https://github.com/n-s-fukuoka-cp-book/cp-book-management#%E5%B1%A5%E6%AD%B4%E3%82%B7%E3%83%BC%E3%83%88%E3%81%AE%E8%A3%9C%E8%B6%B3%E3%83%87%E3%83%BC%E3%82%BF

function inventory(){
  var db_sheet = SpreadsheetApp.getActiveSpreadsheet();
  var db_sheet = db_sheet.getSheetByName("棚卸");
  

}
