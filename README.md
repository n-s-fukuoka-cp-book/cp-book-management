# NS高福岡キャンパス図書管理システム
### 2022-10-24Version
***
## このリファレンスの見方
> 関数の種類別の注意点を記載しています。
> 構成は関数の説明と関数の分類分けになります。
> 一般ユーザーは基本的に操作を行うだけで利用できるようにしてください。
> このリファレンスの閲覧は図書委員と教職員のみとします。
> プログラムに変更を加えた場合は必ずリファレンスの変更を行ってください。
> リファレンスの変更を行った場合は必ず変更日時を記載してください。
***
## on Openについて
> この関数について
* この関数は当該シートを開いた際にどのシートでも自動的に利用方法を表示するシートになります。
* 当該シートを開いた際には必ず自動的に表示されます。
 > タイプ
 * 所属シート　⇛　すべてのシートに対して
 * 該当ボタン⇛　なし
 * 実行タイプ　⇛　自動実行
 * 補足情報　⇛　シートを開いてから関数実行までに数秒時間がかかります。
 ***
## DBについて
>　この関数について
* データーベース用シートの更新を行うシートです。
* 現在のところは、一日1000件までの制限がかかっています。
* 実行の際は回数に気をつけてください。
* ISBNをもとにデータをGoogleから検索してるため、**一部データが抜ける場合**があります。
 > タイプ
 * 所属シート　⇛ DB
 * 該当ボタン⇛　Register(),big_in(),test_isbn()
 * 実行タイプ　⇛ 手動実行
 * 補足情報　⇛　一日1000件までの上限あり　プログラムの構成上、セルデータの更新のためにA1列から行うので1000冊までが上限になる。　**処理開始位置を変更することで対応してください**
 ***
## Loadingについて
>　この関数について
* データベースに登録するデータを読み込むシートです。
* DB()の関数の1000件の制限に加えてこちら側でもほんのタイトルを取得するので、**DB（）と合わせて***1000/日の制限があります。
* 登録内容は以下のようになります。
> 図書管理コード（A000から始まる8桁コード）<br>
> ISBN　（本のバーコード　日本図書は基本的に9..から始まる。*例外あり*）<br>
> 図書分類コード　（日本の規格の図書の場所を指定するもの）<br>
* ISBNが一部ない本がある場合はスキップ可能です。
* 必ず管理コードだけは読み込んでください。
 > タイプ
 * 所属シート　⇛DB登録
 * 該当ボタン⇛データ入力
 * 実行タイプ　⇛　手動実行
 * 補足情報　⇛　データ入力だけなので、ここに読み込むだけでは、DBに書き込みされることはありません。
***
## Registerについて
>　この関数について
* Loading（）で読み込んだデータをデータベースに書き込むシートになります。
* DB登録に値がない場合はエラーを返却します。
 > タイプ
 * 所属シート　⇛DB登録
 * 該当ボタン⇛登録する
 * 実行タイプ　⇛手動実行
 * 補足情報　⇛なし
***
## big_inについて
>　この関数について
* この関数は大型導入用に使用する関数になります。
* **こちらの関数は基本的には使用をしないでください。**
* こちらの関数は改行などが手動なので、DBを汚すだけでなく、正確な情報が入力されない可能性（人的ミス）があります。
* こちらの関数は範囲シートから情報を取得して書き込むだけなので、バーコードを事前に直接入力する必要があります。
 > タイプ
 * 所属シート　⇛大型導入用
 * 該当ボタン⇛登録する
 * 実行タイプ　⇛手動実行
 * 補足情報　⇛**使用非推奨**
***
## in_new_dataについて
>　この関数について
* 大型導入用で最終レンタル日を1899/12/30に指定します。
* 大型導入用で貸し出し状況を貸出可に指定します。
* 貸し出し回数を00回に指定します。
 > タイプ
 * 所属シート　⇛大型導入用
 * 該当ボタン⇛なし
 * 実行タイプ　⇛big_in()で自動実行
 * 補足情報　⇛なし
***
## word_searchについて
> この関数について
* 特定のワードをデータベースと照合します。
* 検索シートの枠に貸し出し状況等を記載します。
* **検索できるのは完全一致ワードです。**
* 詳細は以下を確認してください。
* 「ITパスポート」という本を調べる場合
> [IT]→○<br>
> [パス]→○<br>
> [ITパスポート]→○<br>
> [ぱすぽーと]→✗<br>
> [あいてぃー]→✗<br>
> [Iパス]→✗（略称不可）<br>
 > タイプ
 * 所属シート　⇛検索
 * 該当ボタン⇛ワード検索
 * 実行タイプ　⇛手動
 * 補足情報　⇛検索がなかなかしびあです。
***
## search_sheet_clearについて
> この関数について
* 検索シートの検索結果欄をクリアします。
* on Openにも組み込まれています。
 > タイプ
 * 所属シート　⇛検索
 * 該当ボタン⇛Crystal
 * 実行タイプ　⇛手動and自動　実行
 * 補足情報　⇛　~~セル枠が消える場合があるかもしれません。~~　対策済み
***
## test_sibnについて
>この関数について
* テスト処理用のISBNをDBに書き込みます。
* **本番環境で利用している場合は使用を禁止します。**
* DBの最終列から書き足しますが、既存のものとかぶる場合があるので、使用禁止です。
* ISBN書き込み後にDB()を実行します。
* 初期設定では300種類のISBNの書き込みを行います。
* **ISBNはランダムで作り出すので、存在しないものが作成される場合があります。**
 > タイプ
 * 所属シート　⇛なし
 * 該当ボタン⇛なし
 * 実行タイプ　⇛手動　（GAS）
 * 補足情報　⇛**使用非推奨**
***
## rental_startについて
>この関数について
* この関数では本の貸し出しを行います。
* 一度に貸し出し可能な冊数は10冊に制限します。（10冊以上の場合は複数回に分けて処理してください。）
* メールアドレスを入力した場合はスムーズに返却できるバーコードを送信したり、返却のお知らせなどを行います。（詳細は[double_deleteについて],[return_noticeについて]を確認してください。）
* メールアドレスを使用した場合のみ、「メール処理用」に情報が追加されます。理論上は999999999999さつまで登録可能です。
* メールアドレスを使用しない場合はニックネームを入力してください。 
* nnn.ed.jpのドメインで判定を行っていますので、それ以外のアドレスは使用不可です。
* DBシートにアドレスがセットされますが、白文字で非表示にしてください。
 > タイプ
 * 所属シート　⇛貸出
 * 該当ボタン⇛貸出する
 * 実行タイプ　⇛　手動
 * 補足情報　⇛　ニックネームかメールアドレスを入力する必要があり
***
## rental_endについて
>この関数について
* この関数では本の返却を行います。
* 一度に返却可能な冊数は10冊に制限します。（10冊以上の場合は複数回に分けて処理してください。）
* 貸し出し時にメールアドレスを入力していただいた場合は、返却をスムーズにできるバーコード（12桁）がメールアドレス宛に送信されています。
* 手動返却時は管理用バーコード（8桁）を読み込んでください。
* 返却後も次のuserが貸し出しを行うまで、メールアドレス・ニックネームはDBに保存され続けます。
 > タイプ
 * 所属シート　⇛返却
 * 該当ボタン⇛返却する
 * 実行タイプ　⇛手動実行
 * 補足情報　⇛処理内容が分岐する
***

## rental_sheet_clearについて
* 貸し出し、返却の処理用のシートを最後にクリアします。
* 基本的には実行はrental_start(),rental_end()で実行するための物になります。
 > タイプ
 * 所属シート　⇛貸出,返却
 * 該当ボタン⇛なし
 * 実行タイプ　⇛rental_start(),rental_end()で実行
 * 補足情報　⇛　**貸出シートと返却シートの両方**を一度にクリアするため、注意
***
## double_deleteについて
>この関数について
* 同じデータが2度入力された場合に新しい方のデータを削除します。
* rental_start()で実行されますが、貸し出し回数10回のときに同じバーコードを2回吸った場合は　貸出冊数が9階に減ります。
 > タイプ
 * 所属シート　⇛貸出
 * 該当ボタン⇛なし
 * 実行タイプ　⇛rental_start()で実行
 * 補足情報　⇛貸し出し可能冊数が一度間違えると一回減るので注意
***
## return_noticeについて
>この関数について
* 本の返却日前日にメールを送信します。
* メール処理用のシートで今日の日付+1の値を取得した上で、処理状況が0のものに対して、メールを送信します。
* 処理状況2のものは自動的に除外されます。
 > タイプ
 * 所属シート　⇛なし
 * 該当ボタン⇛なし
 * 実行タイプ　⇛自動実行　トリガー
 * 補足情報　⇛16−17時に実行（一時間のうちのランダムで実行）
***
## Expired_return_dateについて
>この関数について
* 本の返却の翌日にメールを送信します。
* メール処理用のシートで今日の日付-1の値を取得した上で、処理状況が0か1のものに対して、メールを送信します。
* 処理状況2のものは自動的に除外されます。
 > タイプ
 * 所属シート　⇛なし
 * 該当ボタン⇛なし
 * 実行タイプ　⇛自動実行　トリガー
 * 補足情報　⇛7−8時に実行 （一時間のうちのランダムで実行）
***
## write_rogについて
* 各関数からrog_msgの引数を持って処理内容を履歴に記載します。
* 処理モードはテストか本番かを入力してください。
* on Open関数を使用した場合は、処理ユーザーがonopen関数になります。
* 最新の履歴が一番上に来ます。
 > タイプ
 * 所属シート　⇛履歴
 * 該当ボタン⇛なし
 * 実行タイプ　⇛各関数実行時
 * 補足情報　⇛なし
***
## 履歴シートの補足データ
* 処理番号0番　→　未処理
* 処理番号１番　→　返却日前日のメール処理完了
* 処理番号2番　→　返却処理完了
* 処理番号3番　→　返却日超過
