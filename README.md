# はじめに

## 概要

本記事では、Excelのマクロを使ってExcelシート内の文章をTwitterにツイートする方法についてご紹介します。本サンプルプログラムを応用することで、Twitterの企業アカウントを管理している方が、日々Tweetする内容をExcelファイルで管理してExcelファイルからボタンひとつでTweetすることを想定しています。

## 前提として必要なアプリ- サービス

- [Twitter](https://twitter.com/)アカウント
- Excel 2016 ※2013以上で利用可能
- [CData Excel Add-In for Twitter](http://www.cdata.com/jp/drivers/twitter/download/excel/)　※30日間利用可能な無償評価版あり

## 利用イメージ

Excelファイル内のボタンをクリックすると、Excelシート内の情報をもとにTwitterにツイート。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/27388df1-5050-141a-109f-5a41066d3223.png)

# 実現方法

# TwitterのOAuth情報作成

まずはじめに、TwitterのAPIを利用時の認可に必要なOAuthキー情報を取得します。本キーを用いてTwitter内の各種情報にアクセスしますので情報が漏れないように管理してください。※一部スクリーンショットのキーの値をマスキングしています。

[Twitter Application Management](https://dev.twitter.com/apps)にアクセス、SignInして、Twitterのアプリケーションを作成します。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/d70bb9bb-8920-286b-c3c1-1b0a66ab11d4.png)

Detailsタブ
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/5b7a5513-bcde-284d-3bdf-4e959729d3da.png)

- Name : CDataTweetFromExcel（任意）
- Description : CDataTweetFromExcel（任意）
- Website : http://127.0.0.1/
- CallbackURL : http://127.0.0.1/


Keys and Access Tokensタブ
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/1c868d25-a5a6-c74c-7273-ad1c656a80ba.png)
※Create my access tokenボタンをクリックしてAccessTokenを取得済

以下5つの値を控えます。
- Consumer Key(API Key)
- Consumer Secret
- Access Token
- Access Token Secret
- Callback URL

## CData Excel Add-In for TwitterのインストールとTwitterへの接続確認

[CData Excel Add-In for Twitter](http://www.cdata.com/jp/drivers/twitter/download/excel/)をダウンロードします。ダウンローしたExe形式のインストーラを起動して、EULAの内容を確認したうえでデフォルトの内容でインストールします。

※注1 途中オンラインアクティベーションが行われるため、インターネットに接続できる環境である必要があります。
※注2 Excelは閉じた状態でインストールしてください。Excelを開いているとインストーラ途中でその旨のメッセージが表示されます。

インストールが完了したらExcelを起動して、CDATAタブに「取得元Twitter」アイコンが表示されることを確認します。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/1fbffdcf-b8db-e3e8-665f-ad5c6329f796.png)

「取得元Twitter」アイコンをクリックします。下記のような「CDataデータ選択」ダイアログが表示されるので「編集」ボタンをクリックします。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/bc5261f0-4fa5-fe72-fd82-4b50cc755867.png)
「接続ウィザード」が起動するので「TwitterのOAuth情報作成」の手順で設定した以下の5つの情報をセットします。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/e3ae03bc-df0e-56b1-11cb-ba7bf78f93c4.png)

| Twitter Application Management設定 |接続ウィザード設定|
|:-----------------|:------------------|
| Consumer Key(API Key)| OAuth Client Id | 
| Consumer Secret| OAuth Client Secret |
| Access Token | OAuth Access Token |
| Access Token Secret| OAuth Access Token Secret |
| Callback URL| Callback URL |

セットしたら「接続テスト」ボタンをクリックして「サーバーに接続できました」ダイアログが表示されます。これでTwitterへの接続ができました。
※接続に失敗した場合は、再度設定内容を確認のうえ、誤りがない場合は、[CDataJapanサポート窓口](http://www.cdata.com/jp/support/)までご連絡ください。

「CDataデータ選択」ダイアログに戻り、テーブル名「Tweetsテーブル」を選択して「OK」ボタンをクリックします。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/d95c095c-41c1-f51e-3202-17c980fbb3bc.png)

すると、Tweetsシートにお持ちのTwitterアカウントのライムライン（メインページ）のツイートが100レコード表示されます。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/77bc2938-ade3-cf18-dfba-037fd9a2e900.png)

これは「CDataデータ選択」ダイアログの「クエリー」でセットされたSQL「SELECT * FROM [CData].[Twitter].[Tweets] LIMIT 100」の結果です。これで、ExcelからTwitterの情報を取得できることを確認できました。

## ツイート用Excelブックの作成

上記で利用したExcelブックを「Excelマクロ有効ブック」で保存します。デフォルトで作成されるSheet1に以下の内容をセットします。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/a73f8502-1c2b-e12f-0ec0-e75567bd6d48.png)

|セルID(A列)|値|セルID(B列) |値|
|:----------|:----------|:----------|:----------|
| A1| OAuth Client Id |  B1| OAuth Client Idの値 | 
| A2| OAuth Client Secret | B2| OAuth Client Secretの値 | 
| A3 | OAuth Access Token | B3|  OAuth Access Tokenの値 | 
| A4| OAuth Access Token Secret | B4| OAuth Access Token Secretの値 |
| A5| ツイート内容 | B4| ツイートする文言 |
| A6| 画像パス1 | B4| ツイートに添付する画像（その１） |
| A7| 画像パス2 | B4| ツイートに添付する画像（その２） |


次に、マクロ起動用のボタンを追加します。Excelの「挿入」タブ内の図形から選択してシートに貼り付けます。テキストに任意の文字列（例では「ツイート」）を追加してください。追加したボタンを右クリックして「マクロの登録」を選択します。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/761c2f66-615f-9377-0751-15a4d782394e.png)
「マクロの登録」ダイアログが表示されるのでマクロ名に「TweetFromExcel」と入力して「新規作成」ボタンをクリックします。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/8820d068-fe98-cf6a-9e74-157e68e1cf65.png)
そうすると、下記のようなExcelVBAを記述するWindowsが起動します。これで、Excelブックの準備はできました。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/ae845ff5-7c70-eef8-8e30-cf2013ee1de3.png)

## ExcelVBAの記述

以下のコードを上記マクロを記述する「標準モジュール」配下のModule1に貼り付けてください。

```
Sub TweetFromExcel()

  Dim OAAT 'OAuth Access Token
  Dim OAATS 'OAuth Access Token Secret
  Dim OACID 'OAuth ClientID
  Dim OACIDS 'OAuth ClientID Secret
  Dim TweetsText 'Tweetする文章
  Dim PicPath1 '画像パス1
  Dim PicPath2 '画像パス2

  'ツイートするための情報をシートから取得
  With Sheets("Sheet1")
    OAAT = .Cells(1, 2)
    OAATS = .Cells(2, 2)
    OACID = .Cells(3, 2)
    OACIDS = .Cells(4, 2)
    TweetsText = .Cells(5, 2)
    PicPath1 = .Cells(6, 2)
    PicPath2 = .Cells(7, 2)
  End With

 'CData ExcelAddInモジュール呼び出し
  Set Module = CreateObject("CData.ExcelAddIn.ExcelComModule")
  Module.SetProviderName ("Twitter")
 '配列の準備
  Dim nameArray
  nameArray = Array("Textparam", "PicPath1", "PicPath2")
  Dim valueArray
  valueArray = Array(TweetsText, PicPath1, PicPath2)
 'Twitterへの接続文字列の定義
  Module.SetConnectionString ("CallbackUrl=http://127.0.0.1/;OAuthAccessToken=" + OAAT + ";OAuthAccessTokenSecret=" + OAATS + ";OAuthClientID=" + OACID + ";OAuthClientSecret= " + OACIDS + ";")
 'Twitterへのツイート実行（Inert文）
  If Module.Insert("INSERT INTO   Twitter.Tweets (Text, MediaFilePath#1, MediaFilePath#2) VALUES (@Textparam, @PicPath1, @PicPath2)", nameArray, valueArray) Then
    MsgBox "Tweet Sccess:" & TweetsText
  Else
    MsgBox "Tweet Failed."
  End If
  Module.Close

End Sub
```
ソースコードは[GitHub](https://github.com/kuwazzy/TweetFromExcel/blob/master/code/TweetFromExcel.bas)からのダウンロードも可能です。

ソースコードを貼り付けたらマクロを実行してみましょう。下記のようなダイアログが出れば成功です。
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/e11da7e9-279f-64bf-25c2-55daaae00109.png)
ExcelVBAを記述するWindowsを保存して閉じます。
※ エラーが発生した場合はデバッグ実行してエラーが発生したところを確認ください。

## Twitterからの実行結果の確認

Twitterにブラウザからログインしてみてみましょう。Excelからツイートした文言、および、画像がツイートされていれば成功です！！
![image.png](https://qiita-image-store.s3.amazonaws.com/0/123181/ab4c9f8d-aa8e-8b27-b63f-7a1967a1b7c2.png)

## 制限事項

- 添付できる画像はTwitterのAPIの制限で3枚までです。
- TwitterのAPIの制約で同じツイート内容を複数回ツイートしようとするとエラーとなります。
- TwitterのAPIはRateLimitが厳しいので繰り返しTweet内容を取得したりTweetしたりするとRateLimit超過のエラーが発生する場合があります。その場合は一定時間置くことで実行できるようになります。

# まとめ

ExcelマクロからTwitterへツイートできるサンプルをご覧いただきました。企業のSNS管理者の方でTwitterの投稿内容をExcelで管理されているような方がいらっしゃれば是非お試しください。ちなみにCDataSoftwareJapanの企業アカウントもCData Excel Add-In for Twitterを利用して各ツイートごとの「リツイート数」などのパフォーマンスを計測しています。本手順のなかで利用している[CData Excel Add-In for Twitter](http://www.cdata.com/jp/drivers/twitter/download/excel/)は30日間ご利用頂ける評価版がございますのでダウンロードして是非お試しください。
