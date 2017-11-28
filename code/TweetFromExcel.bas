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
