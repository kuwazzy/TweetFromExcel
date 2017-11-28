Sub TweetFromExcel()

  Dim OAAT 'OAuth Access Token
  Dim OAATS 'OAuth Access Token Secret
  Dim OACID 'OAuth ClientID
  Dim OACIDS 'OAuth ClientID Secret
  Dim TweetsText 'Tweet���镶��
  Dim PicPath1 '�摜�p�X1
  Dim PicPath2 '�摜�p�X2

  '�c�C�[�g���邽�߂̏����V�[�g����擾
  With Sheets("Sheet1")
    OAAT = .Cells(1, 2)
    OAATS = .Cells(2, 2)
    OACID = .Cells(3, 2)
    OACIDS = .Cells(4, 2)
    TweetsText = .Cells(5, 2)
    PicPath1 = .Cells(6, 2)
    PicPath2 = .Cells(7, 2)
  End With

 'CData ExcelAddIn���W���[���Ăяo��
  Set Module = CreateObject("CData.ExcelAddIn.ExcelComModule")
  Module.SetProviderName ("Twitter")
 '�z��̏���
  Dim nameArray
  nameArray = Array("Textparam", "PicPath1", "PicPath2")
  Dim valueArray
  valueArray = Array(TweetsText, PicPath1, PicPath2)
 'Twitter�ւ̐ڑ�������̒�`
  Module.SetConnectionString ("CallbackUrl=http://127.0.0.1/;OAuthAccessToken=" + OAAT + ";OAuthAccessTokenSecret=" + OAATS + ";OAuthClientID=" + OACID + ";OAuthClientSecret= " + OACIDS + ";")
 'Twitter�ւ̃c�C�[�g���s�iInert���j
  If Module.Insert("INSERT INTO   Twitter.Tweets (Text, MediaFilePath#1, MediaFilePath#2) VALUES (@Textparam, @PicPath1, @PicPath2)", nameArray, valueArray) Then
    MsgBox "Tweet Sccess:" & TweetsText
  Else
    MsgBox "Tweet Failed."
  End If
  Module.Close

End Sub
