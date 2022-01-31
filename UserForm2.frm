VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "2年遡及判定"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnBack_Click()
    Unload UserForm2 'UserForm2を非表示
    UserForm1.Show 'UserForm1を表示
End Sub

Private Sub btnJudgement_Click()
    Dim inputValue As String
    Dim convertdate As Date '出力値をDate型に変換(ローカル変数)
    Dim standardDate As Date '起算日の取得
    Dim lastDate As Date '遡及年度6月末日
    Dim result As VbMsgBoxResult 'メッセージボックスの結果
    
    lastDate = CDate(objectYear & "/" & "06" & "/" & "30") '文字列をDATE型に変換
    inputValue = Me.txtYearInput.Text '出力値を変数に代入
    On Error GoTo ErrorLbl
    convertdate = CDate(inputValue) '出力値をDate型に変換(ローカル変数)
    standardDate = convertdate + 1 '遡及起算日を取得(ローカル変数)
    
    '入力値が遡及年度の終了日より後を入力させないための条件分岐
    If convertdate <= finDate Then
        result = MsgBox("『国民健康保険資格取得日』を登録してよろしいですか?" & vbCrLf & "登録年月日:" & CStr(convertdate), Buttons:=vbYesNo)  'MsgBoxの戻り値変数に代入
        'MsgBoxの戻り値で条件分岐
        If result = vbNo Then
            MsgBox "登録を取り消しました。"
            Exit Sub
        Else
            MsgBox "登録しました。" & vbCrLf & "登録年月日:" & CStr(convertdate), Buttons:=vbInformation
        End If
        
        '遡及年度の4月～第一期納付期限は、goBackAbleDate=遡及年度の第一期納付期限
        If convertdate < firstDeadline Then
            goBackAbleDate = DateAdd("yyyy", 2, firstDeadline) '遡及可能年月日の取得
        Else
            goBackAbleDate = DateAdd("yyyy", 2, standardDate) '遡及可能年月日の取得
        End If
        
        'いいえ(基準日後の加入者がいる)が選択された場合の起算日の比較さるかの条件分岐
        If optbtnNotResult = True Then
            If goBackAbleDate > goBackAbleDateComparison Then
               Debug.Print (goBackAbleDate)
               goBackAbleDate = goBackAbleDateComparison
               Debug.Print (goBackAbleDateComparison)
            Else
               Debug.Print (goBackAbleDate)
               Debug.Print (goBackAbleDateComparison)
            End If
        Else
        End If
        
        Debug.Print (goBackAbleDate)
        Me.Hide
        UserForm4.Show
    
    Else
         MsgBox "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度を超過しています。" & vbCrLf & "～" & "令和" & convertToWareki - 1 & "(" & objectYear + 1 & ")" & "年３月31日で入力してください。", Buttons:=vbExclamation
    End If
    
    Exit Sub
ErrorLbl:
        MsgBox "入力値が不正です。"
        Me.txtYearInput.Text = ""
End Sub

Private Sub UserForm_initialize()
    Application.Visible = False
    If optbtnOkResult = True Then
        Me.lblQDisplay2 = "同世帯の加入者の『国民健康保険資格取得日』を入力してください。" & vbCrLf & vbCrLf & "本日は" & ConvertToday & "です。"
    Else
        Me.lblQDisplay2 = "対象者の『国民健康保険資格取得日』を入力してください。" & vbCrLf & vbCrLf & "本日は" & ConvertToday & "です。"
    End If
End Sub
