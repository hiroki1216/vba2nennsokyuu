VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "同世帯加入者資格取得日確認"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack2_Click()
    Me.Hide
    UserForm1.Show
End Sub

Private Sub btnNext2_Click()
    Dim inputValue As String
    Dim convertdate As Date '出力値をDate型に変換(ローカル変数)
    Dim standardDate As Date '起算日の取得
    Dim result As VbMsgBoxResult 'メッセージボックスの結果
    Dim convertFirstDeadline As String '遡及年度の第1期納付期限(文字列)
    
    inputValue = Me.txtInputYear2.Text '出力値を変数に代入
    On Error GoTo ErrorLbl
    convertdate = CDate(inputValue) '出力値をDate型に変換(ローカル変数)
    standardDate = convertdate + 1 '遡及起算日を取得(ローカル変数)
    convertFirstDeadline = Format(firstDeadline + 1, "yyyy年m月d日") '遡及年度の第1期納付期限を文字列に変換
    
    '遡及年度の第１期納付期限〜遡及年度の終了日で入力させるための条件分岐
    If convertdate > firstDeadline And convertdate <= finDate Then
        result = MsgBox("『国民健康保険資格取得日』を登録してよろしいですか?" & vbCrLf & "登録年月日:" & CStr(convertdate), Buttons:=vbYesNo)  'MsgBoxの戻り値変数に代入
        'MsgBoxの戻り値で条件分岐
        If result = vbNo Then
            MsgBox "登録を取り消しました。"
            Exit Sub
        Else
            MsgBox "登録しました。" & vbCrLf & "登録年月日:" & CStr(convertdate), Buttons:=vbInformation
        End If
        
        '遡及年度の4月〜6月は、goBackAbleDate=遡及年度の第一期納付期限
        If convertdate < firstDeadline Then
            goBackAbleDateComparison = DateAdd("yyyy", 2, firstDeadline) '遡及可能年月日(比較用)の取得
        Else
            goBackAbleDateComparison = DateAdd("yyyy", 2, standardDate) '遡及可能年月日(比較用)の取得
        End If
        
        Me.Hide
        UserForm2.Show
    Else
         MsgBox "範囲外です。" & vbCrLf & convertFirstDeadline & "〜" & objectYear + 1 & "年３月31日で入力してください。", Buttons:=vbExclamation
    End If
        Exit Sub
ErrorLbl:
        MsgBox "入力値が不正です。"
        Me.txtInputYear2.Text = ""
    
End Sub


Private Sub UserForm_initialize()
Application.Visible = False
Me.lblQD2 = "同世帯の加入者の『国民健康保険資格取得日』を入力してください。" & vbCrLf & vbCrLf & "本日は" & ConvertToday & "です。" & vbCrLf & vbCrLf & "※加入者が複数いる場合は、その中で1番初めに資格を取得した者の取得日を記入してください。"
End Sub
