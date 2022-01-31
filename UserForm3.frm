VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "遡及起算日登録"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBackToTop_Click()
    '基準日登録の有無で条件分岐
    If firstDeadline = "0:00:00 " Then
        MsgBox "基準日を登録してください。", Buttons:=vbExclamation
    Else
        Unload Me
        Unload UserForm1
        UserForm1.Show
    End If
End Sub

Private Sub btnDebug_Click()
    Application.Visible = True
End Sub

Private Sub btnRegistration_Click()
    Dim inputValue As String '入力値
    Dim result As VbMsgBoxResult 'メッセージボックスの結果
    Dim convertdate As Date '入力値をDATE型に変換
    inputValue = Me.txtInputStDate.Text '出力値を変数に代入
    
    '基準日入力値の有無で条件分岐
    If inputValue = "" Then
        MsgBox "基準日が入力されていません。", Buttons:=vbExclamation
    Else
        On Error GoTo ErrorLbl
        convertdate = CDate(inputValue) '出力値をDate型に変換
        
        '入力値が遡及年度内であるかの条件分岐
        If initDate <= convertdate And convertdate <= finDate Then
             result = MsgBox("第1期納付期限を登録してよろしいですか?" & vbCrLf & "登録年月日:" & CStr(convertdate), Buttons:=vbYesNo)
             
            'メッセージボックスの結果で条件分岐
            If result = vbNo Then
                firstDeadline = "0:00:00"
                MsgBox "登録を取り消しました。", Buttons:=vbInformation
            Else
                firstDeadline = convertdate + 1 'グローバル変数に出力値+1を代入
                goBackAbleDate = DateAdd("yyyy", 2, firstDeadline) '遡及可能年月日の取得
                MsgBox "登録しました。" & vbCrLf & "登録年月日:" & CStr(convertdate), Buttons:=vbInformation
                Me.txtInputStDate.Text = "" '入力値を初期化
                Unload Me
                Unload UserForm1
                UserForm1.Show
            End If
        Else
            MsgBox "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度の範囲外です。" & vbCrLf & objectYear & "年４月１日～" & objectYear + 1 & "年３月31日の間で入力してください。", Buttons:=vbExclamation
            
        End If
    End If
        Exit Sub
        
ErrorLbl:
        MsgBox "入力値が不正です。", Buttons:=vbExclamation
        Me.txtInputStDate.Text = ""
    
End Sub


Private Sub UserForm_initialize()
    Application.Visible = False
    today = Date '本日の年月日を取得
    ConvertToday = Format(today, "yyyy年mm月dd日") '本日の年月日を文字列に変換
    thisYear = fiscalYear(today) '現年度の取得(int) 【fiscalYearメソッド】は、標準モジュール内で定義している。
    convertToWareki = thisYear - 2018 '現年度を和暦(令和に変換)
    objectYear = thisYear - 2 '遡及年度の取得(int)
    initDate = CDate(objectYear & "/" & "04" & "/" & "01") '遡及年度の開始日を取得({遡及年度}年4月1日)
    finDate = CDate(objectYear + 1 & "/" & "03" & "/" & "31") '遡及年度の終了日の取得({遡及年度+1}年 3月31日)
    
    Me.lblQdisplayStDate = "現年度:" & "令和" & convertToWareki & "(" & thisYear & ")" & "年度" & vbCrLf & vbCrLf & "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度の第1期の納付期限を入力してください。"
End Sub
