VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "判定結果"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFin_Click()
    goBackAbleDate = DateAdd("yyyy", 2, firstDeadline) '遡及可能年月日の初期化
    'UserFormの初期化
    Unload Me
    Unload UserForm1
    Unload UserForm2
    Unload UserForm5
    'UserForm1を表示
    UserForm1.Show
End Sub

Private Sub BtnGoback2_Click()
    Unload Me
    If TypeOf UserForm Is UserForm2 Then
        UserForm2.Show
    Else
        UserForm1.Show
    End If
End Sub

Private Sub UserForm_initialize()
    Application.Visible = False
    Dim strGoBackAbleDate As String '遡及可能年月日(文字列型)
    strGoBackAbleDate = Format(goBackAbleDate, "yyyy年mm月dd日") '遡及可能年月日を文字列型に変換
        
    '遡及可否の判定
    If goBackAbleDate >= today Then
        Me.lblDisplayResult2 = "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度は、" & "『遡及可』です。" & vbCrLf & vbCrLf & "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度から賦課更正して下さい。"
    Else
        Me.lblDisplayResult2 = "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度は、" & "『遡及不可』です。" & vbCrLf & vbCrLf & "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度以前は、賦課更正しないで下さい。"
    End If
    
    Me.lblDisplayResult = "本日の日付:" & ConvertToday & vbCrLf & vbCrLf & "遡及可能届出日:" + strGoBackAbleDate
End Sub
