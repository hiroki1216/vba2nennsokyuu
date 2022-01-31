VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "同世帯国保加入者確認"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnNext_Click()
    optbtnOkResult = optbtnOk.Value
    optbtnNoResult = optbtnNo.Value
    optbtnNotResult = optbtnNot.Value
    
    If optbtnOkResult = False And optbtnNoResult = False And optbtnNotResult = False Then
        MsgBox "回答が未選択です。", Buttons:=vbExclamation
    Else
        'オプションボタンの出力結果で条件分岐
        If optbtnOkResult = True Then
            Me.Hide
            UserForm4.Show
        ElseIf optbtnNoResult = True Then
            Me.Hide
            UserForm2.Show
        Else
            Me.Hide
            UserForm5.Show
        End If
    End If
End Sub

Private Sub btnRegistrateDate_Click()
    Me.Hide
    UserForm3.Show
End Sub


Private Sub UserForm_initialize()
    Application.Visible = False
    Me.lblDisplayStDate = "基準日：" & vbCrLf & "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度" & vbCrLf & "第１期納付期限の翌日:" + CStr(firstDeadline)
    Me.lblQD = "現年度:" & "令和" & convertToWareki & "(" & CStr(thisYear) & ")" & "年度" & vbCrLf & vbCrLf & "令和" & convertToWareki - 2 & "(" & objectYear & ")" & "年度について、上記の基準日時点で世帯に国保加入者はいますか?"
End Sub
