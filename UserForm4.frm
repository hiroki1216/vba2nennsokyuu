VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "���茋��"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFin_Click()
    goBackAbleDate = DateAdd("yyyy", 2, firstDeadline) '�k�y�\�N�����̏�����
    'UserForm�̏�����
    Unload Me
    Unload UserForm1
    Unload UserForm2
    Unload UserForm5
    'UserForm1��\��
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
    Dim strGoBackAbleDate As String '�k�y�\�N����(������^)
    strGoBackAbleDate = Format(goBackAbleDate, "yyyy�Nmm��dd��") '�k�y�\�N�����𕶎���^�ɕϊ�
        
    '�k�y�ۂ̔���
    If goBackAbleDate >= today Then
        Me.lblDisplayResult2 = "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x�́A" & "�w�k�y�x�ł��B" & vbCrLf & vbCrLf & "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x���畊�ۍX�����ĉ������B"
    Else
        Me.lblDisplayResult2 = "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x�́A" & "�w�k�y�s�x�ł��B" & vbCrLf & vbCrLf & "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x�ȑO�́A���ۍX�����Ȃ��ŉ������B"
    End If
    
    Me.lblDisplayResult = "�{���̓��t:" & ConvertToday & vbCrLf & vbCrLf & "�k�y�\�͏o��:" + strGoBackAbleDate
End Sub
