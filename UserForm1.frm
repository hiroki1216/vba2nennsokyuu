VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�����э��ۉ����Ҋm�F"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
        MsgBox "�񓚂����I���ł��B", Buttons:=vbExclamation
    Else
        '�I�v�V�����{�^���̏o�͌��ʂŏ�������
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
    Me.lblDisplayStDate = "����F" & vbCrLf & "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x" & vbCrLf & "��P���[�t�����̗���:" + CStr(firstDeadline)
    Me.lblQD = "���N�x:" & "�ߘa" & convertToWareki & "(" & CStr(thisYear) & ")" & "�N�x" & vbCrLf & vbCrLf & "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x�ɂ��āA��L�̊�����_�Ő��тɍ��ۉ����҂͂��܂���?"
End Sub
