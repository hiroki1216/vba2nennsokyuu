VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "�����щ����Ҏ��i�擾���m�F"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
    Dim convertdate As Date '�o�͒l��Date�^�ɕϊ�(���[�J���ϐ�)
    Dim standardDate As Date '�N�Z���̎擾
    Dim result As VbMsgBoxResult '���b�Z�[�W�{�b�N�X�̌���
    Dim convertFirstDeadline As String '�k�y�N�x�̑�1���[�t����(������)
    
    inputValue = Me.txtInputYear2.Text '�o�͒l��ϐ��ɑ��
    On Error GoTo ErrorLbl
    convertdate = CDate(inputValue) '�o�͒l��Date�^�ɕϊ�(���[�J���ϐ�)
    standardDate = convertdate + 1 '�k�y�N�Z�����擾(���[�J���ϐ�)
    convertFirstDeadline = Format(firstDeadline + 1, "yyyy�Nm��d��") '�k�y�N�x�̑�1���[�t�����𕶎���ɕϊ�
    
    '�k�y�N�x�̑�P���[�t�����`�k�y�N�x�̏I�����œ��͂����邽�߂̏�������
    If convertdate > firstDeadline And convertdate <= finDate Then
        result = MsgBox("�w�������N�ی����i�擾���x��o�^���Ă�낵���ł���?" & vbCrLf & "�o�^�N����:" & CStr(convertdate), Buttons:=vbYesNo)  'MsgBox�̖߂�l�ϐ��ɑ��
        'MsgBox�̖߂�l�ŏ�������
        If result = vbNo Then
            MsgBox "�o�^���������܂����B"
            Exit Sub
        Else
            MsgBox "�o�^���܂����B" & vbCrLf & "�o�^�N����:" & CStr(convertdate), Buttons:=vbInformation
        End If
        
        '�k�y�N�x��4���`6���́AgoBackAbleDate=�k�y�N�x�̑����[�t����
        If convertdate < firstDeadline Then
            goBackAbleDateComparison = DateAdd("yyyy", 2, firstDeadline) '�k�y�\�N����(��r�p)�̎擾
        Else
            goBackAbleDateComparison = DateAdd("yyyy", 2, standardDate) '�k�y�\�N����(��r�p)�̎擾
        End If
        
        Me.Hide
        UserForm2.Show
    Else
         MsgBox "�͈͊O�ł��B" & vbCrLf & convertFirstDeadline & "�`" & objectYear + 1 & "�N�R��31���œ��͂��Ă��������B", Buttons:=vbExclamation
    End If
        Exit Sub
ErrorLbl:
        MsgBox "���͒l���s���ł��B"
        Me.txtInputYear2.Text = ""
    
End Sub


Private Sub UserForm_initialize()
Application.Visible = False
Me.lblQD2 = "�����т̉����҂́w�������N�ی����i�擾���x����͂��Ă��������B" & vbCrLf & vbCrLf & "�{����" & ConvertToday & "�ł��B" & vbCrLf & vbCrLf & "�������҂���������ꍇ�́A���̒���1�ԏ��߂Ɏ��i���擾�����҂̎擾�����L�����Ă��������B"
End Sub
