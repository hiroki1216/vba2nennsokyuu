VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "�k�y�N�Z���o�^"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBackToTop_Click()
    '����o�^�̗L���ŏ�������
    If firstDeadline = "0:00:00 " Then
        MsgBox "�����o�^���Ă��������B", Buttons:=vbExclamation
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
    Dim inputValue As String '���͒l
    Dim result As VbMsgBoxResult '���b�Z�[�W�{�b�N�X�̌���
    Dim convertdate As Date '���͒l��DATE�^�ɕϊ�
    inputValue = Me.txtInputStDate.Text '�o�͒l��ϐ��ɑ��
    
    '������͒l�̗L���ŏ�������
    If inputValue = "" Then
        MsgBox "��������͂���Ă��܂���B", Buttons:=vbExclamation
    Else
        On Error GoTo ErrorLbl
        convertdate = CDate(inputValue) '�o�͒l��Date�^�ɕϊ�
        
        '���͒l���k�y�N�x���ł��邩�̏�������
        If initDate <= convertdate And convertdate <= finDate Then
             result = MsgBox("��1���[�t������o�^���Ă�낵���ł���?" & vbCrLf & "�o�^�N����:" & CStr(convertdate), Buttons:=vbYesNo)
             
            '���b�Z�[�W�{�b�N�X�̌��ʂŏ�������
            If result = vbNo Then
                firstDeadline = "0:00:00"
                MsgBox "�o�^���������܂����B", Buttons:=vbInformation
            Else
                firstDeadline = convertdate + 1 '�O���[�o���ϐ��ɏo�͒l+1����
                goBackAbleDate = DateAdd("yyyy", 2, firstDeadline) '�k�y�\�N�����̎擾
                MsgBox "�o�^���܂����B" & vbCrLf & "�o�^�N����:" & CStr(convertdate), Buttons:=vbInformation
                Me.txtInputStDate.Text = "" '���͒l��������
                Unload Me
                Unload UserForm1
                UserForm1.Show
            End If
        Else
            MsgBox "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x�͈̔͊O�ł��B" & vbCrLf & objectYear & "�N�S���P���`" & objectYear + 1 & "�N�R��31���̊Ԃœ��͂��Ă��������B", Buttons:=vbExclamation
            
        End If
    End If
        Exit Sub
        
ErrorLbl:
        MsgBox "���͒l���s���ł��B", Buttons:=vbExclamation
        Me.txtInputStDate.Text = ""
    
End Sub


Private Sub UserForm_initialize()
    Application.Visible = False
    today = Date '�{���̔N�������擾
    ConvertToday = Format(today, "yyyy�Nmm��dd��") '�{���̔N�����𕶎���ɕϊ�
    thisYear = fiscalYear(today) '���N�x�̎擾(int) �yfiscalYear���\�b�h�z�́A�W�����W���[�����Œ�`���Ă���B
    convertToWareki = thisYear - 2018 '���N�x��a��(�ߘa�ɕϊ�)
    objectYear = thisYear - 2 '�k�y�N�x�̎擾(int)
    initDate = CDate(objectYear & "/" & "04" & "/" & "01") '�k�y�N�x�̊J�n�����擾({�k�y�N�x}�N4��1��)
    finDate = CDate(objectYear + 1 & "/" & "03" & "/" & "31") '�k�y�N�x�̏I�����̎擾({�k�y�N�x+1}�N 3��31��)
    
    Me.lblQdisplayStDate = "���N�x:" & "�ߘa" & convertToWareki & "(" & thisYear & ")" & "�N�x" & vbCrLf & vbCrLf & "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x�̑�1���̔[�t��������͂��Ă��������B"
End Sub
