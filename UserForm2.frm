VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "2�N�k�y����"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnBack_Click()
    Unload UserForm2 'UserForm2���\��
    UserForm1.Show 'UserForm1��\��
End Sub

Private Sub btnJudgement_Click()
    Dim inputValue As String
    Dim convertdate As Date '�o�͒l��Date�^�ɕϊ�(���[�J���ϐ�)
    Dim standardDate As Date '�N�Z���̎擾
    Dim lastDate As Date '�k�y�N�x6������
    Dim result As VbMsgBoxResult '���b�Z�[�W�{�b�N�X�̌���
    
    lastDate = CDate(objectYear & "/" & "06" & "/" & "30") '�������DATE�^�ɕϊ�
    inputValue = Me.txtYearInput.Text '�o�͒l��ϐ��ɑ��
    On Error GoTo ErrorLbl
    convertdate = CDate(inputValue) '�o�͒l��Date�^�ɕϊ�(���[�J���ϐ�)
    standardDate = convertdate + 1 '�k�y�N�Z�����擾(���[�J���ϐ�)
    
    '���͒l���k�y�N�x�̏I�����������͂����Ȃ����߂̏�������
    If convertdate <= finDate Then
        result = MsgBox("�w�������N�ی����i�擾���x��o�^���Ă�낵���ł���?" & vbCrLf & "�o�^�N����:" & CStr(convertdate), Buttons:=vbYesNo)  'MsgBox�̖߂�l�ϐ��ɑ��
        'MsgBox�̖߂�l�ŏ�������
        If result = vbNo Then
            MsgBox "�o�^���������܂����B"
            Exit Sub
        Else
            MsgBox "�o�^���܂����B" & vbCrLf & "�o�^�N����:" & CStr(convertdate), Buttons:=vbInformation
        End If
        
        '�k�y�N�x��4���`�����[�t�����́AgoBackAbleDate=�k�y�N�x�̑����[�t����
        If convertdate < firstDeadline Then
            goBackAbleDate = DateAdd("yyyy", 2, firstDeadline) '�k�y�\�N�����̎擾
        Else
            goBackAbleDate = DateAdd("yyyy", 2, standardDate) '�k�y�\�N�����̎擾
        End If
        
        '������(�����̉����҂�����)���I�����ꂽ�ꍇ�̋N�Z���̔�r���邩�̏�������
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
         MsgBox "�ߘa" & convertToWareki - 2 & "(" & objectYear & ")" & "�N�x�𒴉߂��Ă��܂��B" & vbCrLf & "�`" & "�ߘa" & convertToWareki - 1 & "(" & objectYear + 1 & ")" & "�N�R��31���œ��͂��Ă��������B", Buttons:=vbExclamation
    End If
    
    Exit Sub
ErrorLbl:
        MsgBox "���͒l���s���ł��B"
        Me.txtYearInput.Text = ""
End Sub

Private Sub UserForm_initialize()
    Application.Visible = False
    If optbtnOkResult = True Then
        Me.lblQDisplay2 = "�����т̉����҂́w�������N�ی����i�擾���x����͂��Ă��������B" & vbCrLf & vbCrLf & "�{����" & ConvertToday & "�ł��B"
    Else
        Me.lblQDisplay2 = "�Ώێ҂́w�������N�ی����i�擾���x����͂��Ă��������B" & vbCrLf & vbCrLf & "�{����" & ConvertToday & "�ł��B"
    End If
End Sub
