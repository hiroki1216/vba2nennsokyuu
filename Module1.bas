Attribute VB_Name = "Module1"
'�O���[�o���ϐ��̒�`

Public optbtnOkResult As Boolean '�I�v�V�����{�^���̐^�U�l(�͂�)
Public optbtnNoResult As Boolean '�I�v�V�����{�^���̐^�U�l(������ �����щ����҂Ȃ�)
Public optbtnNotResult As Boolean '�I�v�V�����{�^���̐^�U�l(������ ������_�����҂���)

Public ConvertToday As String '�����̓��t(������)
Public today As Date '�����̓��t
Public thisYear As Long '���N�x
Public objectYear As Long '�k�y�N�x
Public convertToWareki As Long '�a��ϊ��p

Public goBackAbleDate As Date '�k�y�\�N����
Public goBackAbleDateComparison As Date '�k�y�\�N����(��r�p)

Public firstDeadline As Date '�k�y�N�x�̑����[�t����

Public initDate As Date '�k�y�N�x�̊J�n��
Public finDate As Date '�k�y�N�x�̏I����


'�{���̔N��������N�x�����߂�֐�
Function fiscalYear(ByVal today As Date) As Long
    If Month(today) >= 4 Then
        fiscalYear = Year(today)
    Else
        fiscalYear = Year(today) - 1
    End If
End Function



