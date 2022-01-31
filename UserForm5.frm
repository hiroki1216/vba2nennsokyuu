VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "“¯¢‘Ñ‰Á“üÒ‘Šiæ“¾“úŠm”F"
   ClientHeight    =   8565.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'ƒI[ƒi[ ƒtƒH[ƒ€‚Ì’†‰›
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
    Dim convertdate As Date 'o—Í’l‚ğDateŒ^‚É•ÏŠ·(ƒ[ƒJƒ‹•Ï”)
    Dim standardDate As Date '‹NZ“ú‚Ìæ“¾
    Dim result As VbMsgBoxResult 'ƒƒbƒZ[ƒWƒ{ƒbƒNƒX‚ÌŒ‹‰Ê
    Dim convertFirstDeadline As String '‘k‹y”N“x‚Ì‘æ1Šú”[•tŠúŒÀ(•¶š—ñ)
    
    inputValue = Me.txtInputYear2.Text 'o—Í’l‚ğ•Ï”‚É‘ã“ü
    On Error GoTo ErrorLbl
    convertdate = CDate(inputValue) 'o—Í’l‚ğDateŒ^‚É•ÏŠ·(ƒ[ƒJƒ‹•Ï”)
    standardDate = convertdate + 1 '‘k‹y‹NZ“ú‚ğæ“¾(ƒ[ƒJƒ‹•Ï”)
    convertFirstDeadline = Format(firstDeadline + 1, "yyyy”NmŒd“ú") '‘k‹y”N“x‚Ì‘æ1Šú”[•tŠúŒÀ‚ğ•¶š—ñ‚É•ÏŠ·
    
    '‘k‹y”N“x‚Ì‘æ‚PŠú”[•tŠúŒÀ`‘k‹y”N“x‚ÌI—¹“ú‚Å“ü—Í‚³‚¹‚é‚½‚ß‚ÌğŒ•ªŠò
    If convertdate > firstDeadline And convertdate <= finDate Then
        result = MsgBox("w‘–¯Œ’N•ÛŒ¯‘Šiæ“¾“úx‚ğ“o˜^‚µ‚Ä‚æ‚ë‚µ‚¢‚Å‚·‚©?" & vbCrLf & "“o˜^”NŒ“ú:" & CStr(convertdate), Buttons:=vbYesNo)  'MsgBox‚Ì–ß‚è’l•Ï”‚É‘ã“ü
        'MsgBox‚Ì–ß‚è’l‚ÅğŒ•ªŠò
        If result = vbNo Then
            MsgBox "“o˜^‚ğæ‚èÁ‚µ‚Ü‚µ‚½B"
            Exit Sub
        Else
            MsgBox "“o˜^‚µ‚Ü‚µ‚½B" & vbCrLf & "“o˜^”NŒ“ú:" & CStr(convertdate), Buttons:=vbInformation
        End If
        
        '‘k‹y”N“x‚Ì4Œ`6Œ‚ÍAgoBackAbleDate=‘k‹y”N“x‚Ì‘æˆêŠú”[•tŠúŒÀ
        If convertdate < firstDeadline Then
            goBackAbleDateComparison = DateAdd("yyyy", 2, firstDeadline) '‘k‹y‰Â”\”NŒ“ú(”äŠr—p)‚Ìæ“¾
        Else
            goBackAbleDateComparison = DateAdd("yyyy", 2, standardDate) '‘k‹y‰Â”\”NŒ“ú(”äŠr—p)‚Ìæ“¾
        End If
        
        Me.Hide
        UserForm2.Show
    Else
         MsgBox "”ÍˆÍŠO‚Å‚·B" & vbCrLf & convertFirstDeadline & "`" & objectYear + 1 & "”N‚RŒ31“ú‚Å“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", Buttons:=vbExclamation
    End If
        Exit Sub
ErrorLbl:
        MsgBox "“ü—Í’l‚ª•s³‚Å‚·B"
        Me.txtInputYear2.Text = ""
    
End Sub


Private Sub UserForm_initialize()
Application.Visible = False
Me.lblQD2 = "“¯¢‘Ñ‚Ì‰Á“üÒ‚Ìw‘–¯Œ’N•ÛŒ¯‘Šiæ“¾“úx‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B" & vbCrLf & vbCrLf & "–{“ú‚Í" & ConvertToday & "‚Å‚·B" & vbCrLf & vbCrLf & "¦‰Á“üÒ‚ª•¡”‚¢‚éê‡‚ÍA‚»‚Ì’†‚Å1”Ô‰‚ß‚É‘Ši‚ğæ“¾‚µ‚½Ò‚Ìæ“¾“ú‚ğ‹L“ü‚µ‚Ä‚­‚¾‚³‚¢B"
End Sub
