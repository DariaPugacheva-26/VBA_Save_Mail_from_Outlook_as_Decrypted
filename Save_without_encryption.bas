Attribute VB_Name = "Save_without_encryption"
Sub save_decrypted()
Dim Path As String 'Path for saving email
Dim SLCT As Outlook.Selection 'Select one or few mails to save
Dim Mails As Outlook.MailItem 
Dim MyForward As Outlook.MailItem 'Making forward from selected mails
Dim MailSubj As String 'Filename will be as Mail subject

Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003" 
'Decryption flag. More info: https://docs.microsoft.com/ru-ru/archive/blogs/dvespa/how-to-sign-or-encrypt-a-message-programmatically-from-oom
Dim oProp As Long
Dim D As String
Dim junk() As Variant, i As Variant

'FileNames at Windows can not have any of followed characters, so we are replacing them them to spaces
junk = Array("\", "|", "/", ":", "?", "<", ">", "+", "*", Chr(34)) 'chr(34) is "

D = Format(Date, "dd.mm.yy") 'date for filename

Set SLCT = Application.ActiveExplorer.Selection
For Each Mails In SLCT
    Set MyForward = Mails.Forward
    oProp = CLng(MyForward.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS))
    uFlags = 0
    ulFlags = ulFlags Or &H2
        With MyForward
        MailSubj = .Subject
            For Each i In junk
            MailSubj = Replace(MailSubj, i, " ") 'replacing forbidden characters
            Next i
        MailSubj = Trim(Right(MailSubj, Len(MailSubj) - 3)) 'removing FW and space from mail subject
        Path = "C:\UserName\Saved_Mails" & "\" & MailSubj & " from " & D & ".msg" 'Insert needed path
        .PropertyAccessor.SetProperty PR_SECURITY_FLAGS, ulFlags
        .SaveAs Path
        End With

Next Mails

End Sub
