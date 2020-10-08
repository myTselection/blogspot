VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'myT 04/2015, http://myTselection.blogspot.com
'Instructions:
'Open Outlook and press [ALT+F11] to open the VBS developer editor
'Navigate to 'Project 1 VbaProject.OTM' > Microsoft Outlook Objects > ThisOutlookSession
'Copy and Paste the macro below and close the VBS editor

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
Dim m As Variant, t As Variant
Dim strBody As String
Dim intIn As Long
Dim intAttachCount As Integer, intStandardAttachCount As Integer, limitBody As Integer

On Error GoTo handleError

'Edit the following line if you have a signature on your email that includes images or other files. Make intStandardAttachCount equal the number of files in your signature.
intStandardAttachCount = 0

'If CheckProperty(Item, "HTMLBody") Then
'    strBody = LCase(Item.HTMLBody) & " " & LCase(Item.Subject)
'Else
'    strBody = LCase(Item.Body) & " " & LCase(Item.Subject)
'End If

On Error GoTo NoHtml
strBody = LCase(Item.HTMLBody) & " " & LCase(Item.Subject)
'Show Debug via View > Immediate Window (Ctrl+G)
'Debug.Print strBody

On Error GoTo handleError
'detect line break where previous messages (reply/forward) are starting, only scan the newly create message ignoring forwarded/replied text
'HTML original message splitter
limitBody = InStr(1, strBody, "<div style='border:none;border-top:solid #b5c4df 1.0pt;padding:3.0pt 0cm 0cm 0cm'>")
'RichText has HTMLBody but no border separator
If limitBody = 0 Then GoTo RichText
GoTo commonHandling

RichText:
    On Error GoTo handleError
    'could test on availability of 'text/rtf format' in body to make sure this is rich text case
    strBody = LCase(Item.Body) & " " & LCase(Item.Subject)
    'rich text original message splitter
    limitBody = InStr(1, strBody, "_____________________________________________")
GoTo commonHandling

NoHtml:
    On Error GoTo handleError
    strBody = LCase(Item.Body) & " " & LCase(Item.Subject)
    'detect line break where previous messages (reply/forward) are starting, only scan the newly create message ignoring forwarded/replied text
    'Plain text message splitter
    limitBody = InStr(1, strBody, "-----")
GoTo commonHandling

commonHandling:
    On Error GoTo handleError
    'Show Debug via View > Immediate Window (Ctrl+G)
    'Debug.Print strBody
    
    If limitBody = 0 Then limitBody = Len(strBody)
    
    
    intIn = InStr(1, Left(strBody, limitBody), "attach")
    intIn = intIn + InStr(1, Left(strBody, limitBody), "attached")
    intIn = intIn + InStr(1, Left(strBody, limitBody), "bijlage")
    intIn = intIn + InStr(1, Left(strBody, limitBody), "bijgevoegd")
    'Copy previous line to add extra keywords
    
    intAttachCount = Item.Attachments.Count

    
    If (intIn <> 0 And intAttachCount = intStandardAttachCount) Then
        
        m = MsgBox("It appears that you mean to send an attachment, " & vbCrLf & "but there is no attachment to this message." & vbCrLf & vbCrLf & "Do you still want to send without attachments?", vbQuestion + vbYesNo + vbMsgBoxSetForeground, "Missing attachment")
    
        If m = vbNo Then
            Cancel = True
        End If
    End If

handleError:

    If (Err.Number <> 0) Then
        t = MsgBox("Outlook Attachment Reminder Error " & Err.Description & ", " & Err.Number & ", " & Err.Source, vbExclamation, "Outlook Attachment Reminder Error")
    End If

End Sub
