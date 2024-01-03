Sub SendEmail(addr As String, subj As String, body As String)
    Dim outlookApp As Object
    Dim newMail As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set newMail = outlookApp.CreateItem(0)

    With newMail
        .To = addr
        .Subject = subj
        .Body = body
        .Send
    End With
End Sub

' Call SendEmail("recipient@example.com", "Test Subject", "This is a test email.")

Sub CheckForReplyAndMove(acct As String, subj As String, addr As String, savefolder As String)
    Dim outlookApp As Object
    Dim namespace As Object
    Dim inbox As Object
    Dim destinationFolder As Object
    Dim mail As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set namespace = outlookApp.GetNamespace("MAPI")
    Set inbox = namespace.GetDefaultFolder(6) ' 6 refers to the inbox

    ' Try to get the destination folder, create it if it doesn't exist
    On Error Resume Next
    Set destinationFolder = namespace.Folders(acct).Folders(savefolder)
    If destinationFolder Is Nothing Then
        Set destinationFolder = namespace.Folders("Your Email Account").Folders.Add(savefolder)
    End If
    On Error GoTo 0

    For Each mail In inbox.Items
        If InStr(mail.Subject, subj) > 0 And mail.SenderEmailAddress = addr Then
            Debug.Print mail.Body ' Print the email body to the immediate window
            mail.Move destinationFolder
        End If
    Next mail
End Sub

' Call CheckForReplyAndMove("myaccount@example.com", "Test Subject", "sender@example.com", "Destination Folder")

