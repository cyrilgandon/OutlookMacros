Public Sub MoveItIfTooMuchReceiver(item As Outlook.MailItem)

    Dim maxAllowedReceivers As Integer
    maxAllowedReceivers = 4
    
    Dim targetFolder As String
    targetFolder = "Tous"
    
    Dim countTo As Integer
    countTo = ReceiversCount(item)
    If (countTo > maxAllowedReceivers) Then
        Debug.Print "Should move this email '" & item.Subject & "', it has " & countTo & " receivers"
        Call MoveTo(item, targetFolder)
    Else
        Debug.Print "Message '" & item.Subject & " has " & countTo & " receivers, it can stay in the inbox"
    End If
End Sub

Sub MoveTo(item As Outlook.MailItem, folderName As String)
    Dim olApp As New Outlook.Application
    Dim olNameSpace As Outlook.NameSpace
    Set olNameSpace = olApp.GetNamespace("MAPI")
    Dim olDestFolder As Outlook.MAPIFolder
    Set olDestFolder = olNameSpace.GetDefaultFolder(olFolderInbox).Folders(folderName)
    Debug.Print "Move to folder  '" & olDestFolder.Name & "'"
    item.Move olDestFolder
End Sub

Function ReceiversCount(item As Outlook.MailItem) As Integer
    Dim strNames As String, i As Integer, j As Integer
    j = 1
    strNames = item.To
    ' count the number of semi-comma
    For i = 1 To Len(strNames)
        If Mid(strNames, i, 1) = ";" Then
            j = j + 1
        End If
    Next i
    ReceiversCount = j
End Function


