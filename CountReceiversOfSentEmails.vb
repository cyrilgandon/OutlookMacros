' Display a Message Box with How many people you annoy
Public Sub SentFolderStats()
  
    Dim TotalRecipients As Long
    Dim LastMonthSentEmailCount As Long
    Dim objSentFolder As Outlook.MAPIFolder
    Dim objNS As Outlook.NameSpace
    'Dim objItem As Outlook.MailItem
    
    Dim myAddrList As AddressList
    Dim myAddrEntries As AddressEntry

    Set objNS = Application.GetNamespace("MAPI")
    Set objSentFolder = objNS.GetDefaultFolder(olFolderSentMail)
    Set myAddrList = objNS.AddressLists("Liste d'adresses globale")
    
    TotalRecipients = 0
    LastMonthSentEmailCount = 0
    ExtraDistListMembers = 0
    
    For Each objItem In objSentFolder.Items
       If objItem.Class = olMail Then
         If Now() - objItem.SentOn < 30 Then 'est-ce un mail du mois dernier?
            For Each objRecipient In objItem.Recipients
                
                If objRecipient.DisplayType = olDistList Then 'est-ce une liste de distribution?
                    Set myAddrEntries = myAddrList.AddressEntries(objRecipient.Name)
                    ExtraDistListMembers = ExtraDistListMembers + myAddrEntries.Members.Count - 1 '-1 parce que la liste est comptée comme un recipient
                End If
            Next
             TotalRecipients = TotalRecipients + objItem.Recipients.Count + ExtraDistListMembers
             LastMonthSentEmailCount = LastMonthSentEmailCount + 1
             ExtraDistListMembers = 0
         End If
       End If
    Next

MsgBox "Vous avez envoyé " & Str(LastMonthSentEmailCount) & " mails le mois dernier. Ils ont été distribués à " & Str(TotalRecipients) & " destinataires!", vbOKOnly, "Statistiques de messages envoyés"

End Sub
