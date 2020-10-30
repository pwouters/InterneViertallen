Attribute VB_Name = "modOutlook"
Function OutlookSendMail(sTo As Variant, sCC As Variant, sBCC As Variant, sSubject As Variant, sBody As Variant, Optional sAttachments As Variant = "", Optional sAttachmentNames As Variant = "")
    Dim oOutlookApp As Outlook.Application       'Outlook.Application
    Dim olMail As Outlook.MailItem           'Outlook.MailItem
    Dim olRecip As Outlook.Recipient
 
    Dim asAttachments() As String, asAttachmentNames() As String
    Dim lThisAttachment As Long, lLastAttachName As Long
    Dim asRecipients() As String, lThisRecip As Long
    
     On Error Resume Next

    '' Set oOutlookApp = GetObject(, "Outlook.Application")
    
   '' If Err <> 0 Then
    'Outlook wasn't running, start it from code
     Set oOutlookApp = CreateObject("Outlook.Application")
   '' End If

    
    On Error GoTo ErrFailed
    If (oOutlookApp Is Nothing) = False Then
        'Create a new mail item
        Set olMail = oOutlookApp.CreateItem(0)      'olMailItem
        olMail.SentOnBehalfOfName = mailclient
         'Set the mail fields
        olMail.Subject = sSubject
        'Add the list of recipients
        sTo = Replace(sTo, ";", ",")
        asRecipients = Split(sTo, ",")
        For lThisRecip = 0 To UBound(asRecipients)
            'olMail.Recipients.Add Trim$(asRecipients(lThisRecip))
            Set oRecip = olMail.Recipients.Add(Trim$(asRecipients(lThisRecip)))
            oRecip.Type = olTo
        Next
       sCC = Replace(sCC, ";", ",")
        asRecipients = Split(sCC, ",")
        For lThisRecip = 0 To UBound(asRecipients)
            'olMail.Recipients.Add Trim$(asRecipients(lThisRecip))
            Set oRecip = olMail.Recipients.Add(Trim$(asRecipients(lThisRecip)))
            oRecip.Type = olCC
        Next
       sBCC = Replace(sBCC, ";", ",")
       asRecipients = Split(sBCC, ",")
        For lThisRecip = 0 To UBound(asRecipients)
            'olMail.Recipients.Add Trim$(asRecipients(lThisRecip))
            Set oRecip = olMail.Recipients.Add(Trim$(asRecipients(lThisRecip)))
            oRecip.Type = olBCC
        Next
        olMail.HTMLBody = "<pre>" & sBody & "</pre>"
      
       
        If Len(sAttachments) > 0 Then
            'Add attachments
            On Error Resume Next
                    
            sAttachments = Replace(sAttachments, ";", ",")
            asAttachments = Split(sAttachments, ",")
            If Len(sAttachmentNames) Then
                sAttachmentNames = Replace(sAttachmentNames, ";", ",")
                asAttachmentNames = Split(sAttachmentNames, ",")
                lLastAttachName = UBound(asAttachmentNames)
            Else
                lLastAttachName = -1
            End If
            For lThisAttachment = 0 To UBound(asAttachments)
                'Check the attachment exists
                asAttachments(lThisAttachment) = Trim$(asAttachments(lThisAttachment))
                If Len(Dir$(asAttachments(lThisAttachment))) > 0 Then
                    'Attachment exists, add it
                    With olMail.Attachments.Add(asAttachments(lThisAttachment), 1)       'Where 1 = olByValue (Embed attachment in the item)
                        If lThisAttachment <= lLastAttachName Then
                            .DisplayName = asAttachmentNames(lThisAttachment)
                        End If
                    End With
                End If
            Next
            On Error GoTo ErrFailed
        End If
        
        
           'Send Mail
            olMail.Send
         
        'Clear object pointers
        Set olMail = Nothing
        Set oOutlookApp = Nothing
        OutlookSendMail = True
    Else
        'Failed to create an outlook
        OutlookSendMail = False
    End If
    Exit Function

ErrFailed:
    'Failed to send mail
    Debug.Print "Error in OutlookSendMail: " & Err.Description
    If (olMail Is Nothing) = False Then
        olMail.Delete
    End If
    Set olMail = Nothing
    Set oOutlookApp = Nothing
    OutlookSendMail = False
End Function

