Attribute VB_Name = "modWeb"
Option Compare Text

Global MACAddress As String
Global Release_date As Variant
Global Personid As Long

Public Enum ePort
   INTERNET_DEFAULT_HTTP_PORT = 80
   INTERNET_DEFAULT_HTTPS_PORT = 443
End Enum

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_SERVICE_HTTP = 3

Private Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_SECURE = &H800000
Private Const INTERNET_FLAG_FROM_CACHE = &H1000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Const BUFFER_LENGTH As Long = 1024
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_PROXY = 3

Public Const PROJECTDATA As String = "http://prod.sparc.sc.philips.com:4080/niku/app?action=nmc.executeAvailReport&hidden_compusers_id=&hidden_sharegroups_id=&job=&hidden_compgroups_id=&compUsers=&hidden_failgroups_id=&header=&when_description=Run%20Once&failUsers=&reportId=5000803&hidden_shareusers_id=&hidden_failusers_id="

Public Const STEPDATA As String = "http://admin.stepbridge.nl/show.php?page=tournamentinfo&activityid="


Public Const scUserAgent = "VB OpenUrl"

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
(ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, _
ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

Private Declare Function InternetConnect Lib "wininet.dll" Alias _
"InternetConnectA" (ByVal hInternetSession As Long, ByVal ServerName As String, _
ByVal ServerPort As Integer, ByVal UserName As String, ByVal Password As _
String, ByVal Service As Long, ByVal Flags As Long, ByVal Context As Long) As _
Long


Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias _
"HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal Verb As String, ByVal _
ObjectName As String, ByVal Version As String, ByVal Referer As String, ByVal _
AcceptTypes As Long, ByVal Flags As Long, Context As Long) As Long

Private Declare Function HttpSendRequest Lib "wininet.dll" Alias _
"HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal Headers As String, ByVal _
HeadersLength As Long, ByVal sOptional As String, ByVal OptionalLength As Long) _
As Boolean

Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


Private hHTTP As Long
Private hConnection As Long

Private Const FIELDS_BUFFER_LENGTH As Long = 10
Private Const FIELDS_NAME_INDEX As Long = 0
Private Const FIELDS_VALUE_INDEX As Long = 1

Private DontEncode(255) As Boolean

Private FieldCount As Long
Private mFields() As String



'Then, here's a function that returns the entire text of the specified URL:
Public Function GetHTMLFromURL(sUrl As Variant) As String

Dim s As String
Dim hOpen As Long
Dim hOpenUrl As Long
Dim bDoLoop As Boolean
Dim bRet As Boolean
Dim sReadBuffer As String * 2048
Dim lNumberOfBytesRead As Long

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

bDoLoop = True
While bDoLoop
    sReadBuffer = vbNullString
    bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
    s = s & Left$(sReadBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
Wend

If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
If hOpen <> 0 Then InternetCloseHandle (hOpen)

GetHTMLFromURL = s

End Function

Public Sub GetHTMLFromURLandSaveAs(sUrl As Variant, sPath As Variant)
     ' URL string
Dim intFile As Integer   ' FreeFile variable
intFile = FreeFile()
Dim s As String
Dim hOpen As Long
Dim hOpenUrl As Long
Dim sLenght As Long

Dim bDoLoop As Boolean
Dim bRet As Boolean
Dim sReadBuffer As String * 2048
Dim lNumberOfBytesRead As Long

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

bDoLoop = True
While bDoLoop
    sReadBuffer = vbNullString
    bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
    s = s & Left$(sReadBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
Wend

If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
If hOpen <> 0 Then InternetCloseHandle (hOpen)

Open sPath For Output As #intFile
Print #intFile, s
Close #intFile


End Sub



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
