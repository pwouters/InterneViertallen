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


Public Function grap_table(sTable As Variant) As String
Dim teller_1, teller_2 As Integer
Dim s As String

s = sTable
s = Replace(s, Chr(10), "")

If s = "" Then
    grap_table = ""
        Exit Function
End If

teller_1 = InStr(s, "<tbody")

If teller_1 = 0 Then
    grap_table = ""
        Exit Function
End If

s = Mid(s, teller_1 + 6)
teller_1 = InStr(s, ">")
s = Mid(s, teller_1 + 1)
s = Trim(s)
'If Left(s, 1) = Chr(10) Then
'    s = Mid(s, 2)
'    s = Trim(s)
'End If


teller_1 = InStr(s, "</tbody>")
s = Left(s, teller_1 - 1)
grap_table = s

End Function

Public Function grap_rijen(sRijen As Variant) As Variant
Dim Rijen() As String
Dim i As Integer
Dim teller As Integer
Rijen = Split(sRijen, "</tr>")
For i = LBound(Rijen) To UBound(Rijen)
    Rijen(i) = Trim(Rijen(i))
    teller = InStr(Rijen(i), ">")
    Rijen(i) = Mid(Rijen(i), teller + 1)
Next
grap_rijen = Rijen
End Function

Public Function grap_cellen(sCellen As Variant) As Variant
Dim Cellen() As String
Dim i As Integer
Dim teller As Integer
Cellen = Split(sCellen, "</td>")
For i = LBound(Cellen) To UBound(Cellen)
    Cellen(i) = Replace(Cellen(i), Chr(34), "")
    Cellen(i) = Trim(Cellen(i))
    teller = InStr(Cellen(i), ">")
    Cellen(i) = Trim(Mid(Cellen(i), teller + 1))
Next
grap_cellen = Cellen
End Function


Public Function extract_kleur(strKleur As Variant) As String
Dim s As String
s = strKleur
s = Replace(s, "<img src=" & Chr(34) & "/images/suit411.gif" & Chr(34) & " alt=" & Chr(34) & "S" & Chr(34) & ">", "S")
s = Replace(s, "<img src=" & Chr(34) & "/images/suit311.gif" & Chr(34) & " alt=" & Chr(34) & "H" & Chr(34) & ">", "H")
s = Replace(s, "<img src=" & Chr(34) & "/images/suit211.gif" & Chr(34) & " alt=" & Chr(34) & "D" & Chr(34) & ">", "D")
s = Replace(s, "<img src=" & Chr(34) & "/images/suit111.gif" & Chr(34) & " alt=" & Chr(34) & "C" & Chr(34) & ">", "C")
s = Replace(s, "<img src=" & Chr(34) & "/images/suit409.gif" & Chr(34) & " alt=" & Chr(34) & "S" & Chr(34) & ">", "S")
s = Replace(s, "<img src=" & Chr(34) & "/images/suit309.gif" & Chr(34) & " alt=" & Chr(34) & "H" & Chr(34) & ">", "H")
s = Replace(s, "<img src=" & Chr(34) & "/images/suit209.gif" & Chr(34) & " alt=" & Chr(34) & "D" & Chr(34) & ">", "D")
s = Replace(s, "<img src=" & Chr(34) & "/images/suit109.gif" & Chr(34) & " alt=" & Chr(34) & "C" & Chr(34) & ">", "C")
extract_kleur = s
End Function