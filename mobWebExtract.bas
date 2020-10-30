Attribute VB_Name = "mobWebExtract"
Option Compare Text


'Extract uitslag

'we hebben de Url nodig + een ID, elke wedstrijd heeft een id
'http://admin.stepbridge.nl/show.php?page=tournamentinfo&activityid=20434



Public Function GetStepUitslag(sUrl As Variant, ActiveID As Variant) As String
    Dim s, a, b, c As String
    Dim Uitslag As String
    
    Dim Rang, Spelers, Score As String
    
    
    Dim teller, teller2, tel2, tel3, telr2, telr3 As Integer
    s = GetHTMLFromURL(sUrl & ActiveID)
    
    'zoek tag body
    
     teller = InStr(s, "<body>")
     If teller > 0 Then
        s = Mid(s, teller + 6)
     Else
        'geen body tag dan exit functie
        GetStepUitslag = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
      Exit Function
     End If
     
     'zoek tag <tbody> is een tabel met inhoud
     
    teller = InStr(s, "<tbody>")
    If teller = 0 Then
       'indien geen tabelinhoud dan exit function
        GetStepUitslag = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
        Exit Function
    End If
    
    'we hebben de tweede tag nodig
    
    s = Mid(s, teller + 7)
    teller = InStr(s, "<tbody>")
    If teller = 0 Then
       'indien geen tweede tabelinhoud dan exit function
        GetStepUitslag = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
        Exit Function
    End If
        
    s = Mid(s, teller + 7)
    
    'strip wat tag's en text die we niet nodig hebben
    s = StripHREF(s)
   'zoek body

    s = Replace(s, " align=" & Chr(34) & "right" & Chr(34), "")
    s = Replace(s, " align=" & Chr(34) & "left" & Chr(34), "")
    
    'do zolang er rijen zij
    
    Do While Left(s, 8) <> "</tbody>"
        'zoek tr
        b = tr_tag(s)
  
        '3 kolommen
        
        'Kolom rang
       
        Rang = Chr(34) & td_tag(b) & Chr(34)
        
        'Kolom spelers
        
        Spelers = Chr(34) & td_tag(b) & Chr(34)
 
        'Kolom score
        
        Score = Chr(34) & td_tag(b) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & Spelers & "," & Score & vbCr
       
    Loop
     
     'uitslag is twee kolommen
     
     teller = InStr(s, "<tbody>")
     'mocht er geen twee kolom zijn dan exit functie
     
     If teller = 0 Then
        GetStepUitslag = Uitslag
        Exit Function
    End If
        
        s = Mid(s, teller + 7)
    
    Do While Left(s, 8) <> "</tbody>"
        'zoek tr
        b = tr_tag(s)
  
        '3 kolommen
        
        'Kolom rang
       
        Rang = Chr(34) & td_tag(b) & Chr(34)
        
        'Kolom spelers
        
        Spelers = Chr(34) & td_tag(b) & Chr(34)
 
        'Kolom score
        
        Score = Chr(34) & td_tag(b) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & Spelers & "," & Score & vbCr
     Loop
    
    
    GetStepUitslag = Uitslag
End Function

Public Function tr_tag(Inhoud As Variant) As String
 Dim telr2, telr3 As Integer
 Dim b As String
        telr2 = InStr(Inhoud, "<tr>")
        telr3 = InStr(Inhoud, "</tr>")
        b = Mid(Inhoud, telr2 + 4, telr3 - telr2 + 1)
        Inhoud = Mid(Inhoud, telr3 + 6)
        tr_tag = b
End Function


Public Function td_tag(Inhoud As Variant) As String
 Dim tel2, tel3 As Integer
 Dim b As String
        tel2 = InStr(Inhoud, "<td>")
        tel3 = InStr(Inhoud, "</td>")
        b = Mid(Inhoud, tel2 + 4, tel3 - tel2 - 4)
        Inhoud = Mid(Inhoud, tel3 + 5)
        td_tag = b
End Function

' destilleer paren

' paar -->  username1 , username2

'http://admin.stepbridge.nl/show.php?page=tournamentinfo&activityid=20434&username=JanvGe


Public Function GetStepUserData(sUrl As Variant, ActiveID As Variant, User As Variant) As String
    Dim s, a, b, c As String
    Dim Scorestaat As String
    Dim TestNietGespeeld As Integer
    Dim TestKunstMatig, TestGewoon As Integer
    Dim Spelnr, Contract, Resultaat, Door, Score, ImpsButler, kleur As String
    
    
    Dim teller, teller2, tel2, tel3, tel4, telr2, telr3, hoogte As Integer
    s = GetHTMLFromURL(sUrl & ActiveID & "&" & "username=" & User)
    
    'zoek body
    
    teller = InStr(s, "<body>")
    
    
    If teller > 0 Then
     s = Mid(s, teller + 6)
    Else
     GetStepUserData = "-----"
     Exit Function
    End If
    
    teller = InStr(s, "<tbody>")
    s = Mid(s, teller + 7)
    
    
    s = StripHREF(s)
   'zoek body
    
    s = Replace(s, " align=" & Chr(34) & "right" & Chr(34), "")
    s = Replace(s, " align=" & Chr(34) & "left" & Chr(34), "")
    s = Replace(s, " align=" & Chr(34) & "center" & Chr(34), "")
    s = Replace(s, "&nbsp;", "")
     
    Do While Left(s, 8) <> "</tbody>"
        'zoek tr
        b = tr_tag(s)
        TestNietGespeeld = False
        TestKunstMatig = False
        TestGewoon = True
        
        'test op niet gespeeld
        'test op kunstmatige score
        '<td align="right"><a href="#1">1</a></td>
        '<td align="left" colspan="4" align="center">Kunstmatige Score</td>
        '<td align="right">50,00%</td>
        
        If InStr(b, "spel niet gespeeld") > 0 Then
           TestNietGespeeld = True
           TestGewoon = False
        End If
         
         If InStr(b, "kunstmatige") > 0 Then
          TestKunstMatig = True
          TestGewoon = False
        End If
        
        If TestNietGespeeld Then
            'spelnr
            Spelnr = Chr(34) & td_tag(b) & Chr(34)
            Scorestaat = Scorestaat & Spelnr & "," & "NGSP" & "," & "" & "," & "" & "," & "" & "," & "" & vbCr
        End If
        
        If TestKunstMatig Then
       
        'spelnr
            Spelnr = Chr(34) & td_tag(b) & Chr(34)
            Contract = Chr(34) & td_tag(b) & Chr(34)
            ImpsButler = Chr(34) & Mid(b, tel2 + 4, tel3 - tel2 - 4) & Chr(34)
            Scorestaat = Scorestaat & Spelnr & "," & "ARB" & "," & "" & "," & "" & "," & "" & "," & ImpsButler & vbCr
        End If
        
        If TestGewoon Then
             'spelnr
            Spelnr = Chr(34) & td_tag(b) & Chr(34)
            
            'contract
            a = td_tag(b)
            hoogte = Val(a)
            If hoogte <> 0 Then
                tel4 = InStr(a, "alt")
                If tel4 = 0 Then
                    Contract = Chr(34) & a & Chr(34)
                    Else
                        kleur = Mid(a, tel4)
                        
                        kleur = Replace(kleur, "alt=" & Chr(34), "")
                        kleur = Replace(kleur, Chr(34) & ">", "")
                       Contract = Chr(34) & hoogte & kleur & Chr(34)
                 End If
            Else
                Contract = Chr(34) & "Pass" & Chr(34)
            End If
   
           'Resultaat
            Resultaat = Chr(34) & td_tag(b) & Chr(34)
            
            'Door
            Door = Chr(34) & td_tag(b) & Chr(34)
              
            'Score
            Score = Chr(34) & td_tag(b) & Chr(34)
                  
            'Impsbutler
           
            ImpsButler = Chr(34) & td_tag(b) & Chr(34)
            
            Scorestaat = Scorestaat & Spelnr & "," & Contract & "," & Resultaat & "," & Door & "," & Score & "," & ImpsButler & vbCr
        
        End If
     Loop
    
    
    GetStepUserData = Scorestaat
     
End Function


Public Function StripHREF(a As Variant)

Dim teller, lengte As Long
Dim b, c As Long
Dim strA As String


strA = a

teller = 0


Do While InStr(strA, "<A") > 0
    b = InStr(strA, "<A")
    c = InStr(b, strA, ">")
    strA = Left(strA, b - 1) + Mid(strA, c + 1)
Loop

strA = Replace(strA, "</A>", "")

StripHREF = strA

End Function
Public Sub AlleScoreStatenExcel(Avond As Integer)
Dim i As Integer
Dim x

For i = 1 To 15
x = VulScoreKaartInSheet(i, Avond, True)
Next

End Sub

Public Sub AlleScoreStaten(Avond As Integer)
Dim i As Integer
Dim x
Dim ScoreSheetName   As String

'ScoreSheetName = "Avond_" & Avond & "_Teamnr_*"
'Call deleteSheet(ScoreSheetName)
 
For i = 1 To 15
x = VulScoreKaartInSheet(i, Avond, False)
Next

End Sub

Public Sub VulUitslagIn(Avond As Variant)
Dim Uitslag As String
Dim rijteller, teller, teller2, kolomteller, kolomteller2, Webinfo, kolomoffset, i, j As Integer
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim Rijen() As String
Dim Kolommen() As String

Set StartBook = Workbooks(1)

Set MySheet = StartBook.Sheets("WebInfo")
rijteller = 2
Do While True
If MySheet.Cells(rijteller, 1).Value = Avond Then
    Info = True
    Webinfo = MySheet.Cells(rijteller, 2).Value
    Exit Do
End If


 If MySheet.Cells(rijteller, 1).Value = "" Then
   Exit Sub
 End If
 rijteller = rijteller + 1
Loop

Uitslag = GetStepUitslag(STEPDATA, Webinfo)

Rijen = Split(Uitslag, vbCr)

Set MySheet = StartBook.Sheets("Import_Uitslag")
kolomoffset = (Avond - 1) * 4 + 1
rijteller = 1

MySheet.Cells(1, 0 + kolomoffset).Value = "Rang"
MySheet.Cells(1, 1 + kolomoffset).Value = "Spelers"
MySheet.Cells(1, 2 + kolomoffset).Value = "Score"
rijteller = 2
For teller = 0 To UBound(Rijen)
        Kolommen = Split(Rijen(teller), ",")
        For teller2 = 0 To UBound(Kolommen)
            MySheet.Cells(rijteller + teller, kolomoffset + teller2).Value = Replace(Kolommen(teller2), Chr(34), "")
        Next
Next

Set MySheet = Nothing



End Sub

Public Function VulScoreKaartInSheet(Teamnr As Variant, Avond As Variant, AparteExcel As Integer) As String

Dim offset_speler, wedstrijd, wedstrijd1, wedstrijd2 As Integer
Dim rijteller As Integer
Dim kolomteller As Integer
Dim Opstelling As Integer
Dim Webinfo, Info As Integer
Dim Speler1, Speler3, Speler2, Speler4 As String
Dim Wij, Tegenstander1, Tegenstander2 As String
Dim tegenst1, tegen1, tegen2, tegenst2, wij1, Wijzijn As Integer
Dim TemplateExcelfile As String
Dim ScoresheetExcelfile As String
Dim Scorestaat_Speler1, Scorestaat_Speler3, strFormula As String
Dim teller, teller2 As Integer
Dim ScoreSheetName As String
Dim MySheet As Worksheet
Dim TemplateBook As Workbook
Dim StartBook As Workbook
Dim TemplateSheet As Worksheet
Dim kolomscorespeler1, AantalSpelGespeeldWedstrijd1 As Integer
Dim kolomscorespeler3, AantalSpelGespeeldWedstrijd2 As Integer
Dim VPschaal As Integer
Dim WijScore1, WijScore2, ZijScore1, ZijScore2, WijImps1, ZijImps1, WijImps2, ZijImps2 As Double

Dim Rekenkamerfolder As String

Dim RekenkamerOutputfolder As String

Rekenkamerfolder = "C:\Users\pgjmw\Dropbox\DonderdagAvond\Rekenkamer\Viertallen\Seizoen_2020_2021\Excel\"
RekenkamerOutputfolder = "C:\Users\pgjmw\Dropbox\DonderdagAvond\Rekenkamer\Viertallen\Seizoen_2020_2021\"


TemplateExcelfile = Rekenkamerfolder & "\Team_Avond_Template.xlsx"
ScoresheetExcelfile = RekenkamerOutputfolder & "\Team_Avond_" & Avond & "_" & Teamnr & "_" & Format(Now(), "hh_mm") & ".xlsx"


Dim Rijen() As String
Dim Kolommen() As String

' benodigheden

'teamnr
'avond
'spelers

Set StartBook = Workbooks(1)

Set MySheet = StartBook.Sheets("Import_Opstelling")

Opstelling = False
    rijteller = 1
    Do While True
        If MySheet.Cells(rijteller, 1).Value = Avond And MySheet.Cells(rijteller, 2) = Teamnr Then
         Opstelling = True
         Exit Do
        End If
        If MySheet.Cells(rijteller, 1).Value = "" Then
            VulScoreKaartInSheet = "avond+Teamnr Niet gevonden"
            Exit Function
         End If
    rijteller = rijteller + 1
    Loop

Speler1 = MySheet.Cells(rijteller, 3).Value
Speler2 = MySheet.Cells(rijteller, 4).Value
Speler3 = MySheet.Cells(rijteller, 5).Value
Speler4 = MySheet.Cells(rijteller, 6).Value
tegenst1 = MySheet.Cells(rijteller, 7).Value
tegenst2 = MySheet.Cells(rijteller, 8).Value



Set MySheet = StartBook.Sheets("Teams")
rijteller = 2
Do While True
If MySheet.Cells(rijteller, 1).Value = tegenst1 Then
    tegen1 = True
    Tegenstander1 = MySheet.Cells(rijteller, 2).Value
End If
If MySheet.Cells(rijteller, 1).Value = Teamnr Then
    wij1 = True
    Wij = MySheet.Cells(rijteller, 2).Value
End If
If MySheet.Cells(rijteller, 1).Value = tegenst2 Then
    tegen2 = True
    Tegenstander2 = MySheet.Cells(rijteller, 2).Value
End If

If tegen1 And tegen2 And wij1 Then
    Exit Do
End If
 If MySheet.Cells(rijteller, 1).Value = "" Then
   VulScoreKaartInSheet = "Team Niet gevonden"
   Exit Function
 End If
 rijteller = rijteller + 1
Loop

Info = False

Set MySheet = StartBook.Sheets("WebInfo")
rijteller = 2
Do While True
If MySheet.Cells(rijteller, 1).Value = Avond Then
    Info = True
    Webinfo = MySheet.Cells(rijteller, 2).Value
    Exit Do
End If


 If MySheet.Cells(rijteller, 1).Value = "" Then
   VulScoreKaartInSheet = "Geen Webinfo"
   Exit Function
 End If
 rijteller = rijteller + 1
Loop


Scorestaat_Speler1 = GetStepUserData(STEPDATA, Webinfo, Speler1)
Scorestaat_Speler3 = GetStepUserData(STEPDATA, Webinfo, Speler3)





ScoreSheetName = "Avond_" & Avond & "_Teamnr_" & Teamnr

If AparteExcel Then

Set TemplateBook = Workbooks.Open(TemplateExcelfile)
Set TemplateSheet = TemplateBook.Worksheets("Team_Template")

Else

Call deleteSheet(ScoreSheetName)

StartBook.Sheets("Team_Template").Copy After:=StartBook.Sheets(StartBook.Sheets.Count)
ActiveSheet.Name = ScoreSheetName
Set TemplateSheet = StartBook.Sheets(ScoreSheetName)

End If


TemplateSheet.Cells(1, 5).Value = Speler1 & " - " & Speler2
TemplateSheet.Cells(1, 12).Value = Speler3 & " - " & Speler4
TemplateSheet.Cells(5, 20).Value = Wij
TemplateSheet.Cells(6, 20).Value = Tegenstander1
TemplateSheet.Cells(19, 20).Value = Wij
TemplateSheet.Cells(20, 20).Value = Tegenstander2
TemplateSheet.Cells(30, 3).Value = Wij

Rijen = Split(Scorestaat_Speler1, vbCr)
rijteller = 3
kolomteller = 1


For teller = 0 To UBound(Rijen) - 1
    Kolommen = Split(Rijen(teller), ",")
    For teller2 = 0 To UBound(Kolommen)
        TemplateSheet.Cells(rijteller + teller, kolomteller + teller2).Value = Replace(Kolommen(teller2), Chr(34), "")
    Next
Next

Rijen = Split(Scorestaat_Speler3, vbCr)
rijteller = 3
kolomteller = 8
For teller = 0 To UBound(Rijen) - 1
    Kolommen = Split(Rijen(teller), ",")
    For teller2 = 0 To UBound(Kolommen)
        TemplateSheet.Cells(rijteller + teller, kolomteller + teller2).Value = Replace(Kolommen(teller2), Chr(34), "")
    Next
Next

'indien NGSP

kolomscorespeler1 = 5
kolomscorespeler3 = 12
rijteller = 3

For teller = 0 To UBound(Rijen) - 1
 If TemplateSheet.Cells(rijteller + teller, kolomscorespeler1).Value = "" Or TemplateSheet.Cells(rijteller + teller, kolomscorespeler3).Value = "" Then
 'maak de regel leeg
 For teller2 = 2 To 6
    TemplateSheet.Cells(rijteller + teller, teller2).Value = ""
    TemplateSheet.Cells(rijteller + teller, teller2 + 7).Value = ""
 Next
End If

Next

'tel het aantal spellen dat gespeeld

'eerste wedstrijd

AantalSpelGespeeldWedstrijd1 = 0
AantalSpelGespeeldWedstrijd2 = 0
For teller = 0 To 11
If TemplateSheet.Cells(rijteller + teller, kolomscorespeler1).Value = "" Then
    AantalSpelGespeeldWedstrijd1 = AantalSpelGespeeldWedstrijd1 + 1
End If
Next

For teller = 12 To 23
If TemplateSheet.Cells(rijteller + teller, kolomscorespeler1).Value = "" Then
    AantalSpelGespeeldWedstrijd2 = AantalSpelGespeeldWedstrijd2 + 1
End If

Next
AantalSpelGespeeldWedstrijd1 = 12 - AantalSpelGespeeldWedstrijd1
AantalSpelGespeeldWedstrijd2 = 12 - AantalSpelGespeeldWedstrijd2

Select Case AantalSpelGespeeldWedstrijd1
Case Is < 6
VPschaal = 3
Case 6
VPschaal = 3
Case 7
VPschaal = 4
Case 8
VPschaal = 5
Case 9
VPschaal = 6
Case 10
VPschaal = 7
Case Else
VPschaal = 2
End Select
If VPschaal <> 2 Then
    strFormula = "=IF(V5>0,VLOOKUP(V5,VPSchaal," & VPschaal & "),20-VLOOKUP(V6,VPSchaal," & VPschaal & "))"
    TemplateSheet.Cells(5, 23).Formula = strFormula
    strFormula = "=IF(V6>0,VLOOKUP(V6,VPSchaal," & VPschaal & "),20-VLOOKUP(V5,VPSchaal," & VPschaal & "))"
    TemplateSheet.Cells(6, 23).Formula = strFormula
End If


Select Case AantalSpelGespeeldWedstrijd2
Case Is < 6
VPschaal = 3
Case 6
VPschaal = 3
Case 7
VPschaal = 4
Case 8
VPschaal = 5
Case 9
VPschaal = 6
Case 10
VPschaal = 7
Case Else
VPschaal = 2
End Select
If VPschaal <> 2 Then
     strFormula = "=IF(V19>0,VLOOKUP(V19,VPSchaal," & VPschaal & "),20-VLOOKUP(V20,VPSchaal," & VPschaal & "))"
     TemplateSheet.Cells(19, 23).Formula = strFormula
     strFormula = "=IF(V20>0,VLOOKUP(V20,VPSchaal," & VPschaal & "),20-VLOOKUP(V19,VPSchaal," & VPschaal & "))"
     TemplateSheet.Cells(20, 23).Formula = strFormula
End If

'pas formules aan =ALS(V5>0;VERT.ZOEKEN(V5;VPSchaal;2);20-VERT.ZOEKEN(V6;VPSchaal;2))


WijScore1 = TemplateSheet.Cells(5, 23).Value
ZijScore1 = TemplateSheet.Cells(6, 23).Value
WijImps1 = TemplateSheet.Cells(5, 21).Value
ZijImps1 = TemplateSheet.Cells(6, 21).Value

WijScore2 = TemplateSheet.Cells(19, 23).Value
ZijScore2 = TemplateSheet.Cells(20, 23).Value
WijImps2 = TemplateSheet.Cells(19, 21).Value
ZijImps2 = TemplateSheet.Cells(20, 21).Value



'Vul in de kruistabel
Set MySheet = StartBook.Sheets("Kruistabel")

MySheet.Cells(Teamnr + 1, tegenst1 + 1).Value = WijScore1
MySheet.Cells(tegenst1 + 1, Teamnr + 1).Value = ZijScore1
MySheet.Cells(Teamnr + 1, tegenst2 + 1).Value = WijScore2
MySheet.Cells(tegenst2 + 1, Teamnr + 1).Value = ZijScore2



'Vul in de uitslagen in

Set MySheet = StartBook.Sheets("TeamUitslagen")



' zoek of het team thuis spelen is in wedstrijd 1
wedstrijd1 = False


wedstrijd = 1


rijteller = 2

Do While MySheet.Cells(rijteller, 1) <> ""

If MySheet.Cells(rijteller, 1) = Avond And MySheet.Cells(rijteller, 2) = wedstrijd And MySheet.Cells(rijteller, 3) = Teamnr Then
 wedstrijd1 = True
 Exit Do
End If
rijteller = rijteller + 1
Loop

If wedstrijd1 And tegenst1 <> 16 Then
        MySheet.Cells(rijteller, 7).Value = WijImps1
        MySheet.Cells(rijteller, 8).Value = ZijImps1
        MySheet.Cells(rijteller, 9).Value = WijScore1
        MySheet.Cells(rijteller, 10).Value = ZijScore1
End If

' zoek of het team thuis spelen is in wedstrijd 2

wedstrijd2 = False
wedstrijd = 2
rijteller = 2
Do While MySheet.Cells(rijteller, 1) <> ""

If MySheet.Cells(rijteller, 1) = Avond And MySheet.Cells(rijteller, 2) = wedstrijd And MySheet.Cells(rijteller, 3) = Teamnr Then
 wedstrijd2 = True
 Exit Do
End If
rijteller = rijteller + 1
Loop

If wedstrijd2 And tegenst2 <> 16 Then
        MySheet.Cells(rijteller, 7).Value = WijImps2
        MySheet.Cells(rijteller, 8).Value = ZijImps2
        MySheet.Cells(rijteller, 9).Value = WijScore2
        MySheet.Cells(rijteller, 10).Value = ZijScore2
End If




If AparteExcel Then

TemplateSheet.Name = ScoreSheetName
TemplateBook.SaveAs ScoresheetExcelfile
TemplateBook.Close

Else

Set TemplateSheet = Nothing

End If

Set MySheet = Nothing


End Function

Sub deleteSheet(wsName As String)
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Sheets 'loop to find sheet (if it exists)
    Application.DisplayAlerts = False 'hide confirmation from user
    If ws.Name Like wsName Then ws.Delete 'found it! - delete it
    Application.DisplayAlerts = True 'show future confirmations
  Next ws
End Sub


Public Sub OpenHoofdMenu()
 frmStephoofdmenu.Show
End Sub
   
    Public Sub SluitHoofdMenu()
     Unload frmStephoofdmenu
    End Sub


