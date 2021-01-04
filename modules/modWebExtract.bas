Option Compare Text

Public Const ENKEL = 1
Public Const DUBBEL = 2
Public Const ROBIN = 1
Public Const ZWITSERS = 2



Public Team_Tegen_Avond() As Integer
Public WebInfo() As Integer
Public Rang() As Integer
Public Gemiddelde() As Double
Public Gespeeld() As Integer

'Extract uitslag

'we hebben de Url nodig + een ID, elke wedstrijd heeft een id
'http://admin.stepbridge.nl/show.php?page=tournamentinfo&activityid=20434



Public Function GetStepUitslag(sUrl As Variant, ActiveID As Variant, Data_of_Results As Variant) As String
    Dim s, a, b, c As String
    Dim Uitslag As String
    
    Dim Rang, spelers, score As String
    
    
    Dim teller, teller2, tel2, tel3, telr2, telr3 As Integer
    s = GetHTMLFromURL(sUrl & ActiveID)
    
    'zoek tag body
    
    
    
    'url is admin dan structuur van de tabellen is
    
    ' <tbody>
    '   <tbody>   eerste helft van de uitsag
    '   </tbody>  einde eerste helft
    '   <tbody>   tweede helft
    '   </tbody>   einde tweede helft
    ' </tbody  einde
    
     teller = InStr(s, "<body>")
     If teller > 0 Then
        s = Mid(s, teller + 6)
     Else
        'geen body tag dan exit functie
        GetStepUitslag = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
      Exit Function
     End If
     
     'zoek tag <tbody> is een tabel met inhoud
 's = Replace(s, Chr(10), "")
    's = Trim(s)
    teller = InStr(s, "<tbody>")
    If teller = 0 Then
       'indien geen tabelinhoud dan exit function
        GetStepUitslag = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
        Exit Function
    End If
    
    'we hebben de tweede tag nodig
    
    s = Mid(s, teller + 7)
    
    If Data_of_Results = 1 Then
    teller = InStr(s, "<tbody>")
    If teller = 0 Then
       'indien geen tweede tabelinhoud dan exit function
        GetStepUitslag = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
        Exit Function
    End If
     s = Mid(s, teller + 7)
   End If
  
    
    'strip wat tag's en text die we niet nodig hebben
    s = StripHREF(s)
   'zoek body
    If Data_of_Results = 1 Then
    s = Replace(s, " align=" & Chr(34) & "right" & Chr(34), "")
    s = Replace(s, " align=" & Chr(34) & "left" & Chr(34), "")
    Else
    s = Replace(s, "style=" & Chr(34) & "text-align: left" & Chr(34), "")
    s = Replace(s, "style=" & Chr(34) & "text-align: right" & Chr(34), "")
    s = Replace(s, "style=" & Chr(34) & "text-align: center" & Chr(34), "")
  End If
    s = Trim(s)
    'do zolang er rijen zij
   ' s = Replace(s, Chr(10), "")
    
    Do While Left(s, 8) <> "</tbody>"
        'zoek tr
        b = tr_tag(s)
  
        '3 kolommen
        
        'Kolom rang
       
        Rang = Chr(34) & td_tag(b) & Chr(34)
        
        'Kolom spelers
        
        spelers = Chr(34) & td_tag(b) & Chr(34)
 
        'Kolom score
        
        score = Chr(34) & td_tag(b) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & spelers & "," & score & vbCr
       
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
        
        spelers = Chr(34) & td_tag(b) & Chr(34)
 
        'Kolom score
        
        score = Chr(34) & td_tag(b) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & spelers & "," & score & vbCr
     Loop
    
    
    GetStepUitslag = Uitslag
End Function


Public Function GetStepUitslag_deelnemers(sUrl As Variant, ActiveID As Variant, Data_of_Results As Variant) As String
    Dim s, a, b, c As String
    Dim Uitslag As String
    
    Dim Rang, spelers, score As String
    
    
    Dim teller, teller2, tel2, tel3, telr2, telr3 As Integer
    s = GetHTMLFromURL(sUrl & ActiveID)
    
    'zoek tag body
    
    
    
    'url is admin dan structuur van de tabellen is
    
    ' <tbody>
    '   <tbody>   eerste helft van de uitsag
    '   </tbody>  einde eerste helft
    '   <tbody>   tweede helft
    '   </tbody>   einde tweede helft
    ' </tbody  einde
    
     teller = InStr(s, "<body>")
     If teller > 0 Then
        s = Mid(s, teller + 6)
     Else
        'geen body tag dan exit functie
        GetStepUitslag_deelnemers = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
      Exit Function
     End If
     
     'zoek tag <tbody> is een tabel met inhoud
 's = Replace(s, Chr(10), "")
    's = Trim(s)
    teller = InStr(s, "<tbody>")
    If teller = 0 Then
       'indien geen tabelinhoud dan exit function
        GetStepUitslag_deelnemers = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
        Exit Function
    End If
    
    'we hebben de tweede tag nodig
    
    s = Mid(s, teller + 7)
    
    If Data_of_Results = 1 Then
    teller = InStr(s, "<tbody>")
    If teller = 0 Then
       'indien geen tweede tabelinhoud dan exit function
        GetStepUitslag_deelnemers = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
        Exit Function
    End If
     s = Mid(s, teller + 7)
   End If
  
    
    'strip wat tag's en text die we niet nodig hebben
    s = StripHREF(s)
   'zoek body
    If Data_of_Results = 1 Then
    s = Replace(s, " align=" & Chr(34) & "right" & Chr(34), "")
    s = Replace(s, " align=" & Chr(34) & "left" & Chr(34), "")
    Else
    s = Replace(s, "style=" & Chr(34) & "text-align: left" & Chr(34), "")
    s = Replace(s, "style=" & Chr(34) & "text-align: right" & Chr(34), "")
    s = Replace(s, "style=" & Chr(34) & "text-align: center" & Chr(34), "")
  End If
    s = Trim(s)
    'do zolang er rijen zij
   ' s = Replace(s, Chr(10), "")
    
    Do While Left(s, 8) <> "</tbody>"
        'zoek tr
        b = tr_tag(s)
  
        '3 kolommen
        
        'Kolom rang
       
        Rang = Chr(34) & td_tag(b) & Chr(34)
        
      'Kolom spelers
        
        spelers = td_tag(b)
        
       teller2 = InStr(spelers, "-")
       Speler1 = Chr(34) & Trim(Left(spelers, teller2 - 1)) & Chr(34)
        Speler2 = Chr(34) & Trim(Mid(spelers, teller2 + 1)) & Chr(34)
        
        'score = Chr(34) & td_tag(b) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & Speler1 & "," & Speler2 & vbCr
       
    Loop
     
     'uitslag is twee kolommen
     
     teller = InStr(s, "<tbody>")
     'mocht er geen twee kolom zijn dan exit functie
     
     If teller = 0 Then
        GetStepUitslag_deelnemers = Uitslag
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
        
        spelers = td_tag(b)
        
       teller2 = InStr(spelers, "-")
       Speler1 = Chr(34) & Trim(Left(spelers, teller2 - 1)) & Chr(34)
        Speler2 = Chr(34) & Trim(Mid(spelers, teller2 + 1)) & Chr(34)
        
       ' score = Chr(34) & td_tag(b) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & Speler1 & "," & Speler2 & vbCr
     Loop
    
    
    GetStepUitslag_deelnemers = Uitslag
End Function
Public Function GetStepUitslag_d(sUrl As Variant, ActiveID As Variant) As String
    Dim s, a, b, c As String
    Dim Uitslag As String
    Dim Rijen_Uitslag() As String
    Dim Kolommen_Uitslag() As String
    Dim Rang, spelers, score, strTable As String
    Dim TweedeKolom As Integer
    Dim tel_tbody_1, tel_tbody_2 As Integer
    Dim teller, teller2, tel_1, tel_2, telr2, telr3 As Integer
    s = GetHTMLFromURL(sUrl & ActiveID)
    
    
    
    
    'url is admin dan structuur van de tabellen is
    
    ' <tbody>
    '   <tbody>   eerste helft van de uitsag
    '   </tbody>  einde eerste helft
    '   <tbody>   tweede helft
    '   </tbody>   einde tweede helft
    ' </tbody  einde
    
    'url is results dan structuur van de tabellen is
    '   <tbody>   eerste helft van de uitsag
    '   </tbody>  einde eerste helft
    '   <tbody>   tweede helft
    '   </tbody>   einde tweede helft
    
    
    If sUrl = STEPDATA Then
     'pak de tweede tbody
        tel_tbody_1 = InStr(s, "<tbody>")
        tel_tbody_1 = InStr(tel_tbody_1 + 7, s, "<tbody>")
        tel_tbody_2 = InStr(tel_tbody_1 + 7, s, "<tbody>")
    Else
        tel_tbody_1 = InStr(s, "<tbody>")
        tel_tbody_2 = InStr(tel_tbody_1 + 7, s, "<tbody>")
    End If
    
    TweedeKolom = True
    
    If tel_tbody_1 = 0 Then
       'indien geen tabelinhoud dan exit function
        GetStepUitslag2 = Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & Chr(34) & "--" & Chr(34) & vbCr
        Exit Function
    End If
    If tel_tbody_2 = 0 Then
        'geen tweede kolom
       TweedeKolom = False
    End If
        
    'we hebben de tweede tag nodig
    
    strTable = Mid(s, tel_tbody_1)
    
    
    a = grap_table(strTable)
    
    Rijen_Uitslag = grap_rijen(a)
    
    
    For tel_1 = LBound(Rijen_Uitslag) To UBound(Rijen_Uitslag) - 1
    
        Kolommen_Uitslag = grap_cellen(Rijen_Uitslag(tel_1))
        
        
        '3 kolommen
        
        'Kolom rang
       
        Rang = Chr(34) & stripTags(Kolommen_Uitslag(0)) & Chr(34)
        
        'Kolom spelers
        
        spelers = Chr(34) & extract_spelers(stripTags(Kolommen_Uitslag(1))) & Chr(34)
 
        'Kolom score
        
        score = Chr(34) & stripTags(Kolommen_Uitslag(2)) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & spelers & "," & score & vbCr
       
    Next
     
     
     If TweedeKolom = False Then
        GetStepUitslag2 = Uitslag
        Exit Function
    End If
     'uitslag is twee kolommen
    strTable = Mid(s, tel_tbody_2)
    
    
    a = grap_table(strTable)
    
    Rijen_Uitslag = grap_rijen(a)
    
    
   For tel_1 = LBound(Rijen_Uitslag) To UBound(Rijen_Uitslag) - 1
    
        Kolommen_Uitslag = grap_cellen(Rijen_Uitslag(tel_1))
      '3 kolommen
        
        'Kolom rang
       
        Rang = Chr(34) & stripTags(Kolommen_Uitslag(0)) & Chr(34)
        
        'Kolom spelers
        
        spelers = Chr(34) & extract_spelers(stripTags(Kolommen_Uitslag(1))) & Chr(34)
 
        'Kolom score
        
        score = Chr(34) & stripTags(Kolommen_Uitslag(2)) & Chr(34)
        
        Uitslag = Uitslag & Rang & "," & spelers & "," & score & vbCr
       
    Next
   
    GetStepUitslag_d = Uitslag
End Function

Public Function extract_spelers(spelers As Variant) As String
Dim s As String
s = spelers
s = Replace(s, " ", "")
s = Replace(s, "-", " - ")
extract_spelers = s
End Function

Public Function tr_tag(Inhoud As Variant) As String
 Dim telr2, telr3 As Integer
 Dim b As String
        telr2 = InStr(Inhoud, "<tr>")
        telr3 = InStr(Inhoud, "</tr>")
        b = Mid(Inhoud, telr2 + 4, telr3 - telr2 + 1)
        Inhoud = Trim(Mid(Inhoud, telr3 + 6))
        tr_tag = b
End Function

Public Function td_tag(Inhoud As Variant) As String
 Dim tel2, tel3 As Integer
 Dim b As String
        tel2 = InStr(Inhoud, "<td>")
        tel3 = InStr(Inhoud, "</td>")
        b = Mid(Inhoud, tel2 + 4, tel3 - tel2 - 4)
        b = Replace(b, " ", "")
        b = Replace(b, vbCr, "")
        b = Replace(b, vbLf, "")
        
        Inhoud = Trim(Mid(Inhoud, tel3 + 5))
        td_tag = b
End Function

' destilleer paren

' paar -->  username1 , username2

'http://admin.stepbridge.nl/show.php?page=tournamentinfo&activityid=21707&username=JanvGe


Public Function GetStepUserData(sUrl As Variant, ActiveID As Variant, User As Variant) As String
    Dim s, a, b, c As String
    Dim Scorestaat As String
    Dim TestNietGespeeld As Integer
    Dim TestKunstMatig, TestGewoon As Integer
    Dim Spelnr, Contract, resultaat, Door, score, ImpsButler, kleur As String
    
    
    Dim teller, teller2, tel2, tel3, tel4, telr2, telr3, hoogte As Integer
    
    
   s = GetHTMLFromURL(sUrl & ActiveID & "&" & "username=" & User)
   ' s = GetHTMLFromURL(sUrl & ActiveID & "/" & User)
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
    s = Trim(s)
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
            '<td>4<img src="/images/suit409.gif" alt="S">X</td>
            
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
            resultaat = Chr(34) & td_tag(b) & Chr(34)
            
            'Door
            Door = Chr(34) & td_tag(b) & Chr(34)
              
            'Score
            score = Chr(34) & td_tag(b) & Chr(34)
                  
            'Impsbutler
           
            ImpsButler = Chr(34) & td_tag(b) & Chr(34)
            ImpsButler = Replace(ImpsButler, "IMP", " IMP")
            Scorestaat = Scorestaat & Spelnr & "," & Contract & "," & resultaat & "," & Door & "," & score & "," & ImpsButler & vbCr
        
        End If
     Loop
    
    
    GetStepUserData = Scorestaat
     
End Function

Public Function GetStepUserData2(sUrl As Variant, ActiveID As Variant, User As Variant) As String
    Dim s, a, b, c As String
    Dim Scorestaat As String
    Dim TestNietGespeeld As Integer
    Dim TestKunstMatig, TestGewoon As Integer
    Dim Spelnr, Contract, resultaat, Door, score, ImpsButler, kleur As String
    
    
    Dim teller, teller2, tel2, tel3, tel4, telr2, telr3, hoogte As Integer
    
    
    ' GetHTMLFromURL(sUrl & ActiveID & "&" & "username=" & User)
   s = GetHTMLFromURL(sUrl & ActiveID & "/" & User)
   
   
    'zoek body
    
    teller = InStr(s, "<body>")
    
    
    If teller > 0 Then
     s = Mid(s, teller + 6)
    Else
     GetStepUserData2 = "-----"
     Exit Function
    End If
    
    teller = InStr(s, "<table class=" & Chr(34) & "results" & Chr(34) & ">")
    s = Mid(s, teller + 23)
    
    teller = InStr(s, "<tbody>")
    s = Mid(s, teller + 7)
    
    
    s = StripHREF(s)
   'zoek body
    '' style="text-align: center; padding: 0"
    s = Replace(s, " style=" & Chr(34) & "text-align: center; padding: 0" & Chr(34), "")
    s = Replace(s, " style=" & Chr(34) & "text-align: right; padding: 0 15px" & Chr(34), "")
    s = Replace(s, " style=" & Chr(34) & "text-align: right" & Chr(34), "")
    s = Replace(s, "&nbsp;", "")
    s = Trim(s)
    
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
            resultaat = Chr(34) & td_tag(b) & Chr(34)
            
            'Door
            Door = Chr(34) & "" & Chr(34)
              
            'Score
            score = Chr(34) & td_tag(b) & Chr(34)
                  
            'Impsbutler
           
            ImpsButler = Chr(34) & td_tag(b) & Chr(34)
            ImpsButler = Replace(ImpsButler, "IMP", " IMP")
            
            'Scorestaat = Scorestaat & Spelnr & "," & Contract & "," & Resultaat & "," & Score & "," & ImpsButler & vbCr
            Scorestaat = Scorestaat & Spelnr & "," & Contract & "," & resultaat & "," & Door & "," & score & "," & ImpsButler & vbCr
       
        End If
     Loop
    
    
    GetStepUserData2 = Scorestaat
     
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

Public Function stripTags(strCell As Variant) As String
Dim s As String
Dim tel, tel2 As Integer
s = strCell
s = Replace(s, Chr(10), "")
s = Replace(s, "&nbsp;", "")
tel = InStr(s, "<")
Do While tel > 0
    tel2 = InStr(s, ">")
    If tel = 1 Then
        s = Mid(s, tel2 + 1)
    Else
        s = Left(s, tel - 1) & Mid(s, tel2 + 1)
    End If
    tel = InStr(s, "<")
Loop
    
stripTags = Trim(s)
End Function

Public Sub AlleScoreStatenExcel(avond As Integer)
Dim i As Integer
Dim x

For i = 1 To 15
x = VulScoreKaartInSheet(i, avond, 1, lngToernooi)
Next

End Sub

Public Sub AlleScoreStaten_DATA(avond As Integer)
Dim i As Integer
Dim x
Dim Excelfile, Backupfile As String

'copy excel file naar backup.
Excelfile = WORKFOLDER & WORKFILE
If Dir(WORKFOLDER & "Backup\", vbDirectory) = "" Then
    MkDir WORKFOLDER & "Backup\"
End If

Backupfile = WORKFOLDER & "Backup\" & Left(WORKFILE, Len(WORKFILE) - 5) & "_" & Format(Now(), "hh_mm") & ".xlsx"
x = fnCopyfile(Excelfile, Backupfile)


'ScoreSheetName = "Avond_" & Avond & "_Teamnr_*"
'Call deleteSheet(ScoreSheetName)
 
For i = 1 To 15
x = VulScoreKaartInSheet(i, avond, 1, lngToernooi)
Next


End Sub
Public Sub AlleScoreStaten_RESULTS(avond As Integer, ToernID As Variant, SessieID As Variant)
Dim i, j As Integer
Dim x
Dim Excelfile, Backupfile As String

Call InitAll(ToernID, SessieID)

Excelfile = WORKFOLDER & WORKFILE
'copy excel file naar backup.
If Dir(WORKFOLDER & "Backup\", vbDirectory) = "" Then
    MkDir WORKFOLDER & "Backup\"
End If

Backupfile = WORKFOLDER & "Backup\" & Left(WORKFILE, Len(WORKFILE) - 5) & "_" & Format(Now(), "hh_mm") & ".xlsx"
x = fnCopyfile(Excelfile, Backupfile)


'ScoreSheetName = "Avond_" & Avond & "_Teamnr_*"
'Call deleteSheet(ScoreSheetName)
SysCmd acSysCmdInitMeter, "Process Scorestaten...", AANTALTEAMS - IIf(TEAMBYE > 0, 1, 0)
For i = 1 To AANTALTEAMS - IIf(TEAMBYE > 0, 1, 0)
     SysCmd acSysCmdUpdateMeter, i
    x = VulScoreKaartInSheet(i, avond, 2, ToernID, ScorestaatIntern, ScorestaatExcel)
Next
SysCmd acSysCmdRemoveMeter
End Sub

Public Sub InsertScoreSheets(avond As Integer)
Dim MySheet As Object
Dim StartBook As Object
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)



For j = 1 To 15
'verwijder oude indien aanwezig
ScoreSheetName = PREFIX & avond & "_Teamnr_" & j
sheetteller = StartBook.Sheets.Count
    For i = 1 To sheetteller
        If i > sheetteller Then
        Exit For
        End If
        
        If StartBook.Sheets(i).name = ScoreSheetName Then
            StartBook.Sheets(ScoreSheetName).Delete
            Exit For
        End If
    Next

sheetteller = StartBook.Sheets.Count
StartBook.Sheets("Team_Template").Copy After:=StartBook.Sheets(StartBook.Sheets.Count)
sheetteller = StartBook.Sheets.Count
StartBook.Sheets(sheetteller).name = ScoreSheetName
Set TemplateSheet = StartBook.Sheets(ScoreSheetName)
Next

    StartBook.Save
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing



End Sub
Public Sub VulUitslagIn(sUrl As Variant, avond As Variant)
Dim Uitslag, strWebInfo As String
Dim rijteller, teller, teller2, kolomteller, kolomteller2, kolomoffset, i, j As Integer
Dim MySheet As Object
Dim StartBook As Object
Dim Rijen() As String
Dim Kolommen() As String
Dim xlApp As Object



Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets("WebInfo")

rijteller = 2
Do While True
If MySheet.Cells(rijteller, 1).Value = avond Then
    info = True
    strWebInfo = MySheet.Cells(rijteller, 2).Value
    Exit Do
End If


 If MySheet.Cells(rijteller, 1).Value = "" Then
   Exit Sub
 End If
 rijteller = rijteller + 1
Loop

Uitslag = GetStepUitslag_d(sUrl, strWebInfo)

Rijen = Split(Uitslag, vbCr)

Set MySheet = StartBook.Sheets("Import_Uitslag")
kolomoffset = (avond - 1) * 4 + 1
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

    StartBook.Save
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub

Public Sub HTMLViertalUitslagenIn(avond As Variant, ToernID As Variant, SessieID As Variant)
Dim ViertalUitslag, WebInf, StepRef, ScoreSheetHTML, UitslagHTML, UitslagenHTML, HTMLFolder, Prefix_kopje, Suffix_kopje, Linkje, Voetje As String
Dim rijteller, teller, teller2, kolomteller, kolomteller2, kolomoffset, i, j As Integer
Dim intWebInfo As Integer

Dim Thuis, Uit As Variant
Dim refThuis, refUit As Variant
Dim Imps_Thuis As Variant
Dim Imps_Uit As Variant
Dim VPs_Thuis As Variant
Dim VPs_Uit As Variant

Dim MySheet As Object
Dim StartBook As Object
Dim Rijen() As String
Dim Kolommen() As String
Dim xlApp As Object

Dim rs As Recordset
Dim db As Database


Call InitAll(ToernID, SessieID)
 


Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)

'webinfo

Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblSessie  where [id] =" & lngSessie)
intWebInfo = rs!ActivityID
rs.Close
db.Close

'Set MySheet = StartBook.Worksheets("WebInfo")
'
'rijteller = 2
'Do While True
'If MySheet.Cells(rijteller, 1).Value = avond Then
 '   info = True
 '   intWebInfo = MySheet.Cells(rijteller, 2).Value
 '   Exit Do
'End If


' If MySheet.Cells(rijteller, 1).Value = "" Then
'   Exit Sub
 'End If
' rijteller = rijteller + 1
'Loop

'ViertalUitslag = GetStepViertalUitslag(STEPDATA, WebInfo)

'Rijen = Split(Uitslag, vbCr)
HTMLFolder = LOCALHTML


If ToernID > 1 Then
UitslagHTML = HTMLFolder & "Uitslagen_" & ToernID & "_" & PREFIX & avond & ".html"
Else
UitslagHTML = HTMLFolder & "Uitslagen" & PREFIX & avond & ".html"
End If

UitslagenHTML = ""
UitslagenHTML = UitslagenHTML & html_header()
UitslagenHTML = UitslagenHTML & html_Begin_Body()

If PrefixKopjeuitslagen = "" Then
    Prefix_kopje = "Uitslagen sessie " & avond
 Else
    Prefix_kopje = PrefixKopjeuitslagen & " "
End If


If SuffixKopjeuitslagen = "" Then
    Suffix_kopje = ""
 Else
    Suffix_kopje = SuffixKopjeuitslagen & " "
End If


UitslagenHTML = UitslagenHTML & rij_header(Prefix_kopje & Suffix_kopje)

UitslagenHTML = UitslagenHTML & TeamUitslagenResultheader()

Set MySheet = StartBook.Sheets("TeamUitslagen")

rijteller = 2

Do While MySheet.Cells(rijteller, 1).Value <> ""
    
    If MySheet.Cells(rijteller, 1).Value = avond Then
    
        'plot regel
        
        Thuis = MySheet.Cells(rijteller, 5).Value
        Uit = MySheet.Cells(rijteller, 6).Value
        Imps_Thuis = MySheet.Cells(rijteller, 7).Value
        Imps_Uit = MySheet.Cells(rijteller, 8).Value
        VPs_Thuis = MySheet.Cells(rijteller, 9).Value
        VPs_Uit = MySheet.Cells(rijteller, 10).Value
        refThuis = LOCALSITE & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & MySheet.Cells(rijteller, 3).Value & ".html"
        refUit = LOCALSITE & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & MySheet.Cells(rijteller, 4).Value & ".html"
        
        If Thuis = "Bye" Then
            refThuis = refUit
            Imps_Thuis = "&nbsp;"
            Imps_Uit = "&nbsp;"
            VPs_Thuis = "&nbsp;"
            VPs_Uit = "&nbsp;"
        End If
        
        If Uit = "Bye" Then
            refUit = refThuis
            Imps_Thuis = "&nbsp;"
            Imps_Uit = "&nbsp;"
            VPs_Thuis = "&nbsp;"
            VPs_Uit = "&nbsp;"
        End If
        If rijteller Mod (AANTALTEAMS \ 2) = 2 And (rijteller < 2 + avond * (AANTALTEAMS \ 2) * WEDSTRIJDENPERSESSIE) Then
            UitslagenHTML = UitslagenHTML & TeamUnderlineUitslagenResultRow(Thuis, refThuis, Uit, refUit, Imps_Thuis, Imps_Uit, VPs_Thuis, VPs_Uit)
        Else
            UitslagenHTML = UitslagenHTML & TeamUitslagenResultRow(Thuis, refThuis, Uit, refUit, Imps_Thuis, Imps_Uit, VPs_Thuis, VPs_Uit)
        End If
    End If
    rijteller = rijteller + 1
Loop
    
UitslagenHTML = UitslagenHTML & TeamUitslagenResultfooter()
UitslagenHTML = UitslagenHTML & html_Einde_Body()

If fnExists(UitslagHTML) Then
    Kill (UitslagHTML)
End If




Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile(UitslagHTML)
oFile.Write UitslagenHTML
oFile.Close
Set fso = Nothing
Set oFile = Nothing

    
    

    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub

Public Sub HTMLViertalKruistabel(ToernID As Variant)
Dim ViertalUitslag, WebInf, StepRef, ScoreSheetHTML, UitslagHTML, KruisHTML, HTMLFolder As String
Dim rijteller, teller, teller2, kolomteller, kolomteller2, kolomoffset, i, j, avond As Integer


Dim Thuis, Uit As Variant
Dim refThuis, refUit As Variant
Dim Imps_Thuis As Variant
Dim Imps_Uit As Variant
Dim VPs_Thuis As Variant
Dim VPs_Uit As Variant

Dim MySheet As Object
Dim StartBook As Object
Dim Kolommen() As Variant
Dim xlApp As Object
'altijd laatste sessie

Dim db As Database
Dim rs As Recordset
Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblSessie where [ToernooiD] = " & ToernID & " Order by Sessienr")
rs.MoveLast
lngSessie = rs!id
rs.Close
db.Close

'fris altijd sessie gegevens op
lngSessieOld = 0
Call InitAll(ToernID, lngSessie)

ReDim WebInfo(20) As Integer
ReDim Kolommen(AANTALTEAMS + 4) As Variant
ReDim Team_Tegen_Avond(AANTALTEAMS + 4, AANTALTEAMS + 4) As Integer


Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)


Set MySheet = StartBook.Worksheets("Schema")
rijteller = 2
kolomteller = 1
teller = 1
Do While MySheet.Cells(rijteller + teller, 1).Value <> ""

'For teller = 1 To 15
        avond = (teller - 1) \ WEDSTRIJDENPERSESSIE + 1
        For i = 1 To AANTALTEAMS \ 2
            Team_Tegen_Avond(MySheet.Cells(rijteller + teller, i * 2).Value, MySheet.Cells(rijteller + teller, i * 2 + 1).Value) = avond
             Team_Tegen_Avond(MySheet.Cells(rijteller + teller, i * 2 + 1).Value, MySheet.Cells(rijteller + teller, i * 2).Value) = avond
        Next
'Next
    teller = teller + 1
 Loop
 
Set MySheet = StartBook.Worksheets("WebInfo")
rijteller = 2
info = False
Do While MySheet.Cells(rijteller, 2).Value <> ""
    WebInfo(rijteller - 1) = MySheet.Cells(rijteller, 2).Value
    rijteller = rijteller + 1
Loop


'ViertalUitslag = GetStepViertalUitslag(STEPDATA, WebInfo)

'Rijen = Split(Uitslag, vbCr)
HTMLFolder = LOCALHTML
If ToernID > 1 Then
    kruistabelHTML = HTMLFolder & "Kruistabel_" & ToernID & ".html"
Else
    kruistabelHTML = HTMLFolder & "Kruistabel.html"
End If

    
KruisHTML = ""
KruisHTML = KruisHTML & html_header()
KruisHTML = KruisHTML & html_Begin_Body()
KruisHTML = KruisHTML & rij_header("Kruistabel")




Set MySheet = StartBook.Worksheets("Kruistabel")
rijteller = 1
kolomteller = 1
For rijteller = 1 To AANTALTEAMS + 1
    For kolomteller = 1 To AANTALTEAMS + 4

        Kolommen(kolomteller) = MySheet.Cells(rijteller, kolomteller).Value
    
        'plot regel
    Next
        If TEAMBYE > 0 Then
            Select Case rijteller
            Case 1
                KruisHTML = KruisHTML & Teamkruisheader(Kolommen)
            Case TEAMBYE + 1
                KruisHTML = KruisHTML & Byekruisheaderrow(Kolommen)
            Case Else
                 KruisHTML = KruisHTML & Teamkruisheaderrow(Kolommen, rijteller - 1)
            End Select
        Else
            Select Case rijteller
            Case 1
                KruisHTML = KruisHTML & Teamkruisheader(Kolommen)
            Case Else
                 KruisHTML = KruisHTML & Teamkruisheaderrow(Kolommen, rijteller - 1)
            End Select
        End If
        
        
Next

    
KruisHTML = KruisHTML & TeamResultfooter()
KruisHTML = KruisHTML & Teamkruisfooter
KruisHTML = KruisHTML & html_Einde_Body()

If fnExists(kruistabelHTML) Then
    Kill (kruistabelHTML)
End If




Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile(kruistabelHTML)
oFile.Write KruisHTML
oFile.Close
Set fso = Nothing
Set oFile = Nothing

    
    

    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub
Public Function VulScoreKaartInSheet(Teamnr As Variant, avond As Variant, Data_of_Results As Integer, ToernooiID As Variant, Optional Access As Integer, Optional Excel As Integer) As String

Dim offset_speler, intTeamnr, intVorigeTeamnr, intRijtellerEersteTeam As Integer
Dim rijteller, wedstrijdteller As Integer
Dim kolomteller, sheetteller, progressteller, progoffset As Integer
Dim Opstelling, AparteExcel As Integer
Dim intWebInfo, info As Integer
Dim Speler1, Speler3, Speler2, Speler4, Paar1, Paar2 As String
Dim refSpeler1, refSpeler3, refSpeler2, refSpeler4 As String
Dim StepRef, StepUserRef As String
Dim PWORef, PWOTeamRef As String
Dim Criterium, Criterium1, Criterium2 As String
Dim ThuisWedstrijd, UitWedstrijd, AantalspellenGespeeld As Integer
Dim Wij, Tegenstander1, Tegenstander2 As String
Dim refWij, refTegenstander1, refTegenstander2 As String
Dim strE, strF, strG, strH, strI, strJ As String
Dim t_score1, t_score2, t_saldo, t_imps, t_wijimps, t_zijimps As Integer
Dim tegenst1, tegen1, tegen2, tegenst2, wij1, Wijzijn As Integer
Dim TemplateExcelfile As String
Dim ScoresheetExcelfile As String
Dim Scorestaat_Speler1, Scorestaat_Speler3, strFormula, Kopje_prefix, Kopje_suffix, Voetje As String
Dim teller, teller2 As Integer
Dim ScoreSheetName As String
Dim MySheet As Worksheet
Dim TemplateBook As Workbook
Dim StartBook As Workbook
Dim TemplateSheet As Worksheet
Dim kolomscorespeler1, AantalSpelGespeeldWedstrijd1 As Integer
Dim kolomscorespeler3, AantalSpelGespeeldWedstrijd2 As Integer
Dim ThuisID, UitID, UitslagID, WijID, ZijID, TeamByeID As Long

Dim WijVPs1, ZijVPs1 As Double

Dim xlApp As Object
Dim fso As Object
Dim oFile As Object

Dim WijImps() As Double
Dim ZijImps() As Double
Dim WijVPs() As Double
Dim ZijVPs() As Double
Dim Tegenstander() As String
Dim refTegenstander() As String
Dim tegenst() As Integer
Dim tegen() As Integer
Dim tegenID() As Long
Dim AantalSpelGespeeldWedstrijd() As Integer
Dim VPschaal() As Integer
Dim intWedstrijd() As Integer
Dim intWedstr As Integer
Dim resultaat(3, 24) As Integer
Dim TestOpstelling() As Integer
Dim TestSpelersOpstelling() As Integer
Dim ImpsVoor, ImpsTegen As Integer
Dim db As Database
Dim rs As Recordset
Dim us As Recordset

    lngSessie = DLookup("id", "tblSessie", "[ToernooiD] = " & ToernooiID & " and [Sessienr] = " & avond)


Call InitAll(ToernooiID, lngSessie)
'ScorestaatIntern = False
'ScorestaatExcel = False
'If Not IsNull(Access) Then
'    ScorestaatIntern = Access
'    If Not IsNull(Excel) Then
'     ScorestaatExcel = Excel
'    End If
'Else
'    ScorestaatIntern = False
'    ScorestaatExcel = True
'End If

Dim Rekenkamerfolder, HTMLFolder As String

ReDim WijImps(WEDSTRIJDENPERSESSIE)
ReDim ZijImps(WEDSTRIJDENPERSESSIE)
ReDim WijVPs(WEDSTRIJDENPERSESSIE)
ReDim ZijVPs(WEDSTRIJDENPERSESSIE)
ReDim Tegenstander(WEDSTRIJDENPERSESSIE)
ReDim refTegenstander(WEDSTRIJDENPERSESSIE)
ReDim tegenst(WEDSTRIJDENPERSESSIE)
ReDim tegen(WEDSTRIJDENPERSESSIE)
ReDim tegenID(WEDSTRIJDENPERSESSIE)
ReDim AantalSpelGespeeldWedstrijd(WEDSTRIJDENPERSESSIE)
ReDim VPschaal(WEDSTRIJDENPERSESSIE)
ReDim intWedstrijd(WEDSTRIJDENPERSESSIE)



Rekenkamerfolder = "C:\Users\pgjmw\Dropbox\DonderdagAvond\Rekenkamer"
HTMLFolder = LOCALHTML

AparteExcel = False

TemplateExcelfile = Rekenkamerfolder & "\Team_Avond_Template.xlsx"
ScoresheetExcelfile = Rekenkamerfolder & "\Team_" & PREFIX & avond & "_" & Teamnr & "_" & Format(Now(), "hh_mm") & ".xlsx"

'speler1
Dim Rijen() As String
Dim Kolommen() As String

'speler3
Dim Rijen2() As String
Dim Kolommen2() As String


If Data_of_Results = 1 Then
    StepRef = STEPDATA
Else
    StepRef = STEPRESULTS
End If

PWORef = LOCALSITE


' benodigheden

'teamnr
'avond
'spelers
If ScorestaatIntern Then
    'bepaal WijID
    
    If BerekenAlleStaten = False Then
        SysCmd acSysCmdInitMeter, "Bereken scorestaat ...", WEDSTRIJDENPERSESSIE * AANTALSPELLENPERWEDSTRIJD
    End If
    
    WijID = DLookup("id", "tblTeams", "[Teamnr] = " & Teamnr & " and [ToernooiID] =" & lngToernooi)
 
    Wij = DLookup("TeamNaam", "tblTeams", "[id] = " & WijID)
    
    If TEAMBYE <> 0 Then
        TeamByeID = DLookup("id", "tblTeams", "[Teamnr] = " & TEAMBYE & " and [ToernooiID] =" & lngToernooi)
    End If
    
        
    'stepID
    
    intWebInfo = DLookup("ActivityID", "tblSessie", "[id] = " & lngSessie)
    If Data_of_Results = 1 Then
        StepUserRef = StepRef & intWebInfo & "&username="
        Else
        StepUserRef = StepRef & intWebInfo & "/"
    End If
    'opstelling
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Select * from tblOpstelling where [ToernooiID] = " & lngToernooi & " and [SessieID] = " & lngSessie & " and [TeamID] = " & WijID)
    If rs.BOF And rs.EOF Then
        MsgBox ("Opstelling niet gevonden van team " & Teamnr)
        Exit Function
    End If
    
    Speler1 = rs.Fields("Speler1")
    Speler2 = rs.Fields("Speler2")
    Speler3 = rs.Fields("Speler3")
    Speler4 = rs.Fields("Speler4")

    Paar1 = Speler1 & " - " & Speler2
    Paar2 = Speler3 & " - " & Speler4
    refSpeler1 = StepUserRef & Speler1
    refSpeler2 = StepUserRef & Speler2
    refSpeler3 = StepUserRef & Speler3
    refSpeler4 = StepUserRef & Speler4
    
     For i = 1 To WEDSTRIJDENPERSESSIE
        tegenst(i) = rs.Fields("Wedstrijd" & i)
        'bepaal zijID
        tegenID(i) = DLookup("id", "tblTeams", "[Teamnr] = " & tegenst(i) & " and [ToernooiID] =" & lngToernooi)
        Tegenstander(i) = DLookup("TeamNaam", "tblTeams", "[id] = " & tegenID(i))
        
     Next
     refWij = PWORef & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & Teamnr & ".html"
     
     For i = 1 To WEDSTRIJDENPERSESSIE
            refTegenstander(i) = PWORef & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & tegenst(i) & ".html"
     Next

    rs.Close
    db.Close
    Scorestaat_Speler1 = GetStepUserData_d(StepRef, intWebInfo, Speler1)
    Scorestaat_Speler3 = GetStepUserData_d(StepRef, intWebInfo, Speler3)
    ScoreSheetName = PREFIX & avond & "_Teamnr_" & Teamnr
    strSheetName = ScoreSheetName
    
    intAantalWedstrijdenGespeeld = 0
    If Sessienr > 1 Then
        For i = 1 To Sessienr - 1
         intAantalWedstrijdenGespeeld = intAantalWedstrijdenGespeeld + _
         DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & lngToernooi & " and  [sessienr] = " & i)
        Next
    End If
    
      Set db = CurrentDb
      Set rs = db.OpenRecordset("tblScorestaat")
       'eerst test op sheet aanwezig is
      If Not IsNull(DLookup("id", "tblScorestaat", "[SessieID] = " & lngSessie & " and [TeamID] = " & WijID)) Then
          'verwijder records
           db.Execute "DELETE * FROM " _
                & "tblScorestaat WHERE [SessieID] = " & lngSessie & "and [TeamID] = " & WijID
      End If
      'rs.MoveFirst
    Rijen = Split(Scorestaat_Speler1, vbCr)
    Rijen2 = Split(Scorestaat_Speler3, vbCr)
    ' zoek wedstrijdnr op
    intWedstr = intAantalWedstrijdenGespeeld + 1
    AantalspellenGespeeld = 0
    wedstrijdteller = 1
    ImpsVoor = 0
    ImpsTegen = 0
    WijVPs1 = 0
    ZijVPs1 = 0
    
    For teller = 0 To UBound(Rijen) - 1
        If BerekenAlleStaten = False Then
          SysCmd acSysCmdUpdateMeter, teller + 1
        End If

        If teller Mod AANTALSPELLENPERWEDSTRIJD = 0 And teller > 0 Then '0
                ThuisWedstrijd = False
                UitWedstrijd = False
                Criterium1 = "[SessieID] = " & lngSessie & " and [TeamIDThuis] = " & WijID & " and [Wedstrijdnr] = " & intWedstr
                If Not IsNull(DLookup("id", "tblUitslagen", Criterium1)) Then
                   ThuisWedstrijd = True
                   UitslagID = DLookup("id", "tblUitslagen", Criterium1)
                   If DLookup("TeamIDUit", "tblUitslagen", "[id] = " & UitslagID) = TeamByeID Then
                    ThuisWedstrijd = False
                   End If
                Else
                   Criterium2 = "[SessieID] = " & lngSessie & " and [TeamIDUit] = " & WijID & " and [Wedstrijdnr] = " & intWedstr
                   If Not IsNull(DLookup("id", "tblUitslagen", Criterium2)) Then
                        UitWedstrijd = True
                        UitslagID = DLookup("id", "tblUitslagen", Criterium2)
                        If DLookup("TeamIDThuis", "tblUitslagen", "[id] = " & UitslagID) = TeamByeID Then
                        UitWedstrijd = False
                   End If
                   End If
                End If
               
                If ThuisWedstrijd = True Or UitWedstrijd = True Then '1
                    'update uitslag
                    'Bepaal eerst hoeveel spellen er zijn gespeeld
                    
                     
                    Set us = db.OpenRecordset("select * from tblUitslagen where [id] = " & UitslagID)
                    If Not (us.BOF And us.EOF) Then '2
                        If ThuisWedstrijd Then 'begin Thuis wedstrijd
                            us.Edit
                            Select Case UITREKENVORM
                            Case VPs_u
                                us!ImpsThuis = ImpsVoor
                                us!ImpsUit = ImpsTegen
                                WijVPs1 = BepaalVPs(ImpsVoor - ImpsTegen, AantalspellenGespeeld)
                                us!VpsThuis = WijVPs1
                                ZijVPs1 = BepaalVPs(ImpsTegen - ImpsVoor, AantalspellenGespeeld)
                                us!VpsUit = ZijVPs1
                            Case Imps_u
                                us!ImpsThuis = ImpsVoor
                                us!ImpsUit = ImpsTegen
                            Case Patton_u
                                us!ImpsThuis = ImpsVoor
                                us!ImpsUit = ImpsTegen
                            End Select
                            us.Update
                            us.Close
                            WijImps(wedstrijdteller) = ImpsVoor
                            ZijImps(wedstrijdteller) = ImpsTegen
                            Select Case UITREKENVORM
                            Case VPs_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                                WijVPs(wedstrijdteller) = WijVPs1
                                ZijVPs(wedstrijdteller) = ZijVPs1
                            Case Imps_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                            Case Patton_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                            End Select
                        End If  'end Thuis wedstrijd
                        If UitWedstrijd Then 'Begin uit wedstrijd
                            us.Edit
                            Select Case UITREKENVORM
                            Case VPs_u
                                us!ImpsThuis = ImpsTegen
                                us!ImpsUit = ImpsVoor
                                WijVPs1 = BepaalVPs(ImpsVoor - ImpsTegen, AantalspellenGespeeld)
                                ZijVPs1 = BepaalVPs(ImpsTegen - ImpsVoor, AantalspellenGespeeld)
                                us!VpsThuis = ZijVPs1
                                us!VpsUit = WijVPs1
                            Case Imps_u
                                us!ImpsThuis = ImpsTegen
                                us!ImpsUit = ImpsVoor
                            Case Patton_u
                                us!ImpsThuis = ImpsTegen
                                us!ImpsUit = ImpsVoor
                            End Select
                            us.Update
                            us.Close
                            WijImps(wedstrijdteller) = ImpsVoor
                            ZijImps(wedstrijdteller) = ImpsTegen
                            Select Case UITREKENVORM
                            Case VPs_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                                WijVPs(wedstrijdteller) = WijVPs1
                                ZijVPs(wedstrijdteller) = ZijVPs1
                            Case Imps_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                            Case Patton_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                            End Select
                        End If 'eind uit wedstrijd
                 End If 'end 2
              End If 'end 1
                
                'uitslagID = belend
                
                'bepaal vp's

              intWedstr = intWedstr + 1
              wedstrijdteller = wedstrijdteller + 1
              AantalspellenGespeeld = 0
              ImpsVoor = 0
              ImpsTegen = 0
              WijVPs1 = 0
              ZijVPs1 = 0
              ThuisWedstrijd = False
              UitWedstrijd = False
         End If '0
         Kolommen = Split(Rijen(teller), ",")
         Kolommen2 = Split(Rijen2(teller), ",")
         rs.AddNew
         rs!SessieID = lngSessie
         rs!TeamID = WijID
         rs!WEDSTRIJD = intWedstr
         rs!ToernooiID = lngToernooi
         rs!Paar1 = Paar1
         rs!Paar2 = Paar2
         rs!Spelnr = Val(Kolommen(0))
         If Kolommen(1) <> "" Then
            rs!Contract1 = Kolommen(1)
         End If
         If Kolommen(2) <> "" Then
            rs!Resultaat1 = Kolommen(2)
         End If
         If Kolommen(3) <> "" Then
            rs!Door1 = Kolommen(3)
         End If
        
         If Kolommen(4) <> "" Then
            t_score1 = Val(Kolommen(4))
            rs!score1 = t_score1
         End If
         
        If Kolommen2(1) <> "" Then
            rs!Contract2 = Kolommen2(1)
         End If
         If Kolommen2(2) <> "" Then
            rs!Resultaat2 = Kolommen2(2)
         End If
         If Kolommen2(3) <> "" Then
            rs!Door2 = Kolommen2(3)
         End If
        
         If Kolommen2(4) <> "" Then
            t_score2 = Val(Kolommen2(4))
            rs!score2 = t_score2
         End If
         
         If Kolommen(4) <> "" And Kolommen2(4) <> "" Then
            AantalspellenGespeeld = AantalspellenGespeeld + 1
            saldo = t_score1 + t_score2
            rs!saldo = saldo
            imps = BepaalImps(saldo)
            rs!imps = imps
            If imps > 0 Then
                rs!WijImps = imps
                ImpsVoor = ImpsVoor + imps
               Else
              If imps < 0 Then
                rs!ZijImps = Abs(imps)
                ImpsTegen = ImpsTegen + Abs(imps)
              End If
           End If
        End If
        rs.Update
      
      Next
      
      
      If teller Mod AANTALSPELLENPERWEDSTRIJD = 0 And teller > 0 Then
                ThuisWedstrijd = False
                UitWedstrijd = False
                Criterium1 = "[SessieID] = " & lngSessie & " and [TeamIDThuis] = " & WijID & " and [Wedstrijdnr] = " & intWedstr
                If Not IsNull(DLookup("id", "tblUitslagen", Criterium1)) Then
                   ThuisWedstrijd = True
                   UitslagID = DLookup("id", "tblUitslagen", Criterium1)
                   If DLookup("TeamIDUit", "tblUitslagen", "[id] = " & UitslagID) = TeamByeID Then
                    ThuisWedstrijd = False
                   End If
                Else
                   Criterium2 = "[SessieID] = " & lngSessie & " and [TeamIDUit] = " & WijID & " and [Wedstrijdnr] = " & intWedstr
                   If Not IsNull(DLookup("id", "tblUitslagen", Criterium2)) Then
                        UitWedstrijd = True
                        UitslagID = DLookup("id", "tblUitslagen", Criterium2)
                        If DLookup("TeamIDThuis", "tblUitslagen", "[id] = " & UitslagID) = TeamByeID Then
                            UitWedstrijd = False
                        End If
                   End If
                End If
               
                If ThuisWedstrijd = True Or UitWedstrijd = True Then
                    'update uitslag
                    'Bepaal eerst hoeveel spellen er zijn gespeeld
                    
                     
                    Set us = db.OpenRecordset("select * from tblUitslagen where [id] = " & UitslagID)
                    If Not (us.BOF And us.EOF) Then
                        If ThuisWedstrijd Then
                            us.Edit
                            Select Case UITREKENVORM
                             Case VPs_u
                                us!ImpsThuis = ImpsVoor
                                us!ImpsUit = ImpsTegen
                                WijVPs1 = BepaalVPs(ImpsVoor - ImpsTegen, AantalspellenGespeeld)
                                us!VpsThuis = WijVPs1
                                ZijVPs1 = BepaalVPs(ImpsTegen - ImpsVoor, AantalspellenGespeeld)
                                us!VpsUit = ZijVPs1
                            Case Imps_u
                                us!ImpsThuis = ImpsVoor
                                us!ImpsUit = ImpsTegen
                            Case Patton_u
                                us!ImpsThuis = ImpsVoor
                                us!ImpsUit = ImpsTegen
                            End Select
                            us.Update
                            us.Close
                            Select Case UITREKENVORM
                            Case VPs_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                                WijVPs(wedstrijdteller) = WijVPs1
                                ZijVPs(wedstrijdteller) = ZijVPs1
                            Case Imps_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                            Case Patton_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                            End Select
                        End If
                        If UitWedstrijd Then
                            us.Edit
                            Select Case UITREKENVORM
                            Case VPs_u
                                us!ImpsThuis = ImpsTegen
                                us!ImpsUit = ImpsVoor
                                WijVPs1 = BepaalVPs(ImpsVoor - ImpsTegen, AantalspellenGespeeld)
                                ZijVPs1 = BepaalVPs(ImpsTegen - ImpsVoor, AantalspellenGespeeld)
                                us!VpsThuis = ZijVPs1
                                us!VpsUit = WijVPs1
                           Case Imps_u
                                us!ImpsThuis = ImpsTegen
                                us!ImpsUit = ImpsVoor
                           Case Patton_u
                                us!ImpsThuis = ImpsTegen
                                us!ImpsUit = ImpsVoor
                           End Select
                           us.Update
                           us.Close
                           Select Case UITREKENVORM
                           Case VPs_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                                WijVPs(wedstrijdteller) = WijVPs1
                                ZijVPs(wedstrijdteller) = ZijVPs1
                          Case Imps_u
                                WijImps(wedstrijdteller) = ImpsVoor
                                ZijImps(wedstrijdteller) = ImpsTegen
                          Case Patton_u
                               WijImps(wedstrijdteller) = ImpsVoor
                               ZijImps(wedstrijdteller) = ImpsTegen
                          End Select
                        End If
                    End If
                End If
   
           End If
      
     db.Close
     If BerekenAlleStaten = False Then
          SysCmd acSysCmdRemoveMeter
     End If
End If
 

'
'test even of de workfile aanweig is




'test op opstelling



If ScorestaatExcel Then
    If Not fnExists(WORKFOLDER & WORKFILE) Then
        MsgBox ("applicatie kan " & WORKFOLDER & WORKFILE & "niet vinden ")
        Exit Function
    End If
    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    
    Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
    
    'test workfile
    
    Set MySheet = StartBook.Worksheets("Import_Opstelling")
     
    '
    
    'selecteer de sessie
    
        Opstelling = False
        rijteller = 1
        Do While True
            If MySheet.Cells(rijteller, 1).Value = avond Then
             intRijtellerEersteTeam = rijteller
             Opstelling = True
             Exit Do
            End If
            If MySheet.Cells(rijteller, 1).Value = "" Then
                VulScoreKaartInSheet = "avond+Teamnr Niet gevonden"
                Exit Function
             End If
        rijteller = rijteller + 1
        Loop
    'team gevonden
    If Opstelling = False Then
         MsgBox ("applicatie kan in " & WORKFOLDER & WORKFILE & " geen opstelling vinden van sessie/avond/wedstrijd " & avond)
        Exit Function
    End If

' Wij ID op zoeken



    Set MySheet = StartBook.Worksheets("WebInfo")
    rijteller = 2
    info = False
    Do While True
        If MySheet.Cells(rijteller, 1).Value = avond Then
            info = True
            intWebInfo = MySheet.Cells(rijteller, 2).Value
            Exit Do
        End If
          If MySheet.Cells(rijteller, 1).Value = "" Then
           VulScoreKaartInSheet = "Geen Webinfo"
           Exit Function
         End If
         rijteller = rijteller + 1
    Loop
    
    If BerekenAlleStaten = False Then
       
        SysCmd acSysCmdInitMeter, "Bereken scorestaat ...", 100
        
    End If
 
    If Data_of_Results = 1 Then
        StepUserRef = StepRef & intWebInfo & "&username="
        Else
        StepUserRef = StepRef & intWebInfo & "/"
    End If

    Set MySheet = StartBook.Worksheets("Import_Opstelling")
    
    
    Opstelling = False
        rijteller = 1
        Do While True
            If MySheet.Cells(rijteller, 1).Value = avond And MySheet.Cells(rijteller, 2) = Teamnr Then
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
    
    For i = 1 To WEDSTRIJDENPERSESSIE
        tegenst(i) = MySheet.Cells(rijteller, 6 + i).Value
        If ScorestaatIntern Then
        'bepaal zijID
            tegenID(i) = DLookup("id", "tblTeams", "[Teamnr] = " & tegenst(i) & " and [ToernooiID] =" & lngToernooi)
        End If
    Next
    
    Paar1 = Speler1 & " - " & Speler2
    Paar2 = Speler3 & " - " & Speler4
    
    refSpeler1 = StepUserRef & Speler1
    refSpeler2 = StepUserRef & Speler2
    refSpeler3 = StepUserRef & Speler3
    refSpeler4 = StepUserRef & Speler4
    
        refWij = PWORef & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & Teamnr & ".html"
    For i = 1 To WEDSTRIJDENPERSESSIE
        refTegenstander(i) = PWORef & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & tegenst(i) & ".html"
    Next
    
    
    Set MySheet = StartBook.Sheets("Teams")
    
    For i = 1 To WEDSTRIJDENPERSESSIE
    rijteller = 2
    
    Do While True
    
    
        If MySheet.Cells(rijteller, 1).Value = tegenst(i) Then
            tegen(i) = True
            Tegenstander(i) = MySheet.Cells(rijteller, 2).Value
        End If
    
    
    If MySheet.Cells(rijteller, 1).Value = Teamnr Then
        wij1 = True
        'TeamNaam
        Wij = MySheet.Cells(rijteller, 2).Value
    End If
    
    If tegen(i) And wij1 Then
        Exit Do
    End If
     If MySheet.Cells(rijteller, 1).Value = "" Then
       VulScoreKaartInSheet = "Team Niet gevonden"
       Exit Function
     End If
     rijteller = rijteller + 1
    Loop
    
    Next





Scorestaat_Speler1 = GetStepUserData_d(StepRef, intWebInfo, Speler1)
Scorestaat_Speler3 = GetStepUserData_d(StepRef, intWebInfo, Speler3)



ScoreSheetName = PREFIX & avond & "_Teamnr_" & Teamnr
strSheetName = ScoreSheetName
'opbouw ook als html bestand


     If BerekenAlleStaten = False Then
        SysCmd acSysCmdUpdateMeter, 10
     End If
         
        

' SessieID moet bekend zijn
' teamID

'


'
If AparteExcel Then

Set TemplateBook = Workbooks.Open(TemplateExcelfile)
Set TemplateSheet = TemplateBook.Worksheets("Team_Template")

Else
sheetteller = StartBook.Sheets.Count
For i = 1 To sheetteller
If i > sheetteller Then
Exit For
End If

If StartBook.Sheets(i).name = ScoreSheetName Then
    StartBook.Sheets(ScoreSheetName).Delete
    Exit For
End If
Next

sheetteller = StartBook.Sheets.Count
StartBook.Sheets("Team_Template").Copy After:=StartBook.Sheets(StartBook.Sheets.Count)
sheetteller = StartBook.Sheets.Count
StartBook.Sheets(sheetteller).name = ScoreSheetName
Set TemplateSheet = StartBook.Sheets(ScoreSheetName)

End If

   If BerekenAlleStaten = False Then
       SysCmd acSysCmdUpdateMeter, 30
   End If



TemplateSheet.Cells(1, 5).Value = Speler1 & " - " & Speler2
TemplateSheet.Cells(1, 12).Value = Speler3 & " - " & Speler4

For i = 1 To WEDSTRIJDENPERSESSIE
    TemplateSheet.Cells(5 + (i - 1) * AANTALSPELLENPERWEDSTRIJD, 20).Value = Wij
    TemplateSheet.Cells(6 + (i - 1) * AANTALSPELLENPERWEDSTRIJD, 20).Value = Tegenstander(i)
Next



'TemplateSheet.Cells(30, 3).Value = Wij


'left Column

Rijen = Split(Scorestaat_Speler1, vbCr)
rijteller = 3
kolomteller = 1

progressteller = 4

For teller = 0 To UBound(Rijen) - 1
 progressteller = progressteller + (teller + 1) * diff
    Kolommen = Split(Rijen(teller), ",")
    For teller2 = 0 To UBound(Kolommen)
        TemplateSheet.Cells(rijteller + teller, kolomteller + teller2).Value = Replace(Kolommen(teller2), Chr(34), "")
    Next
Next


'right Column

Rijen = Split(Scorestaat_Speler3, vbCr)


rijteller = 3
kolomteller = 8
For teller = 0 To UBound(Rijen) - 1
 
        
    Kolommen = Split(Rijen(teller), ",")
    For teller2 = 0 To UBound(Kolommen)
           TemplateSheet.Cells(rijteller + teller, kolomteller + teller2).Value = Replace(Kolommen(teller2), Chr(34), "")
    Next
Next

   If BerekenAlleStaten = False Then
       SysCmd acSysCmdUpdateMeter, 60
   End If
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
  For teller2 = 15 To 18
    TemplateSheet.Cells(rijteller + teller, teller2).Value = ""
  Next

End If

Next

   If BerekenAlleStaten = False Then
       SysCmd acSysCmdUpdateMeter, 90
   End If
'Vul de formules in rechterkant

'Imps = som avn alle imps

For i = 1 To WEDSTRIJDENPERSESSIE

tel1 = 5 + (i - 1) * AANTALSPELLENPERWEDSTRIJD
tel2 = 6 + (i - 1) * AANTALSPELLENPERWEDSTRIJD
tel3 = 3 + (i - 1) * AANTALSPELLENPERWEDSTRIJD
tel4 = 3 + i * AANTALSPELLENPERWEDSTRIJD

strFormula = "=SUM(Q" & tel3 & ":Q" & tel4 & ")"
TemplateSheet.Cells(tel1, 21).Formula = strFormula
strFormula = "=SUM(R" & tel3 & ":R" & tel4 & ")"
TemplateSheet.Cells(tel2, 21).Formula = strFormula

Next

'tel het aantal spellen dat gespeeld

'eerste wedstrijd

For i = 1 To WEDSTRIJDENPERSESSIE
    AantalSpelGespeeldWedstrijd(i) = 0
Next

For i = 1 To WEDSTRIJDENPERSESSIE
    For teller = (i - 1) * AANTALSPELLENPERWEDSTRIJD To i * AANTALSPELLENPERWEDSTRIJD - 1
        If TemplateSheet.Cells(rijteller + teller, kolomscorespeler1).Value = "" Then
            AantalSpelGespeeldWedstrijd(i) = AantalSpelGespeeldWedstrijd(i) + 1
        End If
    Next
Next

For i = 1 To WEDSTRIJDENPERSESSIE
     AantalSpelGespeeldWedstrijd(i) = AANTALSPELLENPERWEDSTRIJD - AantalSpelGespeeldWedstrijd(i)
Next

For i = 1 To WEDSTRIJDENPERSESSIE
Select Case AantalSpelGespeeldWedstrijd(i)
Case Is < 6
VPschaal(i) = 3
Case 6
VPschaal(i) = 3
Case 7
VPschaal(i) = 4
Case 8
VPschaal(i) = 5
Case 9
VPschaal(i) = 6
Case 10
VPschaal(i) = 7
Case Else
VPschaal(i) = 2
End Select

tel1 = 5 + (i - 1) * AANTALSPELLENPERWEDSTRIJD
tel2 = 6 + (i - 1) * AANTALSPELLENPERWEDSTRIJD
tel3 = 3 + (i - 1) * AANTALSPELLENPERWEDSTRIJD
tel4 = 3 + i * AANTALSPELLENPERWEDSTRIJD - 1

strFormula = "=IF(V" & tel1 & ">0,VLOOKUP(V" & tel1 & ",VPSchaal," & VPschaal(i) & "),20-VLOOKUP(V" & tel2 & ",VPSchaal," & VPschaal(i) & "))"
TemplateSheet.Cells(tel1, 23).Formula = strFormula
strFormula = "=IF(V" & tel2 & ">0,VLOOKUP(V" & tel2 & ",VPSchaal," & VPschaal(i) & "),20-VLOOKUP(V" & tel1 & ",VPSchaal," & VPschaal(i) & "))"
TemplateSheet.Cells(tel2, 23).Formula = strFormula
strFormula = "=SUM(Q" & tel3 & ":Q" & tel4 & ")"
TemplateSheet.Cells(tel1, 21).Formula = strFormula
strFormula = "=SUM(R" & tel3 & ":R" & tel4 & ")"
TemplateSheet.Cells(tel2, 21).Formula = strFormula


Next
 If BerekenAlleStaten = False Then
       SysCmd acSysCmdUpdateMeter, 95
 End If

'pas formules aan =ALS(V5>0;VERT.ZOEKEN(V5;VPSchaal;2);20-VERT.ZOEKEN(V6;VPSchaal;2))
For i = 1 To WEDSTRIJDENPERSESSIE

WijVPs(i) = TemplateSheet.Cells(5 + (i - 1) * AANTALSPELLENPERWEDSTRIJD, 23).Value
ZijVPs(i) = TemplateSheet.Cells(6 + (i - 1) * AANTALSPELLENPERWEDSTRIJD, 23).Value
WijImps(i) = TemplateSheet.Cells(5 + (i - 1) * AANTALSPELLENPERWEDSTRIJD, 21).Value
ZijImps(i) = TemplateSheet.Cells(6 + (i - 1) * AANTALSPELLENPERWEDSTRIJD, 21).Value

Next
    If BerekenAlleStaten = False Then
      SysCmd acSysCmdRemoveMeter
    End If
  
End If



'maak html
If ScorestaatExcel Then

If intUitvoerNaarHTML Then


ScoreSheetNameHTML = ScoreSheetName & ".html"
htmlfile = HTMLFolder & ScoreSheetNameHTML

ScoreSheetHTML = ""

'header

ScoreSheetHTML = ScoreSheetHTML & html_header()
ScoreSheetHTML = ScoreSheetHTML & html_Begin_Body()


' Kopje Scorestaat

If Prefixkopjesscorestaat = "" Then
    Kopje_prefix = "Scorekaart Viertal " & Wij & " (avond) " & avond & " "
Else
    Kopje_prefix = Prefixkopjesscorestaat
End If


If Suffixkopjesscorestaat = "" Then
    Kopje_suffix = ""
    Else
    Kopje_suffix = Suffixkopjesscorestaat
End If


ScoreSheetHTML = ScoreSheetHTML & rij_header(Kopje_prefix & Kopje_suffix)

'ScoreSheetHTML = ScoreSheetHTML & rij_Paren(Paar1, Paar2, refSpeler1, refSpeler3)

ScoreSheetHTML = ScoreSheetHTML & begin_kolommen()

ScoreSheetHTML = ScoreSheetHTML & Scoresheetheader(Paar1, refSpeler1)


For rijteller = 3 To 2 + WEDSTRIJDENPERSESSIE * AANTALSPELLENPERWEDSTRIJD
'links speler1
    strE = TemplateSheet.Cells(rijteller, 1)
    strF = TemplateSheet.Cells(rijteller, 2)
    strG = TemplateSheet.Cells(rijteller, 3)
    strH = TemplateSheet.Cells(rijteller, 4)
    strI = TemplateSheet.Cells(rijteller, 5)
    'strJ = TemplateSheet.Cells(rijteller, 6)
    
    If (rijteller > 3) And rijteller Mod (AANTALSPELLENPERWEDSTRIJD) = 2 Then
        ScoreSheetHTML = ScoreSheetHTML & ScoresheetLastRow(strE, strF, strG, strH, strI)
     Else
        ScoreSheetHTML = ScoreSheetHTML & ScoresheetRow(strE, strF, strG, strH, strI)
    End If
Next


ScoreSheetHTML = ScoreSheetHTML & Scorefooter()

ScoreSheetHTML = ScoreSheetHTML & Scoresheetheader(Paar2, refSpeler3)
'midden speler3


For rijteller = 3 To 2 + WEDSTRIJDENPERSESSIE * AANTALSPELLENPERWEDSTRIJD
    strE = TemplateSheet.Cells(rijteller, 8)
    strF = TemplateSheet.Cells(rijteller, 9)
    strG = TemplateSheet.Cells(rijteller, 10)
    strH = TemplateSheet.Cells(rijteller, 11)
    strI = TemplateSheet.Cells(rijteller, 12)
    'strJ = TemplateSheet.Cells(rijteller, 13)
    If (rijteller > 3) And rijteller Mod (AANTALSPELLENPERWEDSTRIJD) = 2 Then
       ScoreSheetHTML = ScoreSheetHTML & ScoresheetLastRow(strE, strF, strG, strH, strI)
    Else
      ScoreSheetHTML = ScoreSheetHTML & ScoresheetRow(strE, strF, strG, strH, strI)
    End If
'
Next

ScoreSheetHTML = ScoreSheetHTML & Scorefooter()

ScoreSheetHTML = ScoreSheetHTML & ScoreSaldoheader()
'rechts saldi
For rijteller = 3 To 2 + WEDSTRIJDENPERSESSIE * AANTALSPELLENPERWEDSTRIJD
    strE = TemplateSheet.Cells(rijteller, 15)
    strF = TemplateSheet.Cells(rijteller, 16)
    strG = TemplateSheet.Cells(rijteller, 17)
    strH = TemplateSheet.Cells(rijteller, 18)
     If (rijteller > 3) And rijteller Mod (AANTALSPELLENPERWEDSTRIJD) = 2 Then
   
    ScoreSheetHTML = ScoreSheetHTML & SaldoLastRow(strE, strF, strG, strH)
    Else
    ScoreSheetHTML = ScoreSheetHTML & SaldoRow(strE, strF, strG, strH)
    End If
Next

ScoreSheetHTML = ScoreSheetHTML & Scorefooter()
ScoreSheetHTML = ScoreSheetHTML & eind_kolommen()


'uitslag wedstrijd 1

'uitslag wedstrijd 2

ScoreSheetHTML = ScoreSheetHTML & rij_Lege_regel

strRijheader = "Gesp. "

For i = 1 To WEDSTRIJDENPERSESSIE
    strRijheader = strRijheader & 1 + (i - 1) * AANTALSPELLENPERWEDSTRIJD & "-" & i * AANTALSPELLENPERWEDSTRIJD & " -> " & Tegenstander(i) & " "
Next

ScoreSheetHTML = ScoreSheetHTML & rij_header(strRijheader)

ScoreSheetHTML = ScoreSheetHTML & "<tr><td>" & vbCr
ScoreSheetHTML = ScoreSheetHTML & TeamResultheader()

For i = 1 To WEDSTRIJDENPERSESSIE
    ScoreSheetHTML = ScoreSheetHTML & TeamResultRow(Tegenstander(i), refTegenstander(i), WijImps(i), ZijImps(i), WijVPs(i), ZijVPs(i))
Next

ScoreSheetHTML = ScoreSheetHTML & TeamResultfooter()
ScoreSheetHTML = ScoreSheetHTML & "</td></tr>" & vbCr
ScoreSheetHTML = ScoreSheetHTML & html_Einde_Body_scoresheet()

If fnExists(htmlfile) Then
    Kill (htmlfile)
End If





Set fso = CreateObject("Scripting.FileSystemObject")

Set oFile = fso.CreateTextFile(htmlfile)
oFile.Write ScoreSheetHTML
oFile.Close
Set fso = Nothing
Set oFile = Nothing


   End If
End If

If ScorestaatIntern Then
    If intUitvoerNaarHTML Then
    
    ScoreSheetNameHTML = ScoreSheetName & ".html"
    htmlfile = HTMLFolder & ScoreSheetNameHTML
    
    ScoreSheetHTML = ""
    
    'header
    
    ScoreSheetHTML = ScoreSheetHTML & html_header()
    ScoreSheetHTML = ScoreSheetHTML & html_Begin_Body()
    
    
    ' Kopje Scorestaat
    
    If Prefixkopjesscorestaat = "" Then
        Kopje_prefix = "Scorekaart Viertal " & Wij & " (avond) " & avond & " "
    Else
        Kopje_prefix = Prefixkopjesscorestaat
    End If
    
    
    If Suffixkopjesscorestaat = "" Then
        Kopje_suffix = ""
        Else
        Kopje_suffix = Suffixkopjesscorestaat
    End If
    
    Set db = CurrentDb
    
    sql = ""
    sql = sql & "select * from tblScorestaat where "
    sql = sql & "[TeamID] = " & WijID
    sql = sql & " and "
    sql = sql & "[ToernooiID] = " & lngToernooi
    sql = sql & " and "
    sql = sql & "[SessieID] = " & lngSessie
    sql = sql & " Order by wedstrijd, spelnr "
    
    
    
    Set rs = db.OpenRecordset(sql)
    
    ScoreSheetHTML = ScoreSheetHTML & rij_header(Kopje_prefix & Kopje_suffix)
    
    'ScoreSheetHTML = ScoreSheetHTML & rij_Paren(Paar1, Paar2, refSpeler1, refSpeler3)
    
    ScoreSheetHTML = ScoreSheetHTML & begin_kolommen()
    
    ScoreSheetHTML = ScoreSheetHTML & Scoresheetheader(Paar1, refSpeler1)
    
    rs.MoveFirst
    rijteller = 3
    Do While Not rs.EOF
    'links speler1
        strE = rs!Spelnr
        strF = rs!Contract1
        strG = rs!Resultaat1
        strH = rs!Door1
        strI = rs!score1
        'strJ = TemplateSheet.Cells(rijteller, 6)
        
        If (rijteller > 3) And rijteller Mod (AANTALSPELLENPERWEDSTRIJD) = 2 Then
            ScoreSheetHTML = ScoreSheetHTML & ScoresheetLastRow(strE, strF, strG, strH, strI)
         Else
            ScoreSheetHTML = ScoreSheetHTML & ScoresheetRow(strE, strF, strG, strH, strI)
        End If
    rijteller = rijteller + 1
    rs.MoveNext
    Loop
    
    ScoreSheetHTML = ScoreSheetHTML & Scorefooter()
    
    ScoreSheetHTML = ScoreSheetHTML & Scoresheetheader(Paar2, refSpeler3)
    'midden speler3
    rijteller = 3
    rs.MoveFirst
    Do While Not rs.EOF
    'links speler1
        strE = rs!Spelnr
        strF = rs!Contract2
        strG = rs!Resultaat2
        strH = rs!Door2
        strI = rs!score2
    
        'strJ = TemplateSheet.Cells(rijteller, 13)
        If (rijteller > 3) And rijteller Mod (AANTALSPELLENPERWEDSTRIJD) = 2 Then
           ScoreSheetHTML = ScoreSheetHTML & ScoresheetLastRow(strE, strF, strG, strH, strI)
        Else
          ScoreSheetHTML = ScoreSheetHTML & ScoresheetRow(strE, strF, strG, strH, strI)
        End If
    rijteller = rijteller + 1
    rs.MoveNext
    Loop
    
    ScoreSheetHTML = ScoreSheetHTML & Scorefooter()
    
    ScoreSheetHTML = ScoreSheetHTML & ScoreSaldoheader()
    'rechts saldi
    rijteller = 3
    rs.MoveFirst
    Do While Not rs.EOF
        strE = rs!saldo
        strF = rs!imps
        strG = rs!WijImps
        strH = rs!ZijImps
         If (rijteller > 3) And rijteller Mod (AANTALSPELLENPERWEDSTRIJD) = 2 Then
       
        ScoreSheetHTML = ScoreSheetHTML & SaldoLastRow(strE, strF, strG, strH)
        Else
        ScoreSheetHTML = ScoreSheetHTML & SaldoRow(strE, strF, strG, strH)
        End If
        rijteller = rijteller + 1
        rs.MoveNext
    Loop
    ScoreSheetHTML = ScoreSheetHTML & Scorefooter()
    ScoreSheetHTML = ScoreSheetHTML & eind_kolommen()
    
    
    'uitslag wedstrijd 1
    
    'uitslag wedstrijd 2
    
    ScoreSheetHTML = ScoreSheetHTML & rij_Lege_regel
    
    strRijheader = "Gesp. "
    
    For i = 1 To WEDSTRIJDENPERSESSIE
        strRijheader = strRijheader & 1 + (i - 1) * AANTALSPELLENPERWEDSTRIJD & "-" & i * AANTALSPELLENPERWEDSTRIJD & " -> " & Tegenstander(i) & " "
    Next
    
    ScoreSheetHTML = ScoreSheetHTML & rij_header(strRijheader)
    
    ScoreSheetHTML = ScoreSheetHTML & "<tr><td>" & vbCr
    ScoreSheetHTML = ScoreSheetHTML & TeamResultheader()
    
    For i = 1 To WEDSTRIJDENPERSESSIE
        ScoreSheetHTML = ScoreSheetHTML & TeamResultRow(Tegenstander(i), refTegenstander(i), WijImps(i), ZijImps(i), WijVPs(i), ZijVPs(i))
    Next
    
    ScoreSheetHTML = ScoreSheetHTML & TeamResultfooter()
    ScoreSheetHTML = ScoreSheetHTML & "</td></tr>" & vbCr
    ScoreSheetHTML = ScoreSheetHTML & html_Einde_Body_scoresheet()
    
    If fnExists(htmlfile) Then
        Kill (htmlfile)
    End If
    
    
    
    
   
    Set fso = CreateObject("Scripting.FileSystemObject")
   
    Set oFile = fso.CreateTextFile(htmlfile)
    oFile.Write ScoreSheetHTML
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

    End If
End If





If AparteExcel Then

TemplateSheet.name = ScoreSheetName
TemplateBook.SaveAs ScoresheetExcelfile
TemplateBook.Close

Else

Set TemplateSheet = Nothing

End If



If ScorestaatExcel Then
'Vul in de kruistabel
    Set MySheet = StartBook.Sheets("Kruistabel")
    For i = 1 To WEDSTRIJDENPERSESSIE
        MySheet.Cells(Teamnr + 1, tegenst(i) + 1).Formula = Replace("=Value(" & Format(WijVPs(i), "#0.##") & ")", ",", ".")
        MySheet.Cells(tegenst(i) + 1, Teamnr + 1).Formula = Replace("=Value(" & Format(ZijVPs(i), "#0.##") & ")", ",", ".")
    Next

End If


'Vul in de uitslagen in


If ScorestaatExcel Then

    Set MySheet = StartBook.Sheets("TeamUitslagen")
    
    offsetsessie = GespeeldeWedstrijden(avond, lngToernooi)
    
    
    ' zoek of het team thuis spelen is in wedstrijd 1
    
    
    For i = 1 To WEDSTRIJDENPERSESSIE
        Wedstrijd1 = False
        rijteller = 2
        Do While MySheet.Cells(rijteller, 1).Value <> ""
            If (MySheet.Cells(rijteller, 1).Value = avond) And (MySheet.Cells(rijteller, 2).Value = i + offsetsessie) And (MySheet.Cells(rijteller, 3).Value = Teamnr) Then
             Wedstrijd1 = True
             Exit Do
            End If
            rijteller = rijteller + 1
        Loop
        
        If Wedstrijd1 And tegenst(i) <> TEAMBYE Then
                MySheet.Cells(rijteller, 7).Value = WijImps(i)
                MySheet.Cells(rijteller, 8).Value = ZijImps(i)
                'MySheet.Cells(rijteller, 9).Formula = Replace("=Value(" & Format(WijVPs(i), "#0.##") & ")", ",", ".")
                'MySheet.Cells(rijteller, 10).Formula = Replace("=Value(" & Format(ZijVPs(i), "#0.##") & ")", ",", ".")
                MySheet.Cells(rijteller, 9).Value = WijVPs(i)
                MySheet.Cells(rijteller, 10).Value = ZijVPs(i)
        End If
    Next
    
    
        StartBook.Save
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
End If


End Function



Public Function fnExists(FileName) As Integer
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(FileName) Then
    fnExists = True
    Else
    fnExists = False
    End If
Set fs = Nothing

Exit Function
error_fnexists:
Set fs = Nothing

fnExists = False
End Function
Public Function fnCopyfile(FileName As Variant, destination As Variant) As Integer
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")
' First parameter: original location\file
' Second parameter: new location\file
objFSO.CopyFile FileName, destination
Set objFSO = Nothing
End Function
Function plus_add(number As Variant) As String
If IsNull(number) Then
    plus_add = "&nbsp;"
    Exit Function
End If
If number = "" Then
    plus_add = "&nbsp;"
    Exit Function
End If

If Left(number, 1) = "+" Then
    plus_add = number
     Exit Function
End If

If Left(number, 1) = "-" Then
    plus_add = number
     Exit Function
End If
If Val(number) > 0 Then
    plus_add = "+" & number
    Exit Function
End If

plus_add = number


End Function
Public Function GetStepUserData_d(sUrl As Variant, ActiveID As Variant, User As Variant) As String
    Dim s, a, strTable As String
    Dim Scorestaat As String
    Dim intDoor As Integer
    Dim Rijen_Uitslag() As String
    Dim Kolommen_Uitslag() As String
    Dim aantalkolommen As Integer
    Dim tel_tbody_1, tel_tbody_2 As Long
    Dim Spelnr, Contract, resultaat, Door, score, ImpsButler, kleur As String
    
    'variant geschikt voor beide sides
    
    ' GetHTMLFromURL(sUrl & ActiveID & "&" & "username=" & User)
   If sUrl = STEPDATA Then
    s = GetHTMLFromURL(STEPDATA & ActiveID & "&username=" & User)
    intDoor = True
   Else
    s = GetHTMLFromURL(STEPRESULTS & ActiveID & "/" & User)
    intDoor = False
   End If
   
   
    'zoek body
 
    tel_tbody_1 = InStr(s, "<tbody>")
    tel_tbody_2 = InStr(tel_tbody_1 + 7, s, "</tbody>")
    strTable = Mid(s, tel_tbody_1, tel_tbody_2 - tel_tbody_1 + 8)
    
    
    a = grap_table(strTable)
    
    a = extract_kleur(a)
    
    Rijen_Uitslag = grap_rijen(a)
  
  For tel_1 = LBound(Rijen_Uitslag) To UBound(Rijen_Uitslag) - 1
    
        Kolommen_Uitslag = grap_cellen(Rijen_Uitslag(tel_1))
        
        aantalkolommen = UBound(Kolommen_Uitslag)
        Spelnr = ""
        Contract = ""
        resultaat = ""
        Door = ""
        score = ""
        ImpsButler = ""
        
        Spelnr = stripTags(Kolommen_Uitslag(0))
        Contract = stripTags(Kolommen_Uitslag(1))
        
        
        'test op niet gespeeld of kunstmatig
        If aantalkolommen < 4 Then
          ImpsButler = stripTags(Kolommen_Uitslag(2))
        'ImpsButler = Replace(ImpsButler, "IMP", " IMP")
        Else
            If Not intDoor Then
                resultaat = stripTags(Kolommen_Uitslag(2))
                score = stripTags(Kolommen_Uitslag(3))
                ImpsButler = stripTags(Kolommen_Uitslag(4))
                'ImpsButler = Replace(ImpsButler, "IMP", " IMP")
            Else
                resultaat = stripTags(Kolommen_Uitslag(2))
                Door = stripTags(Kolommen_Uitslag(3))
                score = stripTags(Kolommen_Uitslag(4))
                ImpsButler = stripTags(Kolommen_Uitslag(5))
                'ImpsButler = Replace(ImpsButler, "IMP", " IMP")
            End If
        End If
        Scorestaat = Scorestaat & Spelnr & "," & Contract & "," & resultaat & "," & Door & "," & score & "," & ImpsButler & vbCr
        
    Next
    
    GetStepUserData_d = Scorestaat
     
End Function
Public Function ZoekWebsite(van As Variant, Tot As Variant) As String
   Dim url, s, a, b, c As String
    
    url = STEPRESULTS
    
    For i = van To Tot
        s = GetHTMLFromURL(url & i)
        x = InStr(s, "Datum:")
        If x > 0 Then
        Debug.Print i
       Debug.Print Mid(s, x, 50)
       End If
        
    Next
End Function


Public Sub NogGeenSheetsHTML(avond As Variant, Teamnr As Variant)

Dim offset_speler, intWedstrijd, Wedstrijd1, Wedstrijd2 As Integer
Dim rijteller As Integer
Dim kolomteller, sheetteller As Integer
Dim Opstelling, AparteExcel As Integer
Dim intWebInfo, info As Integer
Dim Speler1, Speler3, Speler2, Speler4, Paar1, Paar2 As String
Dim refSpeler1, refSpeler3, refSpeler2, refSpeler4 As String
Dim StepRef, StepUserRef As String
Dim PWORef, PWOTeamRef As String
Dim Wij, Tegenstander1, Tegenstander2 As String
Dim refWij, refTegenstander1, refTegenstander2 As String
Dim strE, strF, strG, strH, strI, strJ As String

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
Dim WijScore1, WijScore2, ZijScore1, ZijScore2, WijImps1, ZijImps1, WijImps2, ZijImps2, WijVPs1, ZijVPs1, WijVPs2, ZijVPs2 As Double
Dim xlApp As Object
Dim Rekenkamerfolder, HTMLFolder As String
Dim Data_of_Results As Integer

Rekenkamerfolder = "C:\Users\pgjmw\Dropbox\DonderdagAvond\Rekenkamer"
HTMLFolder = LOCALHTML
Data_of_Results = 2
If Data_of_Results = 1 Then
    StepRef = STEPDATA
Else
    StepRef = STEPRESULTS
End If

PWORef = LOCALSITE


Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)

Set MySheet = StartBook.Worksheets("WebInfo")
rijteller = 2
info = False
Do While True
    If MySheet.Cells(rijteller, 1).Value = avond Then
        info = True
        intWebInfo = MySheet.Cells(rijteller, 2).Value
        Exit Do
    End If
      If MySheet.Cells(rijteller, 1).Value = "" Then
        Exit Sub
     End If
     rijteller = rijteller + 1
Loop


If Data_of_Results = 1 Then
    StepUserRef = StepRef & intWebInfo & "&username="
Else
    StepUserRef = StepRef & intWebInfo & "/"
End If



Set MySheet = StartBook.Worksheets("Import_Opstelling")


Opstelling = False
    rijteller = 1
    Do While True
        If MySheet.Cells(rijteller, 1).Value = avond And MySheet.Cells(rijteller, 2) = Teamnr Then
         Opstelling = True
         Exit Do
        End If
        If MySheet.Cells(rijteller, 1).Value = "" Then
            
            Exit Sub
         End If
    rijteller = rijteller + 1
    Loop

Speler1 = MySheet.Cells(rijteller, 3).Value
Speler2 = MySheet.Cells(rijteller, 4).Value
Speler3 = MySheet.Cells(rijteller, 5).Value
Speler4 = MySheet.Cells(rijteller, 6).Value
tegenst1 = MySheet.Cells(rijteller, 7).Value
tegenst2 = MySheet.Cells(rijteller, 8).Value


Paar1 = Speler1 & " - " & Speler2
Paar2 = Speler3 & " - " & Speler4

refSpeler1 = StepUserRef & Speler1
refSpeler2 = StepUserRef & Speler2
refSpeler3 = StepUserRef & Speler3
refSpeler4 = StepUserRef & Speler4

refWij = PWORef & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & Teamnr & ".html"
refTegenstander1 = PWORef & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & tegenst1 & ".html"
refTegenstander2 = PWORef & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & tegenst2 & ".html"



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
  
   Exit Sub
 End If
 rijteller = rijteller + 1
Loop


ScoreSheetName = PREFIX & avond & "_Teamnr_" & Teamnr
'opbouw ook als html bestand

ScoreSheetNameHTML = ScoreSheetName & ".html"
htmlfile = HTMLFolder & ScoreSheetNameHTML

ScoreSheetHTML = ""


ScoreSheetHTML = ScoreSheetHTML & html_header()
ScoreSheetHTML = ScoreSheetHTML & html_Begin_Body()

ScoreSheetHTML = ScoreSheetHTML & rij_header("Avond " & avond & ", nog geen resultaten viertal " & Wij)

ScoreSheetHTML = ScoreSheetHTML & rij_Lege_regel()
ScoreSheetHTML = ScoreSheetHTML & rij_header("Speelt spellen 1-12 tegen " & Tegenstander1 & " en spellen 13-24 tegen " & Tegenstander2)

ScoreSheetHTML = ScoreSheetHTML & "<tr><td>" & vbCr
ScoreSheetHTML = ScoreSheetHTML & TeamResultheader()
ScoreSheetHTML = ScoreSheetHTML & TeamNoResultRow(Tegenstander1, refTegenstander1)
ScoreSheetHTML = ScoreSheetHTML & TeamNoResultRow(Tegenstander2, refTegenstander2)
ScoreSheetHTML = ScoreSheetHTML & TeamResultfooter()
ScoreSheetHTML = ScoreSheetHTML & "</td></tr>" & vbCr

ScoreSheetHTML = ScoreSheetHTML & html_Einde_Body()



If fnExists(htmlfile) Then
    Kill (htmlfile)
End If


Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile(htmlfile)
oFile.Write ScoreSheetHTML
oFile.Close
Set fso = Nothing
Set oFile = Nothing

    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub
Public Sub NogGeenUitslagenIn(ToernID As Variant, avond As Variant)
Dim ViertalUitslag, WebInf, StepRef, ScoreSheetHTML, UitslagHTML, UitslagenHTML, HTMLFolder As String
Dim rijteller, teller, teller2, kolomteller, kolomteller2, kolomoffset, i, j As Integer
Dim intWebInfo As Integer

Dim Thuis, Uit As Variant
Dim refThuis, refUit As Variant
Dim Imps_Thuis As Variant
Dim Imps_Uit As Variant
Dim VPs_Thuis As Variant
Dim VPs_Uit As Variant

Dim MySheet As Object
Dim StartBook As Object
Dim Rijen() As String
Dim Kolommen() As String
Dim xlApp As Object
Dim db As Database
Dim rs As Recordset
        
 
        Set db = CurrentDb
        Set rs = db.OpenRecordset("select * from tblSessie where [ToernooiD] = " & ToernID & " and [Sessienr] = " & avond)
        If Not (rs.BOF And rs.EOF) Then
            lngSessie = rs!id
            lngToernooi = rs!ToernooID
        Else
            rs.Close
            db.Close
            Exit Sub
        End If
        rs.Close
        db.Close
        'fris alle gegevens op
        lngSessieOld = 0
        lngToernooiOld = 0
        Call InitAll(lngToernooi, lngSessie)
 

Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets("WebInfo")

rijteller = 2
Do While True
If MySheet.Cells(rijteller, 1).Value = avond Then
    info = True
    intWebInfo = MySheet.Cells(rijteller, 2).Value
    Exit Do
End If


 If MySheet.Cells(rijteller, 1).Value = "" Then
   Exit Sub
 End If
 rijteller = rijteller + 1
Loop

'ViertalUitslag = GetStepViertalUitslag(STEPDATA, WebInfo)

'Rijen = Split(Uitslag, vbCr)
HTMLFolder = LOCALHTML
UitslagHTML = HTMLFolder & "Uitslagen" & PREFIX & avond & ".html"

UitslagenHTML = ""
UitslagenHTML = UitslagenHTML & html_header()
UitslagenHTML = UitslagenHTML & html_Begin_Body()
UitslagenHTML = UitslagenHTML & rij_header("Nog geen uitslagen avond !!!" & avond)

UitslagenHTML = UitslagenHTML & TeamUitslagenResultheader()

Set MySheet = StartBook.Sheets("TeamUitslagen")

rijteller = 2

Do While MySheet.Cells(rijteller, 1).Value <> ""
    
    If MySheet.Cells(rijteller, 1).Value = avond Then
    
        'plot regel
        
        Thuis = MySheet.Cells(rijteller, 5).Value
        Uit = MySheet.Cells(rijteller, 6).Value
        Imps_Thuis = MySheet.Cells(rijteller, 7).Value
        Imps_Uit = MySheet.Cells(rijteller, 8).Value
        VPs_Thuis = MySheet.Cells(rijteller, 9).Value
        VPs_Uit = MySheet.Cells(rijteller, 10).Value
        refThuis = LOCALSITE & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & MySheet.Cells(rijteller, 3).Value & ".html"
        refUit = LOCALSITE & intWebInfo & "/" & PREFIX & avond & "_Teamnr_" & MySheet.Cells(rijteller, 4).Value & ".html"
        
        If Thuis = "Bye" Then
            refThuis = refUit
            Imps_Thuis = "&nbsp;"
            Imps_Uit = "&nbsp;"
            VPs_Thuis = "&nbsp;"
            VPs_Uit = "&nbsp;"
        End If
        
        If Uit = "Bye" Then
            refUit = refThuis
            Imps_Thuis = "&nbsp;"
            Imps_Uit = "&nbsp;"
            VPs_Thuis = "&nbsp;"
            VPs_Uit = "&nbsp;"
        End If
        
        UitslagenHTML = UitslagenHTML & TeamUitslagenNoResultRow(Thuis, refThuis, Uit, refUit)
        
    End If
    rijteller = rijteller + 1
Loop
    
UitslagenHTML = UitslagenHTML & TeamResultfooter()
ScoreSheetHTML = ScoreSheetHTML & html_Einde_Body()

If fnExists(UitslagHTML) Then
    Kill (UitslagHTML)
End If




Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile(UitslagHTML)
oFile.Write UitslagenHTML
oFile.Close
Set fso = Nothing
Set oFile = Nothing

    
    

    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub


Public Function BepaalImps(score As Variant) As Integer
Dim db As Database
Dim rs As Recordset
Set db = CurrentDb
Dim imps As Integer

Set rs = db.OpenRecordset("imps")
rs.MoveFirst
Do While Not rs.EOF

If score > rs!Verschil Then
    imps = rs!imps
    rs.MoveNext
    
    Else
     If score < rs!Verschil Then
        BepaalImps = imps
     Else
        BepaalImps = rs!imps
     End If
        rs.Close
        db.Close
        Exit Function
    End If
Loop
BepaalImps = 24
  rs.Close
        db.Close
End Function


Public Function BepaalVPs(imps As Variant, Aantalspellen As Variant) As Double
Dim db As Database
Dim rs As Recordset
Set db = CurrentDb
Dim abs_imps As Integer
Dim VPs As Double

If imps < 0 Then
    BepaalVPs = 0
    Else
    BepaalVPs = 20
End If

Set rs = db.OpenRecordset("vps")
'laad tabel
Do While Not rs.EOF
 If rs.Fields("imps") = Abs(imps) Then
    VPs = rs.Fields("vps_" & Aantalspellen)
    If imps < 0 Then
    BepaalVPs = 20 - VPs
    Else
    BepaalVPs = VPs
    End If
    
    rs.Close
    db.Close
    Exit Function
 End If
 rs.MoveNext
Loop
  rs.Close
  db.Close
End Function