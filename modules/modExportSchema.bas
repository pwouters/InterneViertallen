Option Compare Text

Public Function ExportStepSchema(avond As Variant) As String
Dim StartBook As Object
Dim IndelingSheet As Object
Dim SchemaSheet As Object

Dim TeamThuis, TeamUit As Integer
Dim ThuisSpeler1, UitSpeler1, ThuisSpeler2, UitSpeler3, ByeSpeler1, ByeSpeler3 As String
Dim teller, teller2, teller3, Tafel, AantalTafels  As Integer
Dim schemarijteller As Integer
Dim offsetopstelling As Integer
Dim StepIndelingRonde() As String

If avond = "" Then
ExportStepSchema = ""
Exit Function
End If
' wedstrijd = 1  halve competitie
' wedstrijd = 2  hele competitie
ReDim StepIndelingRonde(WEDSTRIJDENPERSESSIE * WEDSTRIJD)
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set IndelingSheet = StartBook.Sheets("Import_Opstelling")
Set SchemaSheet = StartBook.Sheets("Schema")

sessieoffset = 2

'bepaal offset schema dus welke wedstrijd ronde

'sessienr = 1 dan offset = 2
'indien sessie> 1 then
'berekenaantalwedstrijden gespeeld
If Sessienr > 1 Then
For i = 1 To Sessienr - 1
 sessieoffset = sessieoffset + DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & lngToernooi & " and  [sessienr] = " & i)
Next
End If

'kan dus roundrobin zijn maar ook  zwitsers



    '   regel 1 Avond   Teamnr  Speler1 Speler2 Speler3 Speler4 Wedstrijd1  Wedstrijd2
    '   regel 2 1       1       ChristR BasParr HansdL  ErikOo      2           4
    '   regel 3 1       2       PauldLa Germ    BrigitC ArnoudB     1           6
    '   regel 4
    
    '   er zijn 16 teams no 16 is de bye
    '   opstelling per avond (16 regels)
    '   de opstelling van de teams bevinden zich op (Avond - 1) * 16 + 1  + teamnr
    '   speler1 = kolom 3
    '   speler3 = kolom 5
    
    
'sheet Import_Opstelling


'sheet schema

            'Paring1     Paring2     Paring3     Paring4     Paring5     Paring6     Paring7     Paring8
'Wedstrijd   Thuis1  Uit1    Thuis2  Uit2    Thuis3  Uit3    Thuis4  Uit4    Thuis5  Uit5    Thuis6  Uit6    Thuis7  Uit7    Thuis8  Uit8
'1              1   2           3   4           5   6           7   8           9   10          11  12          13  14          15  16
'2              1   4           2   6           8   3           5   9           7   12          10  13          11  15          16  14
'3              1   6           8   4           9   2           12  3           13  5           7   15          11  14          16  10


' per rij een wedstrijd wordt in twee helften gespeeld

'per avond hebben we twee wedstrijden
'wedstrijd 1 = regel



'schemarijteller = avond * WEDSTRIJD
schemarijteller = sessieoffset
'wedstrijd 2   schemarijteller  + 1 of wel avond * 2


For teller = 1 To WEDSTRIJDENPERSESSIE * WEDSTRIJD
    StepIndelingRonde(teller) = ""
    StepIndelingRonde(teller) = StepIndelingRonde(teller) & "/schedule clear " & teller & vbCrLf
    StepIndelingRonde(teller) = StepIndelingRonde(teller) & "/schedule set " & teller
Next





For intronde = 1 To WEDSTRIJDENPERSESSIE

    offsetopstelling = (avond - 1) * (AANTALTEAMS) + 1

    For teller = 1 To AANTALTEAMS \ 2

        TeamStilzit = 0
        TeamThuis = SchemaSheet.Cells(schemarijteller + intronde, teller * 2).Value
        TeamUit = SchemaSheet.Cells(schemarijteller + intronde, teller * 2 + 1).Value
        ThuisSpeler1 = IndelingSheet.Cells(offsetopstelling + TeamThuis, 3).Value
        ThuisSpeler3 = IndelingSheet.Cells(offsetopstelling + TeamThuis, 5).Value
        UitSpeler1 = IndelingSheet.Cells(offsetopstelling + TeamUit, 3).Value
        UitSpeler3 = IndelingSheet.Cells(offsetopstelling + TeamUit, 5).Value

        If TeamThuis = TEAMBYE Then
            TeamStilzit = TeamUit
            ByeSpeler1 = IndelingSheet.Cells(offsetopstelling + TeamStilzit, 3).Value
            ByeSpeler3 = IndelingSheet.Cells(offsetopstelling + TeamStilzit, 5).Value
        End If

        If TeamUit = TEAMBYE Then
            TeamStilzit = TeamThuis
            ByeSpeler1 = IndelingSheet.Cells(offsetopstelling + TeamStilzit, 3).Value
            ByeSpeler3 = IndelingSheet.Cells(offsetopstelling + TeamStilzit, 5).Value
        End If


        

            If (teller <> AANTALTEAMS \ 2) Then

            'A1 thuis1 (speler1) - Uit1 (speler1)  ..   A2 Uit2 (Speler3)  - Thuis2 (Speler3)
            'For intWedstrijd = 1 To aantalwedstrijden
                If WEDSTRIJD = ENKEL Then
                    StepIndelingRonde(intronde) = StepIndelingRonde(intronde) & ", A" & (teller * 2 - 1) & " " & ThuisSpeler1 & " " & UitSpeler1 & ", " & "A" & (teller * 2) & " " & UitSpeler3 & " " & ThuisSpeler3 & ""
                 Else
                    StepIndelingRonde((intronde - 1) * 2 + 1) = StepIndelingRonde((intronde - 1) * 2 + 1) & ", A" & (teller * 2 - 1) & " " & ThuisSpeler1 & " " & UitSpeler1 & ", " & "A" & (teller * 2) & " " & UitSpeler3 & " " & ThuisSpeler3 & ""
                    StepIndelingRonde((intronde) * 2) = StepIndelingRonde((intronde) * 2) & ", A" & (teller * 2 - 1) & " " & ThuisSpeler1 & " " & UitSpeler3 & ", " & "A" & (teller * 2) & " " & UitSpeler1 & " " & ThuisSpeler3 & ""
                End If
            Else
                If TEAMBYE <> 0 Then
                    If WEDSTRIJD = ENKEL Then
                        StepIndelingRonde(intronde) = StepIndelingRonde(intronde) & ", A" & (teller * 2 - 1) & (teller * 2 - 1) & " " & ByeSpeler1 & " " & ByeSpeler3 & vbCrLf
                    Else
                        StepIndelingRonde((intronde - 1) * 2 + 1) = StepIndelingRonde((intronde - 1) * 2 + 1) & ", A" & (teller * 2 - 1) & " " & ByeSpeler1 & " " & ByeSpeler3 & vbCrLf
                        StepIndelingRonde((intronde) * 2) = StepIndelingRonde((intronde) * 2) & ", A" & (teller * 2 - 1) & " " & ByeSpeler3 & " " & ByeSpeler1 & vbCrLf
                    End If
                 Else
                     If WEDSTRIJD = ENKEL Then
                        StepIndelingRonde(intronde) = StepIndelingRonde(intronde) & ", A" & (teller * 2 - 1) & " " & ThuisSpeler1 & " " & UitSpeler1 & ", " & "A" & (teller * 2) & " " & UitSpeler3 & " " & ThuisSpeler3 & "" & vbCrLf
                     Else
                        StepIndelingRonde((intronde - 1) * 2 - 1) = StepIndelingRonde((intronde - 1) * 2 + 1) & ", A" & (teller * 2 - 1) & " " & ThuisSpeler1 & " " & UitSpeler1 & ", " & "A" & (teller * 2) & " " & UitSpeler3 & " " & ThuisSpeler3 & "" & vbCrLf
                        StepIndelingRonde((intronde - 1) * 2) = StepIndelingRonde((intronde - 1) * 2) & ", A" & (teller * 2 - 1) & " " & ThuisSpeler1 & " " & UitSpeler3 & ", " & "A" & (teller * 2) & " " & UitSpeler1 & " " & ThuisSpeler3 & "" & vbCrLf
                    End If
                End If
            End If
    Next
Next


strSchema = ""
For teller = 1 To WEDSTRIJDENPERSESSIE * WEDSTRIJD
    strSchema = strSchema & StepIndelingRonde(teller)
Next
ExportStepSchema = strSchema

    'StartBook.Save
    Set IndelingSheet = Nothing
    Set SchemaSheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
End Function