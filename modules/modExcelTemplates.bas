Option Compare Database

' Hieronder het opzetten van een nieuwtoernooi en zijn functies



Public Sub NieuwToernooi()
    Dim db          As Database
    Dim rs          As Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblToernooi")
    rs.AddNew
    rs!ToernooiNaam = ToernooiNaam
    rs!WORKFOLDER = WORKFOLDER
    rs!WORKFILE = WORKFILE
    rs!STEPDATA = STEP_DATA
    rs!STEPRESULTS = STEP_RESULTS
    rs!AANTALSESSIES = AANTALSESSIES
    rs!WEDSTRIJDENPERSESSIE = WEDSTRIJDENPERSESSIE
    rs!WORKTEMPLATE = WORKFOLDER
    rs!LOCALHTML = LOCALHTML
    rs!LOCALHTML = LOCALSITE
    rs!PREFIX = "Wedstrijd_"
    rs.Update
    rs.Bookmark = rs.LastModified
    lngToernooi = rs!Id
    rs.Close
    
    'nu add sessie aan
    Set rs = db.OpenRecordset("tblSessie")
    
    rs.AddNew
    rs!ToernooID = lngToernooi
    rs!Sessienaam = "Eerste Sessie"
    rs!Sessienr = 1
    rs!Aantalspellen = AANTALSPELLENPERWEDSTRIJD
    rs!Competitie = 1        'halve
    rs!Prefixkopjesscorestaat = "Scorekaart van sessie 1 "
    rs!PrefixKopjeuitslagen = "Uitslag van sessie 1"
    rs!Suffixkopjesscorestaat = ""
    rs!SuffixKopjeuitslagen = ""
    rs!Voettekst = "@Bridge"
    rs!Voetlink = "#"
    rs!wedstrijdvormID = 0
    rs!ActivityID = 0
    rs!AANTALTEAMS = AANTALTEAMS
    rs!ByeTeam = False
    rs!AantalWedstrijdenPerSessie = 1
    rs.Update
    rs.Bookmark = rs.LastModified
    lngSessie = rs!Id
    
    rs.Close
    db.Close
    
    'zorgen dat alle globale variabelen goed staan
    
    Call InitAll(lngToernooi, lngSessie)
    
    'Maken van het werkbestand
    
    ' save werkbestand
    Dim xlApp       As Object
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile As String
    
    strWorkfile = WORKFOLDER & WORKFILE
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Add
    StartBook.SaveAs strWorkfile
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
    'genereren van alle sheets
    
    Call CreateImpsSheet(WORKFOLDER, WORKFILE)
    Call CreateVPsSheet(WORKFOLDER, WORKFILE)
    Call CreateWebInfoSheet(WORKFOLDER, WORKFILE)
    Call CreateTeamsSheet(AANTALTEAMS, WORKFOLDER, WORKFILE)
    Call CreateOpstellingSheet(AANTALTEAMS, WEDSTRIJDENPERSESSIE, WORKFOLDER, WORKFILE)
    Call CreateScoreTemplateSheet(WEDSTRIJDENPERSESSIE, AANTALSPELLENPERWEDSTRIJD, WORKFOLDER, WORKFILE)
    Call CreateTeamUitslagenSheet(AANTALTEAMS, WORKFOLDER, WORKFILE)
    Call CreateKruisTabelSheet(AANTALTEAMS, WORKFOLDER, WORKFILE, TEAMBYE)
    
End Sub

Public Sub CreateScoreTemplateSheet(varWedstrijdenPerSessie As Variant, varAantalSpellenPerWedstrijd, wrkFolder As Variant, wrkFile As Variant)
    Dim xlApp       As Object
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile As String
    Dim question    As Integer
    Dim fcount, i, j, start, einde As Integer
    Dim intVPKolom  As Integer
    
    Dim rng         As Range
    Dim strRange    As String
    
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(wrkFolder & wrkFile)
    
    If Not SheetExists("Team_template", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "Team_template"
    Else
        Set MySheet = StartBook.Sheets("Team_Template")
    End If
    
    MySheet.Cells.Clear
    
    'eerste rij
    
    MySheet.Cells(1, 2).Value = "Team"
    MySheet.Cells(1, 2).Interior.color = RGB(244, 176, 132)
    
    MySheet.Cells(1, 3).Interior.color = RGB(208, 206, 206)
    
    MySheet.Cells(1, 4).Value = "Paar_1"
    MySheet.Cells(1, 4).Interior.color = RGB(244, 176, 132)
    
    MySheet.Cells(1, 5).Interior.color = RGB(208, 206, 206)
    
    MySheet.Cells(1, 11).Value = "Paar_2"
    MySheet.Cells(1, 11).Interior.color = RGB(244, 176, 132)
    
    MySheet.Cells(1, 12).Interior.color = RGB(208, 206, 206)
    
    'tweede rij
    
    MySheet.Cells(2, 1).Value = "Spel"
    MySheet.Cells(2, 2).Value = "Contract"
    MySheet.Cells(2, 3).Value = "Resultaat"
    MySheet.Cells(2, 4).Value = "Door"
    MySheet.Cells(2, 5).Value = "Score"
    MySheet.Cells(2, 6).Value = "Imps_butler"
    MySheet.Cells(2, 8).Value = "Spel"
    MySheet.Cells(2, 9).Value = "Contract"
    MySheet.Cells(2, 10).Value = "Resultaat"
    MySheet.Cells(2, 11).Value = "Door"
    MySheet.Cells(2, 12).Value = "Score"
    MySheet.Cells(2, 13).Value = "Imps_butler"
    MySheet.Cells(2, 15).Value = "saldo"
    MySheet.Cells(2, 16).Value = "imps"
    MySheet.Cells(2, 17).Value = "imps wij"
    MySheet.Cells(2, 18).Value = "imps zij"
    
    Set rng = MySheet.Range(MySheet.Cells(2, 1), MySheet.Cells(2, 18))
    rng.Interior.color = RGB(180, 198, 231)
    Set rng = Nothing
    
    rijteller = 2
    For i = 1 To varWedstrijdenPerSessie
        For j = 1 To varAantalSpellenPerWedstrijd
            rijteller = rijteller + 1
            MySheet.Cells(rijteller, 15).Formula = "=E" & rijteller & " + L" & rijteller
            MySheet.Cells(rijteller, 16).Formula = "=VLOOKUP(O" & rijteller & ",Impschaal,2)"
            MySheet.Cells(rijteller, 17).Formula = "=IF(O" & rijteller & ">0,P" & rijteller & "," & Chr(34) & Chr(34) & ")"
            MySheet.Cells(rijteller, 18).Formula = "=IF(O" & rijteller & "<0,-1*P" & rijteller & "," & Chr(34) & Chr(34) & ")"
        Next
        Set rng = MySheet.Range(MySheet.Cells(2 + (i - 1) * varAantalSpellenPerWedstrijd + 1, 1), MySheet.Cells(2 + i * varAantalSpellenPerWedstrijd, 18))
        If i Mod 2 = 1 Then
            rng.Interior.color = RGB(189, 215, 238)
        Else
            rng.Interior.color = RGB(155, 194, 230)
        End If
        Set rng = Nothing
    Next
    
    'varAantalSpellenPerWedstrijd
    
    Select Case varAantalSpellenPerWedstrijd
        Case Is < 7
            intVPKolom = 3
        Case 7
            intVPKolom = 4
        Case 8
            intVPKolom = 5
        Case 9
            intVPKolom = 6
        Case 10
            intVPKolom = 7
        Case 11, 12
            intVPKolom = 2
    End Select
    
    'box
    
    'imps   = =SOM(Q3:Q9)
    
    For i = 1 To varWedstrijdenPerSessie
        start = 2 + (i - 1) * varAantalSpellenPerWedstrijd + 1
        einde = 2 + (i) * varAantalSpellenPerWedstrijd
        
        rijteller = 3 + (i - 1) * varAantalSpellenPerWedstrijd
        MySheet.Cells(rijteller, 20).Value = "Wedstrijd " & i
        MySheet.Cells(rijteller, 21).Value = "Imps"
        MySheet.Cells(rijteller, 23).Value = "VPs"
        
        MySheet.Cells(rijteller + 2, 19).Value = "Wij"
        MySheet.Cells(rijteller + 3, 19).Value = "Zij"
        MySheet.Cells(rijteller + 2, 21).Formula = "=SUM(Q" & start & ":Q" & einde & ")"
        MySheet.Cells(rijteller + 3, 21).Formula = "=SUM(R" & start & ":R" & einde & ")"
        MySheet.Cells(rijteller + 2, 22).Formula = "=U" & rijteller + 2 & " - U" & rijteller + 3
        MySheet.Cells(rijteller + 3, 22).Formula = "=U" & rijteller + 3 & " - U" & rijteller + 2
        MySheet.Cells(rijteller + 2, 23).Formula = "=IF(V" & rijteller + 2 & ">0,VLOOKUP(V" & rijteller + 2 & ",VPSchaal," & intVPKolom & "),20-VLOOKUP(V" & rijteller + 3 & ",VPSchaal," & intVPKolom & "))"
        MySheet.Cells(rijteller + 3, 23).Formula = "=IF(V" & rijteller + 3 & ">0,VLOOKUP(V" & rijteller + 3 & ",VPSchaal," & intVPKolom & "),20-VLOOKUP(V" & rijteller + 2 & ",VPSchaal," & intVPKolom & "))"
        
        Set rng = MySheet.Range(MySheet.Cells(start, 19), MySheet.Cells(start, 23))
        With rng.Borders
            .LineStyle = xlContinuous
            .color = vbBlack
            .Weight = xlThin
        End With
        Set rng = Nothing
        
        Set rng = MySheet.Range(MySheet.Cells(start + 2, 19), MySheet.Cells(start + 2, 23))
        With rng.Borders
            .LineStyle = xlContinuous
            .color = vbBlack
            .Weight = xlThin
        End With
        Set rng = Nothing
        Set rng = MySheet.Range(MySheet.Cells(start + 3, 19), MySheet.Cells(start + 3, 23))
        With rng.Borders
            .LineStyle = xlContinuous
            .color = vbBlack
            .Weight = xlThin
        End With
        Set rng = Nothing
        
        Set rng = MySheet.Range(MySheet.Cells(start, 19), MySheet.Cells(start + 3, 23))
        With rng.BorderAround(xlContinuous, xlThin)
        End With
        Set rng = Nothing
    Next i
    
    StartBook.Save
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub CreateImpsSheet(wrkFolder As Variant, wrkFile As Variant)
    Dim xlApp       As Object
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile, MyRangeName As String
    Dim db          As Database
    Dim rs          As Recordset
    Dim i, j        As Integer
    
    Dim rng         As Range
    Dim strRange    As String
    
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(wrkFolder & wrkFile)
    
    If Not SheetExists("Imptabel", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "Imptabel"
    Else
        MsgBox ("Reeds aanwezig")
        Exit Sub
    End If
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("select * from imps Order by Verschil")
    rs.MoveFirst
    i = 1
    Do While Not rs.EOF
        MySheet.Cells(i, 1).Value = rs!Verschil
        MySheet.Cells(i, 2).Value = rs!imps
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    db.Close
    
    'creeer rangename
    Set rng = MySheet.Range(MySheet.Cells(1, 1), MySheet.Cells(i - 1, 2))
    'specify defined name
    
    MyRangeName = "Impschaal"
    
    StartBook.Names.Add name:=MyRangeName, RefersTo:=rng
    'create named range with workbook scope. Defined name and cell range are as specified
    
    StartBook.Save
    Set rng = Nothing
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub CreateVPsSheet(wrkFolder As Variant, wrkFile As Variant)
    Dim xlApp       As Object
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile, MyRangeName As String
    Dim db          As Database
    Dim rs          As Recordset
    Dim i, j        As Integer
    
    Dim rng         As Range
    Dim strRange    As String
    
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        'Creer een nieuwe
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(wrkFolder & wrkFile)
    
    If Not SheetExists("VPSchaal", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "VPSchaal"
    Else
        MsgBox ("Reeds aanwezig")
        Exit Sub
    End If
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("select * from Vps Order by Imps")
    rs.MoveFirst
    i = 1
    Do While Not rs.EOF
        MySheet.Cells(i, 1).Value = rs!imps
        MySheet.Cells(i, 2).Value = rs!vps_12
        MySheet.Cells(i, 3).Value = rs!vps_6
        MySheet.Cells(i, 4).Value = rs!vps_7
        MySheet.Cells(i, 5).Value = rs!vps_8
        MySheet.Cells(i, 6).Value = rs!vps_9
        MySheet.Cells(i, 7).Value = rs!vps_10
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    db.Close
    
    'creeer rangename
    Set rng = MySheet.Range(MySheet.Cells(1, 1), MySheet.Cells(i - 1, 7))
    'specify defined name
    MyRangeName = "VPSchaal"
    
    StartBook.Names.Add name:=MyRangeName, RefersTo:=rng
    
    StartBook.Save
    Set rng = Nothing
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub CreateWebInfoSheet(wrkFolder As Variant, wrkFile As Variant)
    Dim xlApp       As Object
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile, MyTableName As String
    
    Dim i, j        As Integer
    
    Dim rng         As Range
    Dim strRange    As String
    
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        'Creer een nieuwe
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(wrkFolder & wrkFile)
    
    If Not SheetExists("WebInfo", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "WebInfo"
    Else
        MsgBox ("Reeds aanwezig")
        Exit Sub
    End If
    
    MySheet.Cells(1, 1).Value = "Sessie"
    MySheet.Cells(1, 2).Value = "ActivityID"
    MySheet.Cells(2, 1).Value = 1
    MySheet.Cells(3, 1).Value = 2
    'creeer rangename
    Set rng = MySheet.Range(MySheet.Cells(1, 1), MySheet.Cells(3, 2))
    'specify defined name
    MyTableName = "WebInfo"
    
    MySheet.ListObjects.Add(xlSrcRange, rng, , xlYes).name = MyTableName
    
    StartBook.Save
    Set rng = Nothing
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub CreateTeamsSheet(varTeams As Variant, wrkFolder As Variant, wrkFile As Variant)
    
    Dim xlApp       As Object
    Dim TestExcel   As Integer
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile As String
    Dim question    As Integer
    Dim fcount, i, j As Integer
    Dim strWs       As String
    Dim rng         As Range
    TestExcel = False
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(strWorkfile)
    
    If Not SheetExists("Teams", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "Teams"
        'Teamnr  Teamnaam    Speler1 Speler2 Speler3 Speler4 Speler5 Speler6
        MySheet.Cells(1, 1).Value = "Teamnr"
        MySheet.Cells(1, 2).Value = "Teamnaam"
        MySheet.Cells(1, 3).Value = "Speler1"
        MySheet.Cells(1, 4).Value = "Speler2"
        MySheet.Cells(1, 5).Value = "Speler3"
        MySheet.Cells(1, 6).Value = "Speler4"
        MySheet.Cells(1, 7).Value = "Speler5"
        MySheet.Cells(1, 8).Value = "Speler6"
        MySheet.Cells(1, 9).Value = "Speler7"
        MySheet.Cells(1, 10).Value = "Speler8"
    Else
        Set MySheet = StartBook.Sheets("Teams")
    End If
    
    For i = 1 To varTeams
        If MySheet.Cells(i + 1, 1).Value = "" Then
            MySheet.Cells(i + 1, 1).Value = i
            If TestExcel = False Then TestExcel = True
        End If
        If MySheet.Cells(i + 1, 2).Value = "" Then
            MySheet.Cells(i + 1, 2).Value = "Team" & i
            If TestExcel = False Then TestExcel = True
        End If
    Next
    
    If Not TestExcel Then
        
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("Er zijn geen teams aangemaakt, reeds aanwezig")
        Exit Sub
    Else
        
        'creeer rangename
        Set rng = MySheet.Range(MySheet.Cells(1, 1), MySheet.Cells(varTeams + 1, 10))
        'specify defined name
        MyTableName = "Teams_Leden"
        MySheet.ListObjects.Add(xlSrcRange, rng, , xlYes).name = MyTableName
        StartBook.Save
    End If
    
    Set rng = Nothing
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub CreateKruisTabelSheet(varTeams, wrkFolder, wrkFile, Bye As Variant)
    Dim xlApp       As Object
    Dim TestExcel   As Integer
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile As String
    Dim question    As Integer
    Dim fcount, i, j As Integer
    Dim strWs       As String
    Dim strGemKolom As String
    Dim strTotaalRange As String
    Dim intTeams    As Integer
    Dim ByeTeam     As Integer
    Dim rng         As Range
    Dim strRange    As String
    
    'eerst kijk of er een kruistabel aanwezig
    
    'indien  aanwezig
    '   aanwezig
    '   test de matrix
    '   indien niet correct met het aantal teams
    '       opnieuw inrichten
    '       indien correct Leegmaken en met de uitslagentabel opnieuw inrichten
    '
    
    If Bye = 0 Then
        ByeTeam = False
    Else
        ByeTeam = True
    End If
    
    'Indien niet aanwezig.
    
    'creeer tabblad
    
    'test of er een werkbestand is
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(wrkFolder & wrkFile)
    
    'test eerst of er een tabblad teams is
    If Not SheetExists("Teams", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "Teams"
        'Teamnr  Teamnaam    Speler1 Speler2 Speler3 Speler4 Speler5 Speler6
        MySheet.Cells(1, 1).Value = "Teamnr"
        MySheet.Cells(1, 2).Value = "Teamnaam"
        MySheet.Cells(1, 3).Value = "Speler1"
        MySheet.Cells(1, 4).Value = "Speler2"
        MySheet.Cells(1, 5).Value = "Speler3"
        MySheet.Cells(1, 6).Value = "Speler4"
        MySheet.Cells(1, 7).Value = "Speler5"
        MySheet.Cells(1, 8).Value = "Speler6"
        MySheet.Cells(1, 9).Value = "Speler7"
        MySheet.Cells(1, 10).Value = "Speler8"
        For i = 1 To varTeams
            MySheet.Cells(i + 1, 1).Value = i
            MySheet.Cells(i + 1, 2).Value = "Team" & i
        Next
    End If
    
    If Not SheetExists("Kruistabel", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "Kruistabel"
    Else
        Set MySheet = StartBook.Sheets("Kruistabel")
    End If
    
    MySheet.Cells(1, 1).Value = "Uitslagen"
    MySheet.Cells(1, 1).Interior.color = RGB(155, 194, 230)
    MySheet.Cells(1, varTeams + 2).Value = "Totaal"
    MySheet.Cells(1, varTeams + 2).Interior.color = RGB(155, 194, 230)
    MySheet.Cells(1, varTeams + 3).Value = "Gem"
    MySheet.Cells(1, varTeams + 3).Interior.color = RGB(155, 194, 230)
    MySheet.Cells(1, varTeams + 4).Value = "Rang"
    MySheet.Cells(1, varTeams + 4).Interior.color = RGB(155, 194, 230)
    MySheet.Cells(1, varTeams + 5).Value = "Uitslagen"
    MySheet.Cells(1, varTeams + 5).Interior.color = RGB(155, 194, 230)
    
    For i = 1 To varTeams
        MySheet.Cells(1, i + 1).Formula = "=Teams!$B$" & i + 1
        MySheet.Cells(i + 1, 1).Formula = "=Teams!$B$" & i + 1
        MySheet.Cells(i + 1, varTeams + 5).Formula = "=Teams!$B$" & i + 1
        MySheet.Cells(1, i + 1).Interior.color = RGB(155, 194, 230)
        MySheet.Cells(i + 1, 1).Interior.color = RGB(155, 194, 230)
        MySheet.Cells(i + 1, varTeams + 5).Interior.color = RGB(155, 194, 230)
    Next
    
    'rij N+2
    'Totaal  Gem Rang    uitslagen
    
    If Not ByeTeam Then
        intTeams = varTeams
    Else
        intTeams = varTeams - 1
    End If
    
    strTotaalRange = BerekenKolom(intTeams + 1)
    
    strGemKolom = BerekenKolom(varTeams + 3)
    
    For i = 1 To varTeams
        'Totaal
        
        MySheet.Cells(1 + i, varTeams + 2).Formula = "=Sum($B$" & i + 1 & ":$" & strTotaalRange & "$" & i + 1 & ")"
        MySheet.Cells(1 + i, varTeams + 3).Formula = "=IFERROR(AVERAGEIF($B$" & i + 1 & ":$" & strTotaalRange & "$" & i + 1 & "," & Chr(34) & "<>""" & Chr(34) & Chr(34) & ")" & "," & Chr(34) & "" & Chr(34) & ")"
        MySheet.Cells(1 + i, varTeams + 4).Formula = "=IFERROR(RANK(" & strGemKolom & i + 1 & ",$" & strGemKolom & "$2:$" & strGemKolom & "$" & intTeams + 1 & ",0)," & Chr(34) & "" & Chr(34) & ")"
        MySheet.Cells(i + 1, i + 1).Value = "xxx"
        'MySheet.Cells(i + 1, i + 1).Interior.color = RGB(255, 255, 0)
        MySheet.Cells(i + 1, i + 1).HorizontalAlignment = xlCenter
        
    Next
    
    Set rng = MySheet.Range(MySheet.Cells(1, 1), MySheet.Cells(varTeams + 1, varTeams + 5))
    
    With rng.Borders
        .LineStyle = xlContinuous
        .color = vbBlack
        .Weight = xlThin
    End With
    
    With rng.BorderAround(xlContinuous, xlThick)
    End With
    Set rng = Nothing
    
    'conditional formatting
    Set rng = MySheet.Range(MySheet.Cells(2, 2), MySheet.Cells(varTeams + 1, varTeams + 1))
    With rng.FormatConditions
        .Delete
        With .Add(xlTextString, TextOperator:=xlContains, String:="xxx")
            .Interior.color = RGB(211, 216, 127)
        End With
        
        With .Add(xlCellValue, xlNotEqual, "=" & Chr(34) & Chr(34))
            .Interior.color = RGB(100, 206, 112)
        End With
        
    End With
    
    StartBook.Save
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub CreateOpstellingSheet(varAantalTeams As Variant, varWedstrijdenPerSessie As Variant, wrkFolder As Variant, wrkFile As Variant)
    Dim xlApp       As Object
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile, MyTableName As String
    Dim rng         As Range
    Dim i, j        As Integer
    
    Dim strRange    As String
    
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        'Creer een nieuwe
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(wrkFolder & wrkFile)
    
    If Not SheetExists("Import_Opstelling", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "Import_Opstelling"
    Else
        MsgBox ("Reeds aanwezig")
        Exit Sub
    End If
    
    'sessie Teamnr Speler1 Speler2 Speler3 Speler4  Wedstrijd1  Wedstrijd2
    
    MySheet.Cells(1, 1).Value = "Sessie"
    MySheet.Cells(1, 2).Value = "Teamnr"
    For i = 1 To 4
        MySheet.Cells(1, 2 + i).Value = "Speler" & i
    Next
    
    For i = 1 To varWedstrijdenPerSessie
        MySheet.Cells(1, 6 + i).Value = "Wedstrijd" & i
    Next
    For i = 1 To varAantalTeams
        MySheet.Cells(1 + i, 1).Value = 1
        MySheet.Cells(1 + i, 2).Value = i
    Next
    
    'creeer rangename
    Set rng = MySheet.Range(MySheet.Cells(1, 1), MySheet.Cells(varAantalTeams + 1, 6 + varWedstrijdenPerSessie))
    'specify defined name
    MyTableName = "Indeling"
    
    MySheet.ListObjects.Add(xlSrcRange, rng, , xlYes).name = MyTableName
    
    StartBook.Save
    Set rng = Nothing
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub CreateTeamUitslagenSheet(varAantalTeams As Variant, wrkFolder As Variant, wrkFile As Variant)
    Dim xlApp       As Object
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile, MyTableName As String
    
    Dim i, j        As Integer
    
    Dim rng         As Range
    Dim strRange    As String
    
    strWorkfile = wrkFolder & wrkFile
    
    If Not fnExists(strWorkfile) Then
        'Creer een nieuwe
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(wrkFolder & wrkFile)
    
    If Not SheetExists("TeamUitslagen", StartBook) Then
        Set MySheet = StartBook.Sheets.Add
        MySheet.name = "TeamUitslagen"
    Else
        Set rng = Nothing
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        
        MsgBox ("Reeds aanwezig")
        Exit Sub
    End If
    
    'Avond/Sessie   Wedstrijd   Teamnr_thuis    Teamnr_Uit  TeamThuis   TeamUit ImpsThuis   ImpsUit VPThuis Vpuit
    
    MySheet.Cells(1, 1).Value = "Sessie/Avond"
    MySheet.Cells(1, 2).Value = "Wedstrijd"
    MySheet.Cells(1, 3).Value = "Teamnr_thuis"
    MySheet.Cells(1, 4).Value = "Teamnr_Uit"
    MySheet.Cells(1, 5).Value = "TeamThuis"
    MySheet.Cells(1, 6).Value = "TeamUit"
    MySheet.Cells(1, 7).Value = "ImpsThuis"
    MySheet.Cells(1, 8).Value = "ImpsUit"
    MySheet.Cells(1, 9).Value = "VPThuis"
    MySheet.Cells(1, 10).Value = "VPUit"
    
    'ingevuld wordt de eerste sessie de eerste wedstrijd
    
    For i = 1 To varAantalTeams \ 2
        MySheet.Cells(1 + i, 1).Value = 1
        MySheet.Cells(1 + i, 2).Value = 1
        MySheet.Cells(1 + i, 3).Value = (i - 1) * 2 + 1
        MySheet.Cells(1 + i, 4).Value = i * 2
    Next
    
    '=ALS.FOUT(VERT.ZOEKEN(C2;Teams_Leden;2);"")
    
    'creeer rangename
    Set rng = MySheet.Range(MySheet.Cells(1, 1), MySheet.Cells(1 + varAantalTeams \ 2, 10))
    'specify defined name
    MyTableName = "TeamUitslagen"
    
    MySheet.ListObjects.Add(xlSrcRange, rng, , xlYes).name = MyTableName
    
    MySheet.Cells(2, 5).Formula = "=IFERROR(VLOOKUP(C2,Teams_Leden,2)," & Chr(34) & Chr(34) & ")"
    MySheet.Cells(2, 6).Formula = "=IFERROR(VLOOKUP(D2,Teams_Leden,2)," & Chr(34) & Chr(34) & ")"
    
    StartBook.Save
    Set rng = Nothing
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub

Public Sub ImportUitslagen(varToernooi As Variant, varSessieID As Variant)
    Dim db          As Database
    Dim rs          As Recordset
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile As String
    Dim question    As Integer
    Dim TeamsID()   As Long
    
    Call InitAll(varToernooi, varSessieID)
    ReDim TeamsID(AANTALTEAMS)
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("select * from tblTeams where [ToernooiID] = " & lngToernooi)
    If rs.BOF And rs.EOF Then
        MsgBox ("er zijn nog geen teams geimporteerd van dit toernooi")
        rs.Close
        db.Close
        Exit Sub
    End If
    
    rs.MoveFirst
    Do While Not rs.EOF
        TeamsID(rs!Teamnr) = rs!Id
        rs.MoveNext
    Loop
    rs.Close
    
    Set rs = db.OpenRecordset("select * from tblUitslagen where [ToernooiID] = " & lngToernooi & " And SessieID = " & lngSessie)
    
    If Not (rs.BOF And rs.EOF) Then
        MsgBox ("Indeling Is al geimporteerd")
        rs.Close
        db.Close
        Exit Sub
    End If
    
    'test of er een werkbestand is
    strWorkfile = WORKFOLDER & WORKFILE
    
    If Not fnExists(strWorkfile) Then
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
    Set MySheet = StartBook.Worksheets("Teamuitslagen")
    
    If MySheet.Cells(2, 1) = "" Then
        question = MsgBox("Er zijn geen teamuitslagen aanwezig")
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        Exit Sub
    End If
    
    'methode met transferspreadsheet
    
    If TableExists("TeamUitslagen", MySheet) Or TableExists("TeamUitslagen", StartBook) Then
        
        '
        'DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12, "tbl_" & lngToernooi & "_Teamuitslagen", strWorkFile, True, "TeamUitslagen!"
        'DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12, "tbl_" & lngToernooi & "_Schema", strWorkFile, True, "Schema!"
        
    End If
    
    rijteller = 2
    
    Do While MySheet.Cells(rijteller, 1) <> "" And MySheet.Cells(rijteller, 1) <> Sessienr
        rijteller = rijteller + 1
    Loop
    If MySheet.Cells(rijteller, 1) = "" Then
        MsgBox ("De indeling Is nog niet ingevoerd")
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("select * from tblUitslagen where [ToernooiID] = " & varToernooi & " And SessieID = " & varSessieID)
    If Not (rs.BOF And rs.EOF) Then
        MsgBox ("De indeling Is reeds geimporteerd")
        'Controleren of er verschillen zijn
        'Verwijderen en opnieuw laden
        
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("tblUitslagen")
    
    Do While MySheet.Cells(rijteller, 1) = Sessienr
        rs.AddNew
        rs!ToernooiID = lngToernooi
        rs!SessieID = lngSessie
        rs!Wedstrijdnr = MySheet.Cells(rijteller, 2).Value
        rs!TeamIDThuis = TeamsID(MySheet.Cells(rijteller, 3).Value)
        rs!TeamIDUit = TeamsID(MySheet.Cells(rijteller, 4).Value)
        If MySheet.Cells(rijteller, 7).Value <> "" Then
            rs!ImpsThuis = MySheet.Cells(rijteller, 7).Value
        End If
        If MySheet.Cells(rijteller, 8).Value <> "" Then
            rs!ImpsUit = MySheet.Cells(rijteller, 8).Value
        End If
        If MySheet.Cells(rijteller, 9).Value <> "" Then
            rs!VpsThuis = MySheet.Cells(rijteller, 9).Value
        End If
        If MySheet.Cells(rijteller, 10).Value <> "" Then
            rs!VpsUit = MySheet.Cells(rijteller, 10).Value
        End If
        rs.Update
        rijteller = rijteller + 1
    Loop
    
    rs.Close
    db.Close
    
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub
'Import functies

Public Sub ImportTeams(varToernooi As Variant)
    'kijk eerst op er al teams opgenomen zijn
    Dim xlApp       As Object
    Dim TestExcel, question As Integer
    Dim db          As Database
    Dim rs          As Recordset
    Dim MySheet     As Worksheet
    Dim StartBook   As Workbook
    Dim strWorkfile As String
    
    Dim tbl         As ListObject
    Dim rw          As ListRow
    Dim cl          As Range
    Dim intTeamsGeladen As Integer
    Dim IntTeamNummer As Integer
    
    intTeamsGeladen = False
    Set db = CurrentDb
    Set rs = db.OpenRecordset("select * from tblTeams where [ToernooiID] = " & varToernooi)
    
    If Not (rs.BOF And rs.EOF) Then
        question = MsgBox("Er zijn reeds teams geladen of aanwezig, doorgaan J/N ", vbYesNo)
        If question = vbNo Then
            rs.Close
            db.Close
            Exit Sub
        End If
        intTeamsGeladen = True
    End If
    
    'test of er een werkbestand is
    strWorkfile = WORKFOLDER & WORKFILE
    
    If Not fnExists(strWorkfile) Then
        MsgBox ("Er Is nog geen excel bestand aangemaakt")
        Exit Sub
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
    Set MySheet = StartBook.Worksheets("Teams")
    
    'test of er een tabel is
    
    If TableExists("Teams", MySheet) Or TableExists("Teams", StartBook) Then
        
        Set tbl = MySheet.ListObjects("Teams")
        IntTeamNummer = 1
        For Each rw In tbl.ListRows
            i = 1
            For Each cl In rw.Range
                'eerste element is de TEAMNR
                
                Select Case i
                    Case 1
                        IntTeamNummer = cl.Value
                        If intTeamsGeladen = True Then
                            rs.MoveFirst
                            gevonden = False
                            Do While Not rs.EOF
                                If rs!TeamNummer = IntTeamNummer Then
                                    gevonden = True
                                    Exit Do
                                End If
                                rs.MoveNext
                            Loop
                            If gevonden Then
                                rs.Edit
                            Else
                                rs.AddNew
                            End If
                            
                        Else
                            rs.AddNew
                        End If
                        rs!ToernooiID = varToernooi
                        rs!Teamnr = IntTeamNummer
                    Case 2
                        rs!TeamNaam = cl.Value
                    Case Else
                        If i < 9 Then
                            rs.Fields("Speler" & i - 2) = cl.Value
                        End If
                End Select
                i = i + 1
            Next cl
            rs.Update
        Next rw
        rs.Close
        db.Close
        
        Set rw = Nothing
        Set cl = Nothing
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        
        Exit Sub
        
    Else
        
        teller = 1
        gevonden = False
        Do While MySheet.Cells(teller, 1) <> "" And teller < 9
            teller = teller + 1
        Loop
        
        If teller > 9 Then
            question = MsgBox("Er zijn geen teams aanwezig")
            Set MySheet = Nothing
            Set StartBook = Nothing
            xlApp.Application.DisplayAlerts = True
            xlApp.Application.Quit
            Set xlApp = Nothing
            Exit Sub
        End If
        
        rijteller = rijteller + 1
        TestExcel = False
        Do While MySheet.Cells(rijteller, 1) <> ""
            
            If Not intTeamsGeladen Then
                rs.AddNew
            Else
                IntTeamNummer = MySheet.Cells(rijteller, 1).Value
                rs.MoveFirst
                gevonden = False
                Do While Not rs.EOF
                    If rs!TeamNummer = IntTeamNummer Then
                        gevonden = True
                        Exit Do
                    End If
                    rs.MoveNext
                Loop
                If gevonden Then
                    rs.Edit
                Else
                    rs.AddNew
                End If
            End If
            rs!ToernooiID = varToernooi
            rs!Teamnr = IntTeamNummer
            rs!TeamNaam = MySheet.Cells(rijteller, 2).Value
            'niet meer dan 8 spelers
            Kolom = 3
            speler = 1
            Do While MySheet.Cells(rijteller, Kolom).Value <> "" And speler < 9
                rs.Fields("speler" & speler) = MySheet.Cells(rijteller, Kolom).Value
                Kolom = Kolom + 1
                speler = speler + 1
            Loop
            rs.Update
            rijteller = rijteller + 1
        Loop
        rs.Close
        db.Close
    End If
    
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
End Sub