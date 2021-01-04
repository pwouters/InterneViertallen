Option Compare Database
Public MyBookmark   As Variant
Public CtrlVisible  As Integer
Public recordteller As Long
Public lngPK        As Long
Public WedstrijdenVanaf, WedstrijdenTot, GespeeldeRondes As Integer


Global Const STEP_RESULTS = "https://results.stepbridge.nl/tournament/events/show/"
Global Const STEP_DATA = "http://admin.stepbridge.nl/show.php?page=tournamentinfo&activityid="
Global Const LOCAL_SITE = "https://www.pwobridge.nl/Step/"

Global Const GEKOPPELDINACCESS = "Teams;Schema;Uitslagen;Opstelling;WebInfo"
Global Const EXCELTABNAMEN = "Teams;Schema;TeamUitslagen;Import_Opstelling;WebInfo;Team_template;Kruistabel;VPSchaal;Imptabel"


Global WORKFOLDER   As String
Global WORKID       As Long
Global WORKFILE     As String
Global STEPRESULTS  As String
Global STEPDATA     As String
Global LOCALSITE    As String
Global LOCALHTML    As String
Global ToernooiNaam   As String
' weergave in de tabbladen, Kan avond zijn meerdere wedstrijden etc
Global PREFIX       As String
'Aantal wedstrijden per avond, per dag van hoeveel krijg ik de uitslagen tegelijk binnen.
Global BEKERWEDSTRIJD As Integer
Global UITREKENVORM As Integer

Global AANTALTEAMS   As Integer
Global AANTALSPELLENPERWEDSTRIJD As Integer
Global TEAMBYE      As Integer
Global WEDSTRIJD    As Integer
Global WEDSTRIJDENPERSESSIE  As Integer
Global COMPETITIEVORM  As Integer
Global AANTALSESSIES  As Integer
Global Prefixkopjesscorestaat As String
Global PrefixKopjeuitslagen As String
Global Suffixkopjesscorestaat As String
Global SuffixKopjeuitslagen As String
Global Voettekst    As String
Global Voetlink     As String
Global lngToernooi  As Long
Global lngSessie    As Long
Global lngTeam      As Long
Global lngToernooiOld As Long
Global lngSessieOld As Long
Global intWedstrijdronde As Integer
Global intUitvoerNaarHTML As Integer
Global intExcelZichtbaar As Integer
Global ScorestaatIntern As Integer
Global ScorestaatExcel As Integer
Global BerekenAlleStaten As Integer
Global strSheetName As String
Global ActivityID   As Integer
Global Sessienaam   As String
Global Sessienr     As Integer
Global strExcel_Folder As String
Global strHTML_Folder As String
Global strTemplate_Folder As String
Global strTemplate_File As String

'Leesbaar maken van de code
'

Public Enum Opstelling
    'sessie Teamnr Speler1 Speler2 Speler3 Speler4  Wedstrijd1  Wedstrijd2 etc
    Sessie__nr = 1
    Team__nr = 2
    Speler__1 = 3
    Speler__2 = 4
    Speler__3 = 5
    Speler__4 = 6
    Wedstrijd__1 = 7
    Wedstrijd__2 = 8
    Wedstrijd__3 = 9
    Wedstrijd__4 = 10
    wedstrijd__5 = 11
    Wedstrijd__6 = 12
    Wedstrijd__7 = 13
    Wedstrijd__8 = 14
    Wedstrijd__9 = 15
    Wedstrijd__10 = 16
End Enum

Public Enum Team
    Team_nr = 1
    Team_naam = 2
    Speler_1 = 3
    Speler_2 = 4
    Speler_3 = 5
    Speler_4 = 6
    Speler_5 = 7
    Speler_6 = 8
    Speler_7 = 9
    Speler_8 = 10
    Club_Naam = 11
End Enum

Public Enum Uitslag
    Sessie_nr = 1
    Wedstrijd_nr = 2
    Thuis_nr = 3
    Uit_nr = 4
    Thuis_Naam = 5
    Uit_Naam = 6
    Thuis_Imps = 7
    Uit_Imps = 8
    Thuis_VPs = 9
    Uit_VPs = 10
End Enum

Public Enum Uitrekenmethode
    VPs_u = 0
    Imps_u = 1
    Patton_u = 2
End Enum


Public Const f_Sessie_nr = "Avond"
Public Const f_Wedstrijd_nr = "Wedstrijd"
Public Const f_Thuis_nr = "Teamnr_thuis"
Public Const f_Uit_nr = "Teamnr_Uit"
Public Const f_Thuis_Naam = "TeamThuis"
Public Const f_Uit_Naam = "TeamUit"
Public Const f_Thuis_Imps = "ImpsThuis"
Public Const f_Uit_Imps = "ImpsUit"
Public Const f_Thuis_VPs = "VPThuis"
Public Const f_Uit_VPs = "Vpuit"

Public Enum Scorestaat
    Spel_1 = 1
    Contract_1 = 2
    Resultaat_1 = 3
    Door_1 = 4
    Score_1 = 5
    Imps_butler_1 = 6
    Spel_2 = 8
    Contract_2 = 9
    Resultaat_2 = 10
    Door_2 = 11
    Score_2 = 12
    Imps_butler_2 = 13
    Saldo_staat = 15
    Imps_staat = 16
    Imps_Wij = 17
    Imps_Zij = 18
    Wij_Zij = 19
    Uitslag_Team = 20
    Uitslag_Imps = 21
    Uitslag_Verschil = 22
    Uitslag_VPs = 23
End Enum

Public Enum VPsTabel
    VP_Imps = 1
    VP_12 = 2
    VP_6 = 3
    VP_7 = 4
    VP_8 = 5
    VP_9 = 6
    VP_10 = 7
End Enum


Public Sub InitToernooi(id As Variant)
    Dim db          As Database
    Dim rs          As Recordset
    
    If id <> lngToernooiOld Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset("select * from tblToernooi where id =" & id)
        rs.MoveFirst
        ToernooiNaam = rs.Fields("ToernooiNaam")
        WORKID = rs.Fields("ID")
        WORKFOLDER = rs.Fields("WORKFOLDER")
        WORKFILE = rs.Fields("WORKFILE")
        STEPRESULTS = rs.Fields("STEPRESULTS")
        STEPDATA = rs.Fields("STEPDATA")
        LOCALSITE = rs.Fields("LOCALSITE")
        LOCALHTML = rs.Fields("LOCALHTML")
        AANTALSESSIES = rs.Fields("AANTALSESSIES")
        PREFIX = rs.Fields("PREFIX")
        BEKERWEDSTRIJD = rs.Fields("BEKERWEDSTRIJD")
        UITREKENVORM = rs.Fields("UITREKENVORM")
        
        
        lngToernooi = id
        lngToernooiOld = lngToernooi
        rs.Close
        db.Close
    End If
    
End Sub

Public Sub InitSessie(id As Variant)
    Dim db          As Database
    Dim rs          As Recordset
    Set db = CurrentDb
    If id <> lngSessieOld Then
        Set rs = db.OpenRecordset("select * from tblSessie where id =" & id)
        rs.MoveFirst
        AANTALTEAMS = rs.Fields("AantalTeams")
        AANTALSPELLENPERWEDSTRIJD = rs.Fields("Aantalspellen")
        If rs.Fields("ByeTeam") Then
            TEAMBYE = AANTALTEAMS
        Else
            TEAMBYE = 0
        End If
        WEDSTRIJD = rs.Fields("Competitie")
        WEDSTRIJDENPERSESSIE = rs.Fields("AantalWedstrijdenPerSessie")
        COMPETITIEVORM = rs.Fields("wedstrijdvormID")
        If Not IsNull(rs.Fields("Prefixkopjesscorestaat")) Then
            Prefixkopjesscorestaat = rs.Fields("Prefixkopjesscorestaat")
        Else
            Prefixkopjesscorestaat = ""
        End If
        If Not IsNull(rs.Fields("PrefixKopjeuitslagen")) Then
            PrefixKopjeuitslagen = rs.Fields("PrefixKopjeuitslagen")
        Else
            PrefixKopjeuitslagen = ""
        End If
        If Not IsNull(rs.Fields("Suffixkopjesscorestaat")) Then
            Suffixkopjesscorestaat = rs.Fields("Suffixkopjesscorestaat")
        Else
            Suffixkopjesscorestaat = ""
        End If
        If Not IsNull(rs.Fields("SuffixKopjeuitslagen")) Then
            SuffixKopjeuitslagen = rs.Fields("SuffixKopjeuitslagen")
        Else
            SuffixKopjeuitslagen = ""
        End If
        
        If Not IsNull(rs.Fields("Voettekst")) Then
            Voettekst = rs.Fields("Voettekst")
        Else
            Voettekst = ""
        End If
        If Not IsNull(rs.Fields("Voetlink")) Then
            Voetlink = rs.Fields("Voetlink")
        Else
            Voetlink = ""
        End If
        
        ActivityID = rs.Fields("ActivityID")
        Sessienaam = rs.Fields("Sessienaam")
        Sessienr = rs.Fields("Sessienr")
        lngSessie = id
        lngSessieOld = lngSessie
        rs.Close
        db.Close
    End If
    
End Sub

Public Sub InitAll(ToernooiID As Variant, SessieID As Variant)
    
    Dim db          As Database
    Dim rs          As Recordset
    Set db = CurrentDb
    
    If ToernooiID <> lngToernooiOld Then
        Set rs = db.OpenRecordset("select * from tblToernooi where id =" & ToernooiID)
        rs.MoveFirst
        WORKID = rs.Fields("ID")
        ToernooiNaam = rs.Fields("ToernooiNaam")
        WORKFOLDER = rs.Fields("WORKFOLDER")
        WORKFILE = rs.Fields("WORKFILE")
        STEPRESULTS = rs.Fields("STEPRESULTS")
        STEPDATA = rs.Fields("STEPDATA")
        LOCALSITE = rs.Fields("LOCALSITE")
        LOCALHTML = rs.Fields("LOCALHTML")
        AANTALSESSIES = rs.Fields("AANTALSESSIES")
        PREFIX = rs.Fields("PREFIX")
        BEKERWEDSTRIJD = rs.Fields("BEKERWEDSTRIJD")
        UITREKENVORM = rs.Fields("UITREKENVORM")

        lngToernooi = rs!id
        lngToernooiOld = lngToernooi
        rs.Close
    End If
    If SessieID <> lngSessieOld Then
        Set rs = db.OpenRecordset("select * from tblSessie where id =" & SessieID)
        
        rs.MoveFirst
        AANTALTEAMS = rs.Fields("AantalTeams")
        AANTALSPELLENPERWEDSTRIJD = rs.Fields("Aantalspellen")
        If rs.Fields("ByeTeam") Then
            TEAMBYE = AANTALTEAMS
        Else
            TEAMBYE = 0
        End If
        WEDSTRIJD = rs.Fields("Competitie")
        WEDSTRIJDENPERSESSIE = rs.Fields("AantalWedstrijdenPerSessie")
        COMPETITIEVORM = rs.Fields("wedstrijdvormID")
        If Not IsNull(rs.Fields("Prefixkopjesscorestaat")) Then
            Prefixkopjesscorestaat = rs.Fields("Prefixkopjesscorestaat")
        Else
            Prefixkopjesscorestaat = ""
        End If
        If Not IsNull(rs.Fields("PrefixKopjeuitslagen")) Then
            PrefixKopjeuitslagen = rs.Fields("PrefixKopjeuitslagen")
        Else
            PrefixKopjeuitslagen = ""
        End If
        If Not IsNull(rs.Fields("Suffixkopjesscorestaat")) Then
            Suffixkopjesscorestaat = rs.Fields("Suffixkopjesscorestaat")
        Else
            Suffixkopjesscorestaat = ""
        End If
        If Not IsNull(rs.Fields("SuffixKopjeuitslagen")) Then
            SuffixKopjeuitslagen = rs.Fields("SuffixKopjeuitslagen")
        Else
            SuffixKopjeuitslagen = ""
        End If
        
        If Not IsNull(rs.Fields("Voettekst")) Then
            Voettekst = rs.Fields("Voettekst")
        Else
            Voettekst = ""
        End If
        If Not IsNull(rs.Fields("Voetlink")) Then
            Voetlink = rs.Fields("Voetlink")
        Else
            Voetlink = ""
        End If
        
        ActivityID = rs.Fields("ActivityID")
        Sessienaam = rs.Fields("Sessienaam")
        Sessienr = rs.Fields("Sessienr")
        lngSessie = SessieID
        lngSessieOld = lngSessie
        rs.Close
    End If
    db.Close
    If CurrentProject.AllForms("Start_VT").IsLoaded = True Then
        
        Forms("Start_VT").lblHuidigToernooi.Caption = ToernooiNaam
        Forms("Start_VT").lblHuidigeSessie.Caption = "Sessie " & Sessienr
    End If
    
End Sub

Public Function lngToernooiID() As Long
    lngToernooiID = lngToernooi
End Function
Public Function lngSessieID() As Long
    lngSessieID = lngSessie
End Function
Public Function fnSaveRecords()
    Dim DgDef, response As Integer
    Dim msg, Title  As String
    DgDef = vbYesNo + vbCritical + vbDefaultButton2
    msg = "Het record Is gewijzigd, wilt u deze wijzigingen opslaan?"
    Title = "Attentie"
    response = MsgBox(msg, DgDef, Title)
    
    If response = vbYes Then
        DoCmd.RunCommand acCmdSaveRecord
    Else
        DoCmd.RunCommand acCmdUndo
    End If
    
End Function
Public Sub GotoLastCurrentRecord(MyFormName As Variant, PK As Variant, PKpointer As Variant)
    With Forms(MyFormName).RecordsetClone
        .FindFirst PK & "=" & PKpointer
        If Not .NoMatch Then
            If Forms(MyFormName).Dirty Then
                Forms(MyFormName).Dirty = False
            End If
            Forms(MyFormName).Bookmark = .Bookmark
        End If
    End With
End Sub

Public Function GetFolderName(Optional OpenAt As String, Optional Soortbestanden As String) As String
    Dim lCount      As Long
    Dim Soort       As String
    Soort = IIf(Not (Soortbestanden = ""), Soortbestanden, "excel")
    Dim Folder      As String
    
    GetFolderName = vbNullString
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = OpenAt
        .Title = "Kies folder voor " & Soort & " bestanden"
        
        Dim FileChosen As Integer
        FileChosen = .Show
        If FileChosen Then
            For lCount = 1 To .SelectedItems.Count
                Folder = .SelectedItems(lCount)
            Next lCount
            If Right(Folder, 1) <> "\" Then
                Folder = Folder & "\"
            End If
            GetFolderName = Folder
        End If
    End With
    
End Function

Public Function GetFileName(strExcel_Folder) As String
    Dim lCount      As Long
    Dim FileChosen  As Integer
    GetFileName = vbNullString
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Kies Excel bestand"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx"
        .InitialFileName = strExcel_Folder
        '.Show
        FileChosen = .Show
        If FileChosen Then
            GetFileName = .SelectedItems(1)
        End If
    End With
End Function

Function KiesModelExcelBestandEnCopieer() As Integer
    Dim source, destination, src, dest As String
    
    Dim fd          As Office.FileDialog
    KiesModelExcelBestandEnCopieer = True
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'get the number of the button chosen
    fd.InitialFileName = LaatsteLocatieModelfile
    fd.AllowMultiSelect = False
    fd.Title = "Kies Excel model bestand of maak zelf een excelfile aan"
    
    Dim FileChosen  As Integer
    FileChosen = fd.Show
    If FileChosen = -1 Then
        source = fd.SelectedItems(1)
        
        destination = InputBox("Toernooinaam: (zal dan tevens de naam van het werkbestand worden) ", "Werkbestand", "Nieuw_Toernooi")
        
        'Moet de copie indezelfde folder komen
        
        'indien nee
        
        'folder dialoog
        
        'strFolder
        
        If destination = "" Then
            KiesModelExcelBestandEnCopieer = False
            Exit Function
        End If
        'test of bron en bestemming niet hetzelfde zijn
        
        ToernooiNaam = destination
        
        destination = destination & ".xlsx"
        
        WORKFILE = destination
        destination = FolderFromPath(source) & destination
        
        If source <> destination Then
            FileCopy source, destination
        End If
        
        WORKFOLDER = FolderFromPath(source)
        If Right(WORKFOLDER, 1) <> "\" Then WORKFOLDER = WORKFOLDER & "\"
        
    Else
        KiesModelExcelBestandEnCopieer = False
        Exit Function
    End If
    
End Function

Public Function LaatsteLocatieModelfile() As String
    Dim db          As Database
    Dim rs          As Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblToernooi")
    rs.MoveLast
    LaatsteLocatieModelfile = rs!WORKTEMPLATE
    rs.Close
    db.Close
    
End Function

Function fnFolderExists(fldr As Variant) As Integer
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FolderExists(fldr)) Then
        fnFolderExists = True
    Else
        fnFolderExists = False
    End If
    Set fso = Nothing
End Function

Public Function GetDesktopfolder() As String
    GetDesktopfolder = Environ("USERPROFILE") & "\Desktop"
End Function

Public Function GetDocumentsfolder() As String
    GetDocumentsfolder = Environ("USERPROFILE") & "\Documents"
End Function

Public Function FileNameFromPath(strFullPath As Variant) As String
    FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Public Function FolderFromPath(strFullPath As Variant) As String
    FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
End Function

Public Function BerekenKolom(varGetal As Variant) As String
    Dim Getal       As Integer
    Dim rest        As Integer
    Dim strKolom    As String
    
    Getal = varGetal
    Do While Getal > 26
        rest = Getal Mod 26
        If rest = 0 Then
            rest = 26
            Getal = Getal - 26
        End If
        strKolom = Chr(64 + rest) & strKolom
        Getal = Getal \ 26
    Loop
    rest = Getal Mod 26
    If rest = 0 Then rest = 26
    strKolom = Chr(64 + rest) & strKolom
    BerekenKolom = strKolom
    
End Function

Public Function AccessTableExists(TableName As String) As Boolean
    On Error Resume Next
    AccessTableExists = CurrentDb.TableDefs(TableName).name = TableName
End Function

Public Sub CreateAllExcelLinks(varToernooi As Variant, strWorkfile As Variant)
    Dim excelLinks() As String
    Dim accessLinks() As String
    Dim i           As Integer
    excelLinks = Split(EXCELTABNAMEN, ";")
    accessLinks = Split(GEKOPPELDINACCESS, ";")
    Call DeleteAllExcelLinks(varToernooi)
    For i = 0 To UBound(accessLinks)
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl_" & varToernooi & "_" & accessLinks(i), strWorkfile, True, excelLinks(i) & "!"
    Next
End Sub

Public Sub ImportTempTeamUitslagenTabel(strWorkfile As Variant)
    Call DeleteTable("tbl_temp_Uitslagen")
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl_temp_Uitslagen", strWorkfile, True, "TeamUitslagen!"
End Sub
Public Function DeleteAllExcelLinks(varToernooi As Variant)
    Dim strTable    As String
    Dim tbl         As TableDef
    strTable = "tbl_" & varToernooi & "_*"
    For Each tbl In CurrentDb.TableDefs
        If tbl.name Like strTable Then
            DoCmd.DeleteObject acTable, tbl.name
        End If
    Next
End Function
Public Function DeleteTable(varTabel As Variant)
    Dim strTable    As String
    Dim tbl         As TableDef
    For Each tbl In CurrentDb.TableDefs
        If tbl.name = varTabel Then
            DoCmd.DeleteObject acTable, tbl.name
            Exit For
        End If
    Next
End Function
Function ValidFileName(FileName As String) As Boolean
    Const sBadChar  As String = "\/:*?<>|[]"""
    Dim i           As Long
    
    ' assume valid unless it isn't
    ValidFileName = True
    
    For i = 1 To Len(sBadChar)
        If InStr(FileName, Mid$(sBadChar, i, 1)) > 0 Then
            ' invalid
            ValidFileName = False
            Exit For
        End If
    Next
End Function
Public Function SheetExists(strSheetName As String, wbWorkbook As Object) As Boolean
   Dim obj As Object
    On Error GoTo HandleError
    Set obj = wbWorkbook.Sheets(strSheetName)
    SheetExists = True
    Exit Function
HandleError:
    SheetExists = False
End Function
Public Function TableExists(TableName As String, wbWorkbook As Object) As Boolean
   Dim obj As Object
    On Error GoTo HandleError
    Set obj = wbWorkbook.ListObjects(TableName)
    TableExists = True
    Exit Function
HandleError:
    TableExists = False
End Function
Function DistinctRandomNumbers(NumCount As Long, LLimit As Long, ULimit As Long) As Variant
Dim RandColl As Collection, i As Long, varTemp() As Long
DistinctRandomNumbers = False
If NumCount < 1 Then Exit Function
If LLimit > ULimit Then Exit Function
If NumCount > (ULimit - LLimit + 1) Then Exit Function
Set RandColl = New Collection
Randomize
Do
    On Error Resume Next
    i = CLng(Rnd * (ULimit - LLimit) + LLimit)
    RandColl.Add i, CStr(i)
    On Error GoTo 0
Loop Until RandColl.Count = NumCount
ReDim varTemp(1 To NumCount)
For i = 1 To NumCount
    varTemp(i) = RandColl(i)
Next i
Set RandColl = Nothing
DistinctRandomNumbers = varTemp
Erase varTemp
End Function

Sub test()
Dim qArray() As Long
ReDim qArray(1 To 40)

qArray() = RandomQuestionArray(40)
'loop through your questions

End Sub

Function RandomQuestionArray(varTeams As Integer)
Dim i As Long, n As Long
Dim numArray() As Long
Dim numCollection As New Collection
ReDim numArray(varTeams)
With numCollection
    For i = 1 To varTeams
        .Add i
    Next
    For i = 1 To varTeams
        n = Rnd * (.Count - 1) + 1
        numArray(i) = numCollection(n)
        .Remove n
    Next
End With

RandomQuestionArray = numArray()

End Function
Public Function fnVan() As Integer
    fnVan = WedstrijdenVanaf
End Function
Public Function fnTot() As Integer
    fnTot = WedstrijdenTot
End Function
Public Function fn_Gespeeld() As Integer
fn_Gespeeld = GespeeldeRondes
End Function