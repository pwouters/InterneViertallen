Option Compare Database
Public MyBookmark As Variant
Public CtrlVisible As Integer
Public recordteller As Long
Public lngPK As Long

Global Const STEP_RESULTS = "https://results.stepbridge.nl/tournament/events/show/"
Global Const STEP_DATA = "http://admin.stepbridge.nl/show.php?page=tournamentinfo&activityid="
Global Const LOCAL_SITE = "https://www.pwobridge.nl/Step/"

Global Const GEKOPPELDINACCESS = "Teams;Schema;Uitslagen;Opstelling;WebInfo"
Global Const EXCELTABNAMEN = "Teams;Schema;TeamUitslagen;Import_Opstelling;WebInfo;Team_template;Kruistabel;VPSchaal;Imptabel"

Global WORKFOLDER As String
Global WORKID As Long
Global WORKFILE As String
Global STEPRESULTS As String
Global STEPDATA As String
Global LOCALSITE As String
Global LOCALHTML   As String
Global ToernooiNaam   As String
' weergave in de tabbladen, Kan avond zijn meerdere wedstrijden etc
Global PREFIX As String
'Aantal wedstrijden per avond, per dag van hoeveel krijg ik de uitslagen tegelijk binnen.

Global AANTALTEAMS   As Integer
Global AANTALSPELLENPERWEDSTRIJD As Integer
Global TEAMBYE  As Integer
Global WEDSTRIJD  As Integer
Global WEDSTRIJDENPERSESSIE  As Integer
Global COMPETITIEVORM  As Integer
Global AANTALSESSIES  As Integer
Global Prefixkopjesscorestaat As String
Global PrefixKopjeuitslagen As String
Global Suffixkopjesscorestaat As String
Global SuffixKopjeuitslagen As String
Global Voettekst As String
Global Voetlink As String
Global lngToernooi As Long
Global lngSessie As Long
Global lngTeam As Long
Global lngToernooiOld As Long
Global lngSessieOld As Long
Global intUitvoerNaarHTML As Integer
Global intExcelZichtbaar As Integer
Global ScorestaatIntern As Integer
Global ScorestaatExcel As Integer
Global BerekenAlleStaten As Integer
Global strSheetName As String
Global ActivityID As Integer
Global Sessienaam As String
Global Sessienr As Integer
Global strExcel_Folder As String
Global strHTML_Folder As String
Global strTemplate_Folder As String
Global strTemplate_File As String

'Leesbaar maken van de code
'

Public Enum Opstelling
    'sessie Teamnr Speler1 Speler2 Speler3 Speler4  Wedstrijd1  Wedstrijd2 etc
    Sessie_ = 1
    Team_nr = 2
    Speler_1 = 3
    Speler_2 = 4
    Speler_3 = 5
    Speler_4 = 6
    Wedstrijd_1 = 7
    Wedstrijd_2 = 8
    Wedstrijd_3 = 9
    Wedstrijd_4 = 10
    wedstrijd_5 = 11
    Wedstrijd_6 = 12
    Wedstrijd_7 = 13
    Wedstrijd_8 = 14
    Wedstrijd_9 = 15
    Wedstrijd_10 = 16
End Enum

Public Enum Team
    Team_nr = 1
    Team_Naam = 2
    Speler_1 = 3
    Speler_2 = 4
    Speler_3 = 5
    Speler_4 = 6
    Speler_5 = 7
    Speler_6 = 8
    Speler_7 = 9
    Speler_8 = 10
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

Public Enum Scorestaat
    spel_1 = 1
    Contract_1 = 2
    Resultaat_1 = 3
    Door_1 = 4
    score_1 = 5
    Spel_2 = 6
    Resultaat_2 = 7
    Door_2 = 8
    score_2 = 9
    saldo_ = 10
    imps_ = 11
    Imps_Wij = 12
    Imps_Zij = 13
End Enum


Public Sub InitToernooi(Id As Variant)
    Dim db As Database
    Dim rs As Recordset
    
    If Id <> lngToernooiOld Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset("select * from tblToernooi where id =" & Id)
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
            lngToernooi = Id
            lngToernooiOld = lngToernooi
        rs.Close
        db.Close
    End If
    
    
End Sub

Public Sub InitSessie(Id As Variant)
    Dim db As Database
    Dim rs As Recordset
    Set db = CurrentDb
     If Id <> lngSessieOld Then
        Set rs = db.OpenRecordset("select * from tblSessie where id =" & Id)
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
        lngSessie = Id
        lngSessieOld = lngSessie
        rs.Close
        db.Close
    End If

End Sub

Public Sub InitAll(ToernooiID As Variant, SessieID As Variant)

    Dim db As Database
    Dim rs As Recordset
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
        lngToernooi = rs!Id
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
 Dim msg, Title As String
 DgDef = vbYesNo + vbCritical + vbDefaultButton2
    msg = "Het record is gewijzigd, wilt u deze wijzigingen opslaan?"
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
   Dim lCount As Long
    Dim Soort As String
    Soort = IIf(Not (Soortbestanden = ""), Soortbestanden, "excel")
    Dim Folder As String
        
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
   Dim lCount As Long
   Dim FileChosen As Integer
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

Dim fd As Office.FileDialog
KiesModelExcelBestandEnCopieer = True
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'get the number of the button chosen
fd.InitialFileName = LaatsteLocatieModelfile
fd.AllowMultiSelect = False
fd.Title = "Kies Excel model bestand of maak zelf een excelfile aan"

Dim FileChosen As Integer
FileChosen = fd.Show
If FileChosen = -1 Then
        source = fd.SelectedItems(1)
        
        destination = InputBox("Toernooinaam: (zal dan tevens de naam van het werkbestand worden) ", "Werkbestand", "Nieuw_Toernooi")
        
        
        'Moet de copie indezelfde folder komen ?
        
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
Dim db As Database
Dim rs As Recordset

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
Dim Getal As Integer
Dim rest As Integer
Dim strKolom As String


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

Dim i As Integer
excelLinks = Split(EXCELTABNAMEN, ";")
accessLinks = Split(GEKOPPELDINACCESS, ";")
Call DeleteAllExcelLinks(varToernooi)
For i = 0 To UBound(accessLinks)
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl_" & varToernooi & "_" & accessLinks(i), strWorkfile, True, excelLinks(i) & "!"
Next
End Sub



Public Function DeleteAllExcelLinks(varToernooi As Variant)
Dim strTable As String
Dim tbl As TableDef
strTable = "tbl_" & varToernooi & "_*"
    For Each tbl In CurrentDb.TableDefs
        If tbl.name Like strTable Then
            DoCmd.DeleteObject acTable, tbl.name
        End If
   Next
End Function
Function ValidFileName(FileName As String) As Boolean
  Const sBadChar As String = "\/:*?<>|[]"""
  Dim i As Long

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