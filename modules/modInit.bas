Option Compare Database
Public MyBookmark As Variant
Public CtrlVisible As Integer
Public recordteller As Long
Public lngPK As Long



Global WORKFOLDER As String
Global WORKID As Long
Global WORKFILE As String
Global STEPRESULTS As String
Global STEPDATA As String
Global LOCALSITE As String
Global LOCALHTML   As String

' weergave in de tabbladen, Kan avond zijn meerdere wedstrijden etc
Global PREFIX As String
'Aantal wedstrijden per avond, per dag van hoeveel krijg ik de uitslagen tegelijk binnen.

Global AANTALTEAMS   As Integer
Global AANTALSPELLENPERWEDSTRIJD As Integer
Global TEAMBYE  As Integer
Global WEDSTRIJD  As Integer
Global WEDSTRIJDENPERSESSIE  As Integer
Global COMPETITIEVORM  As Integer
Global Prefixkopjesscorestaat As String
Global PrefixKopjeuitslagen As String
Global Suffixkopjesscorestaat As String
Global SuffixKopjeuitslagen As String
Global Voettekst As String
Global Voetlink As String
Global lngToernooi As Long
Global lngSessie As Long
Global lngToernooiOld As Long
Global lngSessieOld As Long
Global intUitvoerNaarHTML As Integer
Global intExcelZichtbaar As Integer

Global ActivityID As Integer
Global Sessienaam As String
Global Sessienr As Integer
Global strExcel_Folder As String
Global strHTML_Folder As String
Global strTemplate_Folder As String
Global strTemplate_File As String

Public Sub InitToernooi(Id As Variant)
    Dim db As Database
    Dim rs As Recordset
    
    If Id <> lngToernooiOld Then
        Set db = CurrentDb
        Set rs = db.OpenRecordset("select * from tblToernooi where id =" & Id)
            rs.MoveFirst
            WORKID = rs.Fields("ID")
            WORKFOLDER = rs.Fields("WORKFOLDER")
            WORKFILE = rs.Fields("WORKFILE")
            STEPRESULTS = rs.Fields("STEPRESULTS")
            STEPDATA = rs.Fields("STEPDATA")
            LOCALSITE = rs.Fields("LOCALSITE")
            LOCALHTML = rs.Fields("LOCALHTML")
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
        WORKFOLDER = rs.Fields("WORKFOLDER")
        WORKFILE = rs.Fields("WORKFILE")
        STEPRESULTS = rs.Fields("STEPRESULTS")
        STEPDATA = rs.Fields("STEPDATA")
        LOCALSITE = rs.Fields("LOCALSITE")
        LOCALHTML = rs.Fields("LOCALHTML")
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

Sub KiesModelExcelBestandEnCopieer(InitialFile As Variant)
Dim source, destination As String

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
'get the number of the button chosen
fd.InitialFileName = InitialFile

Dim FileChosen As Integer
FileChosen = fd.Show
If FileChosen = -1 Then
        source = fd.SelectedItems(1)
        destination = InputBox("Kies nieuw naam Werkbestand ", "Werkbestand", "Nieuw_Toernooi.xlsx")
        
        
        'Moet de copie indezelfde folder komen ?
        
        'indien nee
        
        'folder dialoog
        
        'strFolder
        
        
        
        destination = FolderFromPath(source) & destination
        FileCopy source, destination
    
End If

End Sub





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
Public Function FileNameFromPath(strFullPath As Variant) As String
    FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function

Public Function FolderFromPath(strFullPath As Variant) As String
    FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
End Function