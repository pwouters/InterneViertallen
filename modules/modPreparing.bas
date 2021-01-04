Option Compare Database



Public Sub TransferUitslagenNaarSchema(WORKFOLDER, WORKFILE As Variant)
Dim ViertalUitslag, WebInf, StepRef, ScoreSheetHTML, UitslagHTML, UitslagenHTML, HTMLFolder As String
Dim rijteller, teller, teller2, kolomteller, kolomteller2, kolomoffset, i, j, intWedstrijd As Integer
Dim intWebInfo As Integer

Dim Thuis, Uit As Variant
Dim refThuis, refUit As Variant
Dim Imps_Thuis As Variant
Dim Imps_Uit As Variant
Dim VPs_Thuis As Variant
Dim VPs_Uit As Variant

Dim MySheet As Object
Dim MySchemaSheet As Object
Dim StartBook As Object
Dim Rijen() As String
Dim Kolommen() As String
Dim xlApp As Object





Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)



Set MySheet = StartBook.Sheets("TeamUitslagen")
Set MySchemaSheet = StartBook.Sheets("Schema")


rijteller = 2
teller2 = 3
Do While MySheet.Cells(rijteller, 1).Value <> ""
    intWedstrijd = MySheet.Cells(rijteller, 2).Value
    MySchemaSheet.Cells(teller2, 1) = intWedstrijd
    Kolom = 1
    Do While intWedstrijd = MySheet.Cells(rijteller, 2).Value
    Kolom = Kolom + 1
    MySchemaSheet.Cells(teller2, Kolom).Value = MySheet.Cells(rijteller, 3)
    Kolom = Kolom + 1
    MySchemaSheet.Cells(teller2, Kolom).Value = MySheet.Cells(rijteller, 4)
    rijteller = rijteller + 1
    Loop
    teller2 = teller2 + 1
Loop

   
    StartBook.Save
    Set MySheet = Nothing
    Set MySchemaSheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub

Public Sub ImportSchema(ToernID As Variant)
Dim rijteller, teller, kolomteller, i, j, avond, ronde, Thuis, Uit, internImport, gevonden As Integer
Dim tRows, tCols As Long
Dim MySheet As Object
Dim StartBook As Object
Dim Kolommen() As Variant
Dim xlApp As Object
Dim tbl As ListObject
Dim rw As ListRow
Dim cl As Range
Dim intSchemaGeladen As Integer
    
  lngSessie = DLookup("id", "tblSessie", "[ToernooiD] = " & ToernID & " and [Sessienr] = " & 1)
  Call InitAll(ToernID, lngSessie)
  

'altijd laatste sessie
    intSchemaGeladen = False
    Dim db As Database
    Dim rs As Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("select * from tblSchema where [ToernooiID] = " & ToernID & " Order By Wedstrijdronde, Paring ")
    If Not (rs.BOF And rs.EOF) Then
        rs.Close
        db.Close
        question = MsgBox("Schema is al geladen, wil je doorgaan J/N ", vbYesNo)
        intSchemaGeladen = True
        If question = vbNo Then
            Exit Sub
        End If
    End If





Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = True
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)

'test even of tabblad schema is aanwezig
If Not SheetExists("Schema", StartBook) Then
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    MsgBox ("Geen tabblad Schema in het werkbestand ")
    Exit Sub
End If


Set MySheet = StartBook.Worksheets("Schema")
If TableExists("Schema", StartBook) Or TableExists("Schema", MySheet) Then

    Set tbl = MySheet.ListObjects("Schema")
    For Each rw In tbl.ListRows
        i = 1
        Paring = 1
        For Each cl In rw.Range
            'eerste element is de ronde
            'rest zijn de paringen
            
            If i = 1 Then
                ronde = cl.Value
            Else
                If i Mod 2 = 0 Then
                 Thuis = cl.Value
                Else
                    Uit = cl.Value
                    If Not Schemageladen Then
                        rs.AddNew
                    Else
                        'zoek ronde en paring op
                        rs.MoveFirst
                        gevonden = False
                        Do While Not rs.EOF
                            If Not (rs!Wedstrijdronde = ronde And rs!Paring = Paring) Then
                            rs.MoveNext
                            Else
                                gevonden = True
                                Exit Do
                            End If
                        Loop
                        If gevonden Then
                            rs.Edit
                        Else
                            rs.AddNew
                        End If
                    End If
                    rs!ToernooiID = ToernID
                    rs!Wedstrijdronde = ronde
                    rs!Paring = Paring
                    rs!TeamThuis = Thuis
                    rs!TeamUit = Uit
                    rs.Update
                    Paring = Paring + 1
                End If
            End If
            i = i + 1
        Next cl
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

'zoek waar de tabel begint

' is er tabel aanwezig



Else

rijteller = 1
Do While MySheet.Cells(rijteller, 1).Value = ""
rijteller = rijteller + 1
If rijteller > 10 Then
   Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
    MsgBox ("Geen Schema gevonden")
    Exit Sub
End If

Loop


kolomteller = 1
teller = 1

Do While MySheet.Cells(rijteller + teller, 1).Value <> ""
Paring = 1
ronde = MySheet.Cells(rijteller + teller, 1)
'For teller = 1 To 15
       For i = 1 To AANTALTEAMS \ 2
        If Not intSchemaGeladen Then
            rs.AddNew
        Else
   'zoek ronde en paring op
            rs.MoveFirst
            gevonden = False
            Do While Not rs.EOF
                If Not (rs!Wedstrijdronde = ronde And rs!Paring = Paring) Then
                rs.MoveNext
                Else
                    gevonden = True
                    Exit Do
                End If
            Loop
            If gevonden Then
                rs.Edit
            Else
                rs.AddNew
            End If
      End If
        rs!ToernooiID = ToernID
        rs!Wedstrijdronde = ronde
        rs!Paring = Paring
        rs!TeamThuis = MySheet.Cells(rijteller + teller, i * 2)
        rs!TeamUit = MySheet.Cells(rijteller + teller, i * 2 + 1)
        rs.Update
        Paring = Paring + 1
        Next
        teller = teller + 1
 Loop
 
 
End If
rs.Close
db.Close

    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub


Public Sub MuteerExcelTeamUitslagen_Schema(ToernID As Variant, SessieID As Variant)
Dim rijteller, teller, kolomteller, i, j, avond, sessieoffset As Integer
Dim MySheet As Object
Dim StartBook As Object
Dim Kolommen() As Variant
Dim xlApp As Object
'altijd laatste sessie

Dim db As Database
Dim rs As Recordset
    
  Call InitAll(ToernID, SessieID)




Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblSchema where [ToernooiID] = " & ToernID)
If Not (rs.BOF And rs.EOF) Then
    MsgBox ("Schema is nog niet geladen of berekend ")
    Exit Sub
End If

'bepaal wedstrijdnummer
If Sessienr > 1 Then
For i = 1 To Sessienr - 1
 sessieoffset = sessieoffset + DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & lngToernooi & " and  [sessienr] = " & i)
Next
End If

Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = True
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)

'test even of tabblad schema is aanwezig
If Not SheetExists("TeamUitslagen", StartBook) Then
    MsgBox ("Geen tabblad Schema in het werkbestand ")
    Exit Sub
End If


Set MySheet = StartBook.Worksheets("TeamUitslagen")
rijteller = 1
kolomteller = 1
teller = 1

'zoek
' for j = 1 to Aantalwedstrijd per sessie
For j = sessieoffset + 1 To sessieofset + WEDSTRIJDENPERSESSIE
    Do While MySheet.Cells(rijteller + teller, 2).Value <> j And MySheet.Cells(rijteller + teller, 2).Value <> ""
        teller = teller + 1
    Loop
    
    If MySheet.Cells(rijteller + teller, 2).Value = "" Then Exit For
     Paring = 1
     Set rs = db.OpenRecordset("select * from tblSchema where [ToernooiID] = " & ToernID And " Wedstrijdronde = " & j & " Order By Wedstrijdronde ")
     rs.MoveFirst
     Do While MySheet.Cells(rijteller + teller, 2).Value = j
         MySheet.Cells(rijteller + teller, 3) = rs!TeamThuis
         MySheet.Cells(rijteller + teller, 4) = rs!TeamUit
         rs.MoveNext
         'pas paring aan
         teller = teller + 1
     Loop
     If MySheet.Cells(rijteller + teller, 2).Value = "" Then Exit For
     rs.Close
Next
 

db.Close

    StartBook.Save
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub

Public Sub AddExcelTeamUitslagen_Schema(ToernID As Variant, SessieID As Variant)
Dim rijteller, teller, kolomteller, i, j, avond, sessieoffset As Integer
Dim MySheet As Object
Dim StartBook As Object
Dim Kolommen() As Variant
Dim xlApp As Object
'altijd laatste sessie

Dim db As Database
Dim rs As Recordset
    
  Call InitAll(ToernID, SessieID)




Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblSchema where [ToernooiID] = " & ToernID)
If rs.BOF And rs.EOF Then
    MsgBox ("Schema is nog niet geladen of berekend ")
    Exit Sub
End If

'bepaal wedstrijdnummer
If Sessienr > 1 Then
For i = 1 To Sessienr - 1
 sessieoffset = sessieoffset + DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & lngToernooi & " and  [sessienr] = " & i)
Next
End If

Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = True
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)

'test even of tabblad schema is aanwezig
If Not SheetExists("TeamUitslagen", StartBook) Then
    MsgBox ("Geen tabblad Schema in het werkbestand ")
    Exit Sub
End If


Set MySheet = StartBook.Worksheets("TeamUitslagen")
rijteller = 1
kolomteller = 1
teller = 1
j = MySheet.Cells(rijteller + teller, 2).Value

    Do While MySheet.Cells(rijteller + teller, 2).Value <> ""
        j = MySheet.Cells(rijteller + teller, 2).Value
        teller = teller + 1
    Loop

If j > sessieoffset Then
    MsgBox ("Schema is al geladen of berekend ")
    Exit Sub
End If
'test even wat de laatste wedstrijd was


'zoek
' for j = 1 to Aantalwedstrijd per sessie
For j = sessieoffset + 1 To sessieoffset + WEDSTRIJDENPERSESSIE
     Set rs = db.OpenRecordset("select * from tblSchema where [ToernooiID] = " & ToernID & " And  [Wedstrijdronde] = " & j & " Order By Paring ")
     rs.MoveFirst
     Do While Not rs.EOF
         MySheet.Cells(rijteller + teller, 1).Value = Sessienr
         MySheet.Cells(rijteller + teller, 2).Value = j
         MySheet.Cells(rijteller + teller, 3).Value = rs!TeamThuis
         MySheet.Cells(rijteller + teller, 4).Value = rs!TeamUit
         rs.MoveNext
         'pas paring aan
         teller = teller + 1
     Loop
     rs.Close
Next
 

db.Close

    StartBook.Save
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing


End Sub

Public Sub AddInternTeamUitslagen_Schema(ToernID As Variant, SessieID As Variant)
Dim rijteller, teller, kolomteller, i, j, avond, sessieoffset As Integer


'altijd laatste sessie

Dim db As Database
Dim rs As Recordset
Dim us As Recordset
  Call InitAll(ToernID, SessieID)




Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblSchema where [ToernooiID] = " & ToernID)
If rs.BOF And rs.EOF Then
    MsgBox ("Schema is nog niet geladen of berekend ")
    Exit Sub
End If

'bepaal wedstrijdnummer
If Sessienr > 1 Then
For i = 1 To Sessienr - 1
 sessieoffset = sessieoffset + DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & lngToernooi & " and  [sessienr] = " & i)
Next
End If


Set us = db.OpenRecordset("select * from tblUitslagen where [ToernooiID] = " & ToernID & " And  [Wedstrijdnr] = " & sessieoffset + 1)

If Not (us.BOF And us.EOF) Then
      MsgBox ("Uitslagen schema is nog geladen")
    Exit Sub
End If


' for j = 1 to Aantalwedstrijd per sessie
For j = sessieoffset + 1 To sessieoffset + WEDSTRIJDENPERSESSIE
     Set rs = db.OpenRecordset("select * from tblSchema where [ToernooiID] = " & ToernID & " And  [Wedstrijdronde] = " & j & " Order By Paring ")
     rs.MoveFirst
     Do While Not rs.EOF
        us.AddNew
         us!SessieID = SessieID
         us!TeamIDThuis = DLookup("id", "tblTeams", "[ToernooiID] = " & ToernID & " and [Teamnr] = " & rs!TeamThuis)
         us!TeamIDUit = DLookup("id", "tblTeams", "[ToernooiID] = " & ToernID & " and [Teamnr] = " & rs!TeamUit)
         us!Wedstrijdnr = j
         us!ToernooiID = ToernID
         us.Update
         rs.MoveNext
         'pas paring aan
     Loop
     rs.Close
Next
 
us.Close

db.Close

End Sub

Public Function GespeeldeWedstrijden(varSessie, varToernooi As Variant) As Integer
Dim sessieoffset, i As Integer
sessieoffset = 0
If varSessie > 1 Then
For i = 1 To varSessie - 1
    sessieoffset = sessieoffset + DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & varToernooi & " and  [sessienr] = " & i)
Next
GespeeldeWedstrijden = sessieoffset
End If

End Function

Public Function GespeeldeWedstrijden_vanaf(varSessieID, varToernooi As Variant) As String
Dim sessieoffset, i As Integer
Dim intSessienr As Integer
sessieoffset = 0
intSessienr = DLookup("sessienr", "tblSessie", "id = " & varSessieID)
If intSessienr > 1 Then
For i = 1 To intSessienr - 1
    sessieoffset = sessieoffset + DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & varToernooi & " and  [sessienr] = " & i)
Next
End If
GespeeldeWedstrijden_vanaf = sessieoffset + 1
End Function
Public Function GespeeldeWedstrijden_tot(varSessieID, varToernooi As Variant) As String
Dim sessieoffset, i As Integer
Dim intSessienr As Integer
sessieoffset = 0
intSessienr = DLookup("sessienr", "tblSessie", "id = " & varSessieID)
For i = 1 To intSessienr
    sessieoffset = sessieoffset + DLookup("AantalWedstrijdenPerSessie", "tblSessie", "[ToernooID] = " & varToernooi & " and  [sessienr] = " & i)
Next
GespeeldeWedstrijden_tot = sessieoffset
End Function