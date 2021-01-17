Option Compare Database

' Zwitsers indelen met gebruikmaken van het werkbestand
'
' gebruik maken van de kruistabel om de volgende ronde te berekenen
' indien gebruikmakend van indelen op de stand van ronde terug moet de uitslagen tab gescand worden

Private Stand()         As Integer
Private Total()         As Double
Private intStand        As Integer
Private intRondesGesp   As Integer
Private intRondeStand   As Integer
Private H()             As Integer
Private Id()            As Integer
Private VanAfViert      As Integer
Private TotEnMet        As Integer
Private T, U            As Integer
Private Paringen        As Integer
Private GeenCombinatie  As Integer
Private strIndeling     As String

Public Function BepaalZwitsers(varTeams As Variant, rondestand As Variant, Aantalrondesgespeeld As Variant, strWorkfolder As Variant, strWorkfile As Variant, Optional MaxRonde As Variant) As Variant
    ' werkt vanuit excel bestand
    Dim Teller          As Integer
    Dim kolomteller     As Integer
    Dim i, j            As Integer
    Dim MySheet         As Object
    Dim StartBook       As Object
    Dim xlApp           As Object
    
    'Variabelen gebruikt bij zwitsers
    
    Dim WedstrijdenGespeeld()    As Integer
    Dim Tegenstanders() As Integer
    
    Dim db              As Database
    Dim rs              As Recordset
    Dim sql             As String
    
    If IsMissing(MaxRonde) Then
        MaxRonde = 16
    End If
    
    intRondesGesp = Int(Aantalrondesgespeeld)
    intRondeStand = Int(rondestand)
    
    ReDim Stand(varTeams)
    ReDim Total(varTeams)
    ReDim Rang(varTeams)
    ReDim Gemiddelde(varTeams)
    ReDim Gespeeld(varTeams, varTeams)
    ReDim H(varTeams)
    ReDim Id(varTeams \ 2, 1)
    ReDim Tegenstanders(varTeams, MaxRonde)
    ReDim WedstrijdenGespeeld(varTeams)
    
    If Not fnExists(strWorkfolder & strWorkfile) Then
        MsgBox ("Werkbestand strWorkfolder & strWorkfile Is niet gevonden")
        Exit Function
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(strWorkfolder & strWorkfile)
    
    If Not SheetExists("Kruistabel", StartBook) Then
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab Kruistabel Is niet gevonden")
        Exit Function
    End If
    
    If Not SheetExists("TeamUitslagen", StartBook) Then
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab TeamUitslagen Is niet gevonden")
        Exit Function
    End If
    
    ' indien een variable indeling
    
    Call ImportTempTeamUitslagenTabel(strWorkfolder & strWorkfile)
    Set db = CurrentDb
    sql = ""
    sql = sql & "Select * from tbl_temp_Uitslagen "
    sql = sql & "Where wedstrijd <= " & intRondesGesp & " "
    sql = sql & "Order by Avond, Wedstrijd;"
    Set rs = db.OpenRecordset(sql)
    If rs.BOF And rs.EOF Then
        rs.Close
        db.Close
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab TeamUitslagen heeft geen uitslagen")
        Exit Function
    End If
    rs.MoveFirst
    
    Do While Not rs.EOF
        Tegenstanders(rs.Fields(f_Thuis_nr), rs.Fields(f_Wedstrijd_nr)) = rs.Fields(f_Uit_nr)
        Tegenstanders(rs.Fields(f_Uit_nr), rs.Fields(f_Wedstrijd_nr)) = rs.Fields(f_Thuis_nr)
        If Not IsNull(rs.Fields(f_Thuis_VPs)) Then
            If rs.Fields(f_Wedstrijd_nr) <= intRondeStand Then
                Total(rs.Fields(f_Thuis_nr)) = Total(rs.Fields(f_Thuis_nr)) + rs.Fields(f_Thuis_VPs)
                Total(rs.Fields(f_Uit_nr)) = Total(rs.Fields(f_Uit_nr)) + rs.Fields(f_Uit_VPs)
                WedstrijdenGespeeld(rs.Fields(f_Thuis_nr)) = WedstrijdenGespeeld(rs.Fields(f_Thuis_nr)) + 1
                WedstrijdenGespeeld(rs.Fields(f_Uit_nr)) = WedstrijdenGespeeld(rs.Fields(f_Uit_nr)) + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Close
    
    For i = 1 To varTeams
        If WedstrijdenGespeeld(i) > 0 Then
            Gemiddelde(i) = Total(i) / WedstrijdenGespeeld(i)
        End If
    Next
    'Bereken rang
    For i = 1 To varTeams
        Rang(i) = i
    Next
    Call Sorteer(Gemiddelde(), Rang(), True)
    
    For i = 1 To varTeams
        Stand(Rang(i)) = i
    Next
    
    For Teller = 1 To varTeams
        For kolomteller = 1 To intRondesGesp
            Gespeeld(Teller, Tegenstanders(Teller, kolomteller)) = True
            Gespeeld(Tegenstanders(Teller, kolomteller), Teller) = True
        Next
    Next
    
    ' start zwitsers
    Call RekenZwitsers(varTeams)
    
    strIndeling = ""
    Dim TotalThuis, TotalUit    As String
    
    If GeenCombinatie Then
        strIndeling = "Geen Combinatie"
    Else
        For i = 1 To varTeams \ 2
            TotalThuis = Format(Total(Id(i, 0)), "#0.00")
            TotalThuis = String(6 - Len(TotalThuis), " ") & TotalThuis
            TotalUit = Format(Total(Id(i, 1)), "#0.00")
            TotalUit = String(6 - Len(TotalUit), " ") & TotalUit
            
            strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & "  |  Total " & TotalThuis & " - " & TotalUit
            strIndeling = strIndeling & " |  " & Format(Rang(Id(i, 0)), "00") & " - " & Format(Rang(Id(i, 1)), "00") & vbCr
            
        Next
    End If
    
    BepaalZwitsers = Id
    
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
End Function


Public Function BepaalDeens(varTeams As Variant, rondestand As Variant, Aantalrondesgespeeld As Variant, strWorkfolder As Variant, strWorkfile As Variant, Optional MaxRonde As Variant) As Variant
    ' werkt vanuit excel bestand
    Dim Teller          As Integer
    Dim kolomteller     As Integer
    Dim i, j            As Integer
    Dim MySheet         As Object
    Dim StartBook       As Object
    Dim xlApp           As Object
    
    'Variabelen gebruikt bij zwitsers
    
    Dim WedstrijdenGespeeld()    As Integer
    Dim Tegenstanders() As Integer
    
    Dim db              As Database
    Dim rs              As Recordset
    Dim sql             As String
    
    If IsMissing(MaxRonde) Then
        MaxRonde = 16
    End If
    
    intRondesGesp = Int(Aantalrondesgespeeld)
    intRondeStand = Int(rondestand)
    
    ReDim Stand(varTeams)
    ReDim Total(varTeams)
    ReDim Rang(varTeams)
    ReDim Gemiddelde(varTeams)
    ReDim Gespeeld(varTeams, varTeams)
    ReDim H(varTeams)
    ReDim Id(varTeams \ 2, 1)
    ReDim Tegenstanders(varTeams, MaxRonde)
    ReDim WedstrijdenGespeeld(varTeams)
    
    If Not fnExists(strWorkfolder & strWorkfile) Then
        MsgBox ("Werkbestand strWorkfolder & strWorkfile Is niet gevonden")
        Exit Function
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(strWorkfolder & strWorkfile)
    
    If Not SheetExists("Kruistabel", StartBook) Then
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab Kruistabel Is niet gevonden")
        Exit Function
    End If
    
    If Not SheetExists("TeamUitslagen", StartBook) Then
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab TeamUitslagen Is niet gevonden")
        Exit Function
    End If
    
    ' indien een variable indeling
    
    Call ImportTempTeamUitslagenTabel(strWorkfolder & strWorkfile)
    Set db = CurrentDb
    sql = ""
    sql = sql & "Select * from tbl_temp_Uitslagen "
    sql = sql & "Where wedstrijd <= " & intRondesGesp & " "
    sql = sql & "Order by Avond, Wedstrijd;"
    Set rs = db.OpenRecordset(sql)
    If rs.BOF And rs.EOF Then
        rs.Close
        db.Close
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab TeamUitslagen heeft geen uitslagen")
        Exit Function
    End If
    rs.MoveFirst
    
    Do While Not rs.EOF
        If Not IsNull(rs.Fields(f_Thuis_VPs)) Then
            If rs.Fields(f_Wedstrijd_nr) <= intRondeStand Then
                Total(rs.Fields(f_Thuis_nr)) = Total(rs.Fields(f_Thuis_nr)) + rs.Fields(f_Thuis_VPs)
                Total(rs.Fields(f_Uit_nr)) = Total(rs.Fields(f_Uit_nr)) + rs.Fields(f_Uit_VPs)
                WedstrijdenGespeeld(rs.Fields(f_Thuis_nr)) = WedstrijdenGespeeld(rs.Fields(f_Thuis_nr)) + 1
                WedstrijdenGespeeld(rs.Fields(f_Uit_nr)) = WedstrijdenGespeeld(rs.Fields(f_Uit_nr)) + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Close
    
    For i = 1 To varTeams
        If WedstrijdenGespeeld(i) > 0 Then
            Gemiddelde(i) = Total(i) / WedstrijdenGespeeld(i)
        End If
    Next
    'Bereken rang
    For i = 1 To varTeams
        Rang(i) = i
    Next
    Call Sorteer(Gemiddelde(), Rang(), True)
    
    For i = 1 To varTeams
        Stand(Rang(i)) = i
    Next
 
'Deens
  
For i = 1 To varTeams \ 2
    Id(i, 0) = Stand((i - 1) * 2 + 1)
    Id(i, 1) = Stand(i * 2)
Next
  
    
    
    strIndeling = ""
    Dim TotalThuis, TotalUit    As String
    
    If GeenCombinatie Then
        strIndeling = "Geen Combinatie"
    Else
        For i = 1 To varTeams \ 2
            TotalThuis = Format(Total(Id(i, 0)), "#0.00")
            TotalThuis = String(6 - Len(TotalThuis), " ") & TotalThuis
            TotalUit = Format(Total(Id(i, 1)), "#0.00")
            TotalUit = String(6 - Len(TotalUit), " ") & TotalUit
            
            strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & "  |  Total " & TotalThuis & " - " & TotalUit
            strIndeling = strIndeling & " |  " & Format(Rang(Id(i, 0)), "00") & " - " & Format(Rang(Id(i, 1)), "00") & vbCr
            
        Next
    End If
    
    BepaalDeens = Id
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
End Function

'sorteer routine nog oude doos gemaakt mid 80 ers

Public Sub Sorteer(ByRef rx() As Double, ByRef rv() As Integer, Optional intVolgorde As Integer)
Dim lw, rw          As Integer
Dim s1, s2, s3, sx, temp, aantal As Integer
Dim sl()            As Integer
Dim sr()            As Integer
Dim hlp()           As Integer


'is eigenlijk de stack (s3) bij 5.000.000 items zal hij ongeveer 14 moeten zijn
'is geen recursie quicksort

ReDim sl(20)
ReDim sr(20)

aantal = UBound(rv)
ReDim hlp(aantal)


s3 = 1: sl(1) = 1: sr(1) = aantal
Do
    lw = sl(s3): rw = sr(s3): s3 = s3 - 1
    Do
        s1 = lw: s2 = rw: sx = rv(Int((lw + rw) / 2))
        Do
            While rx(rv(s1)) < rx(sx)
                s1 = s1 + 1
            Wend
            While rx(rv(s2)) > rx(sx)
                s2 = s2 - 1
            Wend
            If s1 <= s2 Then
                temp = rv(s1)
                rv(s1) = rv(s2)
                rv(s2) = temp
                s1 = s1 + 1
                s2 = s2 - 1
            End If
        Loop Until s1 > s2
        If s2 - lw >= rw - s1 Then
            If lw < s2 Then
                s3 = s3 + 1
                sl(s3) = lw
                sr(s3) = s2
            End If
            lw = s1
        Else
            If s1 < rw Then
                s3 = s3 + 1
                sl(s3) = s1
                sr(s3) = rw
            End If
            rw = s2
        End If
    Loop Until lw >= rw
Loop Until s3 = 0

For s1 = 1 To aantal
    If intVolgorde Then
        hlp(s1) = rv(aantal - s1 + 1)
    Else
        hlp(s1) = rv(s1)
    End If
Next
For s1 = 1 To aantal
    rv(hlp(s1)) = s1
Next


End Sub


Public Sub RekenZwitsers(varTeams)
    Dim P1          As Integer
    
    x = 0
    
    ReDim H(varTeams)
    
600:
    GeenCombinatie = 0
    'TOTENMET
    VanAfViert = 1
    TotEnMet = varTeams
    
    T = VanAfViert
    Paringen = 0
    For i = VanAfViert To TotEnMet: H(i) = False: Next: H(VanAfViert) = True
        
        U = VanAfViert
        
        Do
            
            U = U + 1
            If U > TotEnMet And T = VanAfViert Then
                GeenCombinatie = -1: Exit Do
            End If
            If U > TotEnMet Then
                
                If x = 0 Then P1 = T - 1: AF = TotEnMet - P1
                intStand = Stand(P1)
                Stand(P1) = Stand(P1 + 1)
                Stand(P1 + 1) = intStand
                P1 = P1 + 1: x = -1
                If P1 > TotEnMet - 1 Then AF = AF + 1: P1 = TotEnMet - AF
                If P1 < VanAfViert Then
                    GeenCombinatie = -1
                    Exit Do
                End If
                
                'Wis indeling
                For i = 1 To Paringen
                    Gespeeld(Id(i, 0), Id(i, 1)) = False
                    Gespeeld(Id(i, 1), Id(i, 0)) = False
                    Id(i, 0) = 0
                    Id(i, 1) = 0
                Next
                
                GoTo 600
                
            Else
                If H(U) = -1 Then GoTo 610
            End If
            
            If Gespeeld(Stand(T), Stand(U)) Then GoTo 610
            
            Paringen = Paringen + 1
            If Stand(T) = 0 Then Stand(T) = 16
            If Stand(U) = 0 Then Stand(U) = 16
            Id(Paringen, 0) = Stand(T)
            Id(Paringen, 1) = Stand(U)
            
            Gespeeld(Stand(T), Stand(U)) = True
            Gespeeld(Stand(U), Stand(T)) = True
            
            H(U) = -1
            
            If Paringen = TotEnMet \ 2 Then Exit Do
            Do
                T = T + 1
            Loop Until H(T) <> -1
            H(T) = -1
610:
        Loop
        
    End Sub


Public Function RekenRandom(varTeams As Variant, Aantalrondesgespeeld As Variant, strWorkfolder As Variant, strWorkfile As Variant, Optional MaxRonde As Variant) As Variant
    Dim Teller          As Integer
    Dim kolomteller     As Integer
    Dim i, j, n, k         As Integer
    Dim intParingen     As Integer
    Dim intNietgespeeld As Integer
    Dim Thuis, Uit      As Integer
    Dim MySheet         As Object
    Dim StartBook       As Object
    Dim xlApp           As Object
    
    'Variabelen gebruikt bij zwitsers
    
    Dim WedstrijdenGespeeld()    As Integer
    Dim Tegenstanders() As Integer
    
    Dim db              As Database
    Dim rs              As Recordset
    Dim sql             As String
    
    If IsMissing(MaxRonde) Then
        MaxRonde = 16
    End If
    intRondesGesp = Int(Aantalrondesgespeeld)
    
    ReDim Gespeeld(varTeams, varTeams)
    ReDim H(varTeams)
    ReDim Id(varTeams \ 2, 1)
    ReDim Tegenstanders(varTeams, MaxRonde)
    ReDim WedstrijdenGespeeld(varTeams)

    
    If Not fnExists(strWorkfolder & strWorkfile) Then
        MsgBox ("Werkbestand strWorkfolder & strWorkfile Is niet gevonden")
        Exit Function
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(strWorkfolder & strWorkfile)
    
    If Not SheetExists("Kruistabel", StartBook) Then
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab Kruistabel Is niet gevonden")
        Exit Function
    End If
    
    If Not SheetExists("TeamUitslagen", StartBook) Then
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab TeamUitslagen Is niet gevonden")
        Exit Function
    End If
    
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
    Call ImportTempTeamUitslagenTabel(strWorkfolder & strWorkfile)
    Set db = CurrentDb
    sql = ""
    sql = sql & "Select * from tbl_temp_Uitslagen "
    sql = sql & "Where wedstrijd <= " & intRondesGesp & " "
    sql = sql & "Order by Avond, Wedstrijd;"
    Set rs = db.OpenRecordset(sql)
    If rs.BOF And rs.EOF Then
        rs.Close
        db.Close
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        MsgBox ("tab TeamUitslagen heeft geen uitslagen")
    Else
        rs.MoveFirst
        Do While Not rs.EOF
            Tegenstanders(rs.Fields(f_Thuis_nr), rs.Fields(f_Wedstrijd_nr)) = rs.Fields(f_Uit_nr)
            Tegenstanders(rs.Fields(f_Uit_nr), rs.Fields(f_Wedstrijd_nr)) = rs.Fields(f_Thuis_nr)
            rs.MoveNext
        Loop
        rs.Close
        db.Close
    End If
    
    If intRondesGesp > 0 Then
        For Teller = 1 To varTeams
            For kolomteller = 1 To intRondesGesp
                Gespeeld(Teller, Tegenstanders(Teller, kolomteller)) = True
                Gespeeld(Tegenstanders(Teller, kolomteller), Teller) = True
            Next
        Next
    End If
    
    'maak een array van varTeams
    'do zolang er paringen zijn
    
    Dim numCollection As New Collection
    Randomize
    Do
        With numCollection
            'nieuwe reeks
            For i = 1 To varTeams
                .Add i
            Next
            k = 0
            Do
                ' thuis ploeg
                k = k + 1
                n = Int(Rnd * (.Count - 1) + 1)
                Thuis = numCollection(n)
                Id(k, 0) = Thuis
                .Remove n
                'trek uit ploeg
                Nietgespeeld = False
                Do While .Count > 0
                    n = Int(Rnd * (.Count - 1) + 1)
                    Uit = numCollection(n)
                    If Not Gespeeld(Id(k, 0), Uit) Then
                        Nietgespeeld = True
                        Id(k, 1) = Uit
                        .Remove n
                        Exit Do
                    End If
                Loop
                If k = varTeams \ 2 And Nietgespeeld Then Exit Do
            Loop
        End With
        Set numCollection = New Collection
        If k = varTeams \ 2 And Nietgespeeld Then Exit Do
    Loop
    
    'Display nieuwe indeling
        strIndeling = ""
For i = 1 To varTeams \ 2
    strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & vbCr
Next
    'toevoegen aan
 RekenRandom = Id
End Function

Public Function BepaalZwitsersIntern(varTeams As Variant, rondestand As Variant, Aantalrondesgespeeld As Variant, ToernID As Variant, Optional MaxRonde As Variant) As Variant

Dim Teller          As Integer
Dim kolomteller     As Integer
Dim i, j            As Integer
Dim Thuisnr, UitNr  As Integer

Dim MySheet         As Object



'Variabelen gebruikt bij zwitsers


Dim WedstrijdenGespeeld()    As Integer
Dim Tegenstanders() As Integer

Dim db              As Database
Dim rs              As Recordset
Dim sql             As String

If IsMissing(MaxRonde) Then
    MaxRonde = 16
End If
intRondesGesp = Int(Aantalrondesgespeeld)
intRondeStand = Int(rondestand)
ReDim Stand(varTeams)
ReDim Total(varTeams)
ReDim Rang(varTeams)
ReDim Gemiddelde(varTeams)
ReDim Gespeeld(varTeams, varTeams)
ReDim H(varTeams)
ReDim Id(varTeams \ 2, 1)
ReDim Tegenstanders(varTeams, MaxRonde)
ReDim WedstrijdenGespeeld(varTeams)



' indien een variable indeling


    Set db = CurrentDb
    sql = ""
    sql = sql & "SELECT tblUitslagen.Wedstrijdnr, tblUitslagen.ImpsThuis, tblUitslagen.ImpsUit, "
    sql = sql & "tblUitslagen.VpsThuis, tblUitslagen.VpsUit, tblTeams.Teamnr AS Thuisnr, tblTeams_1.Teamnr AS Uitnr "
    sql = sql & "FROM (tblTeams INNER JOIN tblUitslagen ON (tblTeams.id = tblUitslagen.TeamIDThuis) AND (tblTeams.id = tblUitslagen.TeamIDThuis)) INNER JOIN tblTeams AS tblTeams_1 ON tblUitslagen.TeamIDUit = tblTeams_1.id "
    sql = sql & "WHERE tblUitslagen.ToernooiID = " & ToernID & " And tblUitslagen.Wedstrijdnr <= " & intRondesGesp & " "
    sql = sql & "ORDER BY tblUitslagen.id, tblUitslagen.Wedstrijdnr;"
    
    Set rs = db.OpenRecordset(sql)
    
    If rs.BOF And rs.EOF Then
        rs.Close
        db.Close
       
        MsgBox ("tab TeamUitslagen heeft geen uitslagen")
        Exit Function
    End If
    rs.MoveFirst
    
    Do While Not rs.EOF
        
        Tegenstanders(rs!Thuisnr, rs!Wedstrijdnr) = rs!UitNr
        Tegenstanders(rs!UitNr, rs!Wedstrijdnr) = rs!Thuisnr
        If Not IsNull(rs!VpsThuis) Then
                If rs!Wedstrijdnr <= intRondeStand Then
                    Total(rs!Thuisnr) = Total(rs!Thuisnr) + rs!VpsThuis
                    Total(rs!UitNr) = Total(rs!UitNr) + rs!VpsUit
                    WedstrijdenGespeeld(rs!Thuisnr) = WedstrijdenGespeeld(rs!Thuisnr) + 1
                    WedstrijdenGespeeld(rs!UitNr) = WedstrijdenGespeeld(rs!UitNr) + 1
                End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Close
    
    For i = 1 To varTeams
        If WedstrijdenGespeeld(i) > 0 Then
            Gemiddelde(i) = Total(i) / WedstrijdenGespeeld(i)
        End If
    Next
    'Bereken rang
    For i = 1 To varTeams
        Rang(i) = i
    Next
    Call Sorteer(Gemiddelde(), Rang(), True)
    
    For i = 1 To varTeams
        Stand(Rang(i)) = i
    Next
    
    
    For Teller = 1 To varTeams
        For kolomteller = 1 To intRondesGesp
             Gespeeld(Teller, Tegenstanders(Teller, kolomteller)) = True
             Gespeeld(Tegenstanders(Teller, kolomteller), Teller) = True
        Next
   Next
    
  
' start zwitsers
Call RekenZwitsers(varTeams)

strIndeling = ""
Dim TotalThuis, TotalUit    As String

If GeenCombinatie Then
    strIndeling = "Geen Combinatie"
Else
    For i = 1 To varTeams \ 2
        TotalThuis = Format(Total(Id(i, 0)), "#0.00")
        TotalThuis = String(6 - Len(TotalThuis), " ") & TotalThuis
        TotalUit = Format(Total(Id(i, 1)), "#0.00")
        TotalUit = String(6 - Len(TotalUit), " ") & TotalUit
        
        
        strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & "  |  Total " & TotalThuis & " - " & TotalUit
        strIndeling = strIndeling & " |  " & Format(Rang(Id(i, 0)), "00") & " - " & Format(Rang(Id(i, 1)), "00") & vbCr
    
    Next
End If

BepaalZwitsersIntern = Id


End Function

Public Function RekenRandomIntern(varTeams As Variant, Aantalrondesgespeeld As Variant, ToernID As Variant, Optional MaxRonde As Variant) As Variant
    Dim Teller          As Integer
    Dim kolomteller     As Integer
    Dim i, j, n, k         As Integer
    Dim intParingen     As Integer
    Dim intNietgespeeld As Integer
    Dim Thuis, Uit      As Integer
    
    'Variabelen gebruikt bij zwitsers
    
    Dim WedstrijdenGespeeld()    As Integer
    Dim Tegenstanders() As Integer
    
    Dim db              As Database
    Dim rs              As Recordset
    Dim sql             As String
    
    If IsMissing(MaxRonde) Then
        MaxRonde = 16
    End If
    If IsNull(Aantalrondesgespeeld) Then
        Aantalrondesgespeeld = 0
    End If
    intRondesGesp = Int(Aantalrondesgespeeld)
    
    ReDim Gespeeld(varTeams, varTeams)
    ReDim H(varTeams)
    ReDim Id(varTeams \ 2, 1)
    ReDim Tegenstanders(varTeams, MaxRonde)
    ReDim WedstrijdenGespeeld(varTeams)

    
If intRondesGesp > 0 Then
    Set db = CurrentDb
    sql = ""
    sql = sql & "SELECT tblUitslagen.Wedstrijdnr, tblUitslagen.ImpsThuis, tblUitslagen.ImpsUit, "
    sql = sql & "tblUitslagen.VpsThuis, tblUitslagen.VpsUit, tblTeams.Teamnr AS Thuisnr, tblTeams_1.Teamnr AS Uitnr "
    sql = sql & "FROM (tblTeams INNER JOIN tblUitslagen ON (tblTeams.id = tblUitslagen.TeamIDThuis) AND (tblTeams.id = tblUitslagen.TeamIDThuis)) INNER JOIN tblTeams AS tblTeams_1 ON tblUitslagen.TeamIDUit = tblTeams_1.id "
    sql = sql & "WHERE tblUitslagen.ToernooiID = " & ToernID & " And tblUitslagen.Wedstrijdnr <= " & intRondesGesp & " "
    sql = sql & "ORDER BY tblUitslagen.id, tblUitslagen.Wedstrijdnr;"
    
    Set rs = db.OpenRecordset(sql)
    
    If rs.BOF And rs.EOF Then
        rs.Close
        db.Close
        MsgBox ("tab TeamUitslagen heeft geen uitslagen")
    Else
        rs.MoveFirst
        Do While Not rs.EOF
            Tegenstanders(rs!Thuisnr, rs!Wedstrijdnr) = rs!UitNr
            Tegenstanders(rs!UitNr, rs!Wedstrijdnr) = rs!Thuisnr
            rs.MoveNext
        Loop
        rs.Close
        db.Close
    End If
    
    If intRondesGesp > 0 Then
        For Teller = 1 To varTeams
            For kolomteller = 1 To intRondesGesp
                Gespeeld(Teller, Tegenstanders(Teller, kolomteller)) = True
                Gespeeld(Tegenstanders(Teller, kolomteller), Teller) = True
            Next
        Next
    End If
 End If
 
    'maak een array van varTeams
    'do zolang er paringen zijn
    
    Dim numCollection As New Collection
    
    Do
        With numCollection
            'nieuwe reeks
            For i = 1 To varTeams
                .Add i
            Next
            k = 0
            Do
                ' thuis ploeg
                k = k + 1
                Randomize
                n = Int(Rnd * (.Count) + 1)
                Thuis = numCollection(n)
                Id(k, 0) = Thuis
                .Remove n
                'trek uit ploeg
                Nietgespeeld = False
                Do While .Count > 0
                    n = Int(Rnd * (.Count) + 1)
                    Uit = numCollection(n)
                    If Not Gespeeld(Id(k, 0), Uit) Then
                        Nietgespeeld = True
                        Id(k, 1) = Uit
                        .Remove n
                        Exit Do
                    End If
                Loop
                If k = varTeams \ 2 And Nietgespeeld Then Exit Do
            Loop
        End With
        Set numCollection = New Collection
        If k = varTeams \ 2 And Nietgespeeld Then Exit Do
    Loop
    strIndeling = ""
For i = 1 To varTeams \ 2
    strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & vbCr
Next
    'Display nieuwe indeling
    
    'toevoegen aan
 RekenRandomIntern = Id
End Function
Public Function BepaalDeensIntern(varTeams As Variant, Aantalrondesgespeeld As Variant, ToernID As Variant, Optional MaxRonde As Variant) As Variant

Dim Teller          As Integer
Dim kolomteller     As Integer
Dim i, j            As Integer
Dim Thuisnr, UitNr  As Integer

Dim MySheet         As Object



'Variabelen gebruikt bij zwitsers


Dim WedstrijdenGespeeld()    As Integer
Dim Tegenstanders() As Integer

Dim db              As Database
Dim rs              As Recordset
Dim sql             As String

If IsMissing(MaxRonde) Then
    MaxRonde = 16
End If
intRondesGesp = Int(Aantalrondesgespeeld)
intRondeStand = Int(Aantalrondesgespeeld)
ReDim Stand(varTeams)
ReDim Total(varTeams)
ReDim Rang(varTeams)
ReDim Gemiddelde(varTeams)
ReDim Gespeeld(varTeams, varTeams)
ReDim H(varTeams)
ReDim Id(varTeams \ 2, 1)
ReDim Tegenstanders(varTeams, MaxRonde)
ReDim WedstrijdenGespeeld(varTeams)



' indien een variable indeling


    Set db = CurrentDb
    sql = ""
    sql = sql & "SELECT tblUitslagen.Wedstrijdnr, tblUitslagen.ImpsThuis, tblUitslagen.ImpsUit, "
    sql = sql & "tblUitslagen.VpsThuis, tblUitslagen.VpsUit, tblTeams.Teamnr AS Thuisnr, tblTeams_1.Teamnr AS Uitnr "
    sql = sql & "FROM (tblTeams INNER JOIN tblUitslagen ON (tblTeams.id = tblUitslagen.TeamIDThuis) AND (tblTeams.id = tblUitslagen.TeamIDThuis)) INNER JOIN tblTeams AS tblTeams_1 ON tblUitslagen.TeamIDUit = tblTeams_1.id "
    sql = sql & "WHERE tblUitslagen.ToernooiID = " & ToernID & " And tblUitslagen.Wedstrijdnr <= " & intRondesGesp & " "
    sql = sql & "ORDER BY tblUitslagen.id, tblUitslagen.Wedstrijdnr;"
    
    Set rs = db.OpenRecordset(sql)
    
    If rs.BOF And rs.EOF Then
        rs.Close
        db.Close
       
        MsgBox ("tab TeamUitslagen heeft geen uitslagen")
        Exit Function
    End If
    rs.MoveFirst
    
    Do While Not rs.EOF
        
  '      Tegenstanders(rs!Thuisnr, rs!Wedstrijdnr) = rs!UitNr
  '      Tegenstanders(rs!UitNr, rs!Wedstrijdnr) = rs!Thuisnr
        If Not IsNull(rs!VpsThuis) Then
                If rs!Wedstrijdnr <= intRondeStand Then
                    Total(rs!Thuisnr) = Total(rs!Thuisnr) + rs!VpsThuis
                    Total(rs!UitNr) = Total(rs!UitNr) + rs!VpsUit
                    WedstrijdenGespeeld(rs!Thuisnr) = WedstrijdenGespeeld(rs!Thuisnr) + 1
                    WedstrijdenGespeeld(rs!UitNr) = WedstrijdenGespeeld(rs!UitNr) + 1
                End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    db.Close
    
    For i = 1 To varTeams
        If WedstrijdenGespeeld(i) > 0 Then
            Gemiddelde(i) = Total(i) / WedstrijdenGespeeld(i)
        End If
    Next
    'Bereken rang
    For i = 1 To varTeams
        Rang(i) = i
    Next
    Call Sorteer(Gemiddelde(), Rang(), True)
    
    For i = 1 To varTeams
        Stand(Rang(i)) = i
    Next
    
    
 '   For teller = 1 To varTeams
 '       For kolomteller = 1 To intRondesGesp
 '            Gespeeld(teller, Tegenstanders(teller, kolomteller)) = True
 ''            Gespeeld(Tegenstanders(teller, kolomteller), teller) = True
 '       Next
 '  Next
    
    

For i = 1 To varTeams \ 2
    Id(i, 0) = Stand((i - 1) * 2 + 1)
    Id(i, 1) = Stand(i * 2)
Next



' start zwitsers
'Call RekenZwitsers(varTeams)

strIndeling = ""
Dim TotalThuis, TotalUit    As String

For i = 1 To varTeams \ 2
    TotalThuis = Format(Total(Id(i, 0)), "#0.00")
    TotalThuis = String(6 - Len(TotalThuis), " ") & TotalThuis
    TotalUit = Format(Total(Id(i, 1)), "#0.00")
    TotalUit = String(6 - Len(TotalUit), " ") & TotalUit
    strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & "  |  Total " & TotalThuis & " - " & TotalUit
    strIndeling = strIndeling & " |  " & Format(Rang(Id(i, 0)), "00") & " - " & Format(Rang(Id(i, 1)), "00") & vbCr
Next


BepaalDeensIntern = Id


End Function

'Indeling 1 - 2 3 - 4


Public Function BepaalIndeling1234(varTeams As Variant, ToernID As Variant, Optional MaxRonde As Variant) As Variant
    ReDim Id(varTeams \ 2, 1)
    For i = 1 To varTeams \ 2
        Id(i, 0) = (i - 1) * 2 + 1
        Id(i, 1) = i * 2
    Next
    Bepaal1234 = Id
End Function

'indeling 1 - 3 4 - 2

Public Function BepaalIndeling1342(varTeams As Variant, ToernID As Variant, Optional MaxRonde As Variant) As Variant
    ReDim Id(varTeams \ 2, 1)
    For i = 1 To varTeams \ 2
    If i Mod 2 = 1 Then
        Id(i, 0) = (i - 1) * 2 + 1
        Id(i, 1) = i * 2 + 1
    Else
         Id(i, 0) = i * 2
         Id(i, 1) = (i - 1) * 2
    End If
        
    Next
           strIndeling = ""
For i = 1 To varTeams \ 2
    strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & vbCr
Next
    'toevoegen aan
 BepaalIndeling1342 = Id

End Function
'indeling 1 - 4 2 - 3


Public Function BepaalIndeling1423(varTeams As Variant, ToernID As Variant, Optional MaxRonde As Variant) As Variant
    ReDim Id(varTeams \ 2, 1)
    For i = 1 To varTeams \ 2
    If i Mod 2 = 1 Then
        Id(i, 0) = (i - 1) * 2 + 1
        Id(i, 1) = i * 2 + 2
    Else
         Id(i, 0) = (i - 1) * 2
         Id(i, 1) = (i - 1) * 2 + 1
    End If
        
    Next
               strIndeling = ""
For i = 1 To varTeams \ 2
    strIndeling = strIndeling & Format(Id(i, 0), "00") & " - " & Format(Id(i, 1), "00") & vbCr
Next
    'toevoegen aan
 
    BepaalIndeling1423 = Id
End Function