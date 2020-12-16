dbMemo "SQL" ="Select * from qryThuiswedstrijden\015\012UNION Select * from qryUitwedstrijden;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qryThuiswedstrijden.tblTeams.id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.tblTeams.ToernooiID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.tblUitslagen.SessieID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.VPS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.tblSessie.Sessienr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.TegenTeamNr"
        dbInteger "ColumnWidth" ="2730"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.tblUitslagen.Wedstrijdnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.tblTeams.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.tblTeams.TeamNaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.TegenTeam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.Gespeeld"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryThuiswedstrijden.TegenTeamNaam"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ToernooiID"
    End
    Begin
        dbText "Name" ="Teamnr"
    End
    Begin
        dbText "Name" ="TeamNaam"
    End
    Begin
        dbText "Name" ="Sessienr"
    End
    Begin
        dbText "Name" ="Wedstrijdnr"
    End
    Begin
        dbText "Name" ="id"
    End
    Begin
        dbText "Name" ="SessieID"
    End
    Begin
        dbText "Name" ="VPS"
    End
    Begin
        dbText "Name" ="TegenTeam"
    End
    Begin
        dbText "Name" ="Gespeeld"
    End
    Begin
        dbText "Name" ="TegenTeamNr"
    End
    Begin
        dbText "Name" ="TegenTeamNaam"
    End
    Begin
        dbText "Name" ="ZittingNr"
    End
End
