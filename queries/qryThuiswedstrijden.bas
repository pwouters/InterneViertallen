Operation =1
Option =0
Begin InputTables
    Name ="tblUitslagen"
    Name ="tblTeams"
    Name ="tblSessie"
    Name ="tblTeams"
    Alias ="tblTeams_1"
End
Begin OutputColumns
    Expression ="tblTeams.ToernooiID"
    Expression ="tblTeams.Teamnr"
    Expression ="tblTeams.TeamNaam"
    Expression ="tblSessie.Sessienr"
    Expression ="tblUitslagen.Wedstrijdnr"
    Expression ="tblTeams.id"
    Expression ="tblUitslagen.SessieID"
    Alias ="VPS"
    Expression ="tblUitslagen.VpsThuis"
    Alias ="TegenTeam"
    Expression ="tblUitslagen.TeamIDUit"
    Alias ="Gespeeld"
    Expression ="\"Thuis\""
    Alias ="TegenTeamNr"
    Expression ="tblTeams_1.Teamnr"
    Alias ="TegenTeamNaam"
    Expression ="tblTeams_1.Teamnaam"
    Alias ="ZittingNr"
    Expression ="tblUitslagen.Wedstrijdnr"
End
Begin Joins
    LeftTable ="tblTeams"
    RightTable ="tblUitslagen"
    Expression ="tblTeams.id = tblUitslagen.TeamIDThuis"
    Flag =1
    LeftTable ="tblUitslagen"
    RightTable ="tblTeams_1"
    Expression ="tblUitslagen.TeamIDUit = tblTeams_1.id"
    Flag =1
    LeftTable ="tblSessie"
    RightTable ="tblUitslagen"
    Expression ="tblSessie.id = tblUitslagen.SessieID"
    Flag =1
    LeftTable ="tblSessie"
    RightTable ="tblUitslagen"
    Expression ="tblSessie.id = tblUitslagen.SessieID"
    Flag =1
    LeftTable ="tblSessie"
    RightTable ="tblUitslagen"
    Expression ="tblSessie.id = tblUitslagen.SessieID"
    Flag =1
End
Begin OrderBy
    Expression ="tblTeams.ToernooiID"
    Flag =0
    Expression ="tblTeams.Teamnr"
    Flag =0
    Expression ="tblSessie.Sessienr"
    Flag =0
    Expression ="tblUitslagen.Wedstrijdnr"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblTeams.TeamNaam"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblTeams.id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.SessieID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VPS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.Sessienr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.ToernooiID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.Wedstrijdnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TegenTeam"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Gespeeld"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Teamnr"
        dbInteger "ColumnWidth" ="795"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TegenTeamNr"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TegenTeamNaam"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.VpsThuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.TeamIDUit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams_1.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ZittingNr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =950
    Bottom =715
    Left =-1
    Top =-1
    Right =934
    Bottom =368
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =305
        Top =43
        Right =449
        Bottom =327
        Top =0
        Name ="tblUitslagen"
        Name =""
    End
    Begin
        Left =75
        Top =33
        Right =219
        Bottom =327
        Top =0
        Name ="tblTeams"
        Name =""
    End
    Begin
        Left =526
        Top =104
        Right =670
        Bottom =248
        Top =0
        Name ="tblSessie"
        Name =""
    End
    Begin
        Left =501
        Top =262
        Right =645
        Bottom =406
        Top =0
        Name ="tblTeams_1"
        Name =""
    End
End