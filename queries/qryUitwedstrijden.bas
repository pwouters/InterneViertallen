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
    Expression ="tblTeams.id"
    Expression ="tblUitslagen.SessieID"
    Expression ="tblSessie.Sessienr"
    Expression ="tblUitslagen.Wedstrijdnr"
    Alias ="VPS"
    Expression ="tblUitslagen.VpsUit"
    Alias ="TegenTeam"
    Expression ="tblUitslagen.TeamIDThuis"
    Alias ="Gespeeld"
    Expression ="\"Uit\""
    Alias ="TegenTeamNr"
    Expression ="tblTeams_1.Teamnr"
    Alias ="TegenTeamNaam"
    Expression ="tblTeams_1.TeamNaam"
    Alias ="ZittingNr"
    Expression ="tblUitslagen.Wedstrijdnr"
End
Begin Joins
    LeftTable ="tblTeams"
    RightTable ="tblUitslagen"
    Expression ="tblTeams.id = tblUitslagen.TeamIDUit"
    Flag =1
    LeftTable ="tblUitslagen"
    RightTable ="tblTeams_1"
    Expression ="tblUitslagen.TeamIDThuis = tblTeams_1.id"
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
    Expression ="tblTeams.id"
    Flag =0
    Expression ="tblUitslagen.SessieID"
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
        dbText "Name" ="tblTeams.id"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="tblUitslagen.SessieID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="VPS"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="tblSessie.Sessienr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="tblTeams.ToernooiID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="tblUitslagen.Wedstrijdnr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="TegenTeam"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Gespeeld"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Teamnr"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.TeamNaam"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TegenTeamNr"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TegenTeamNaam"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
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
    Right =1075
    Bottom =833
    Left =-1
    Top =-1
    Right =1059
    Bottom =351
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
        Left =521
        Top =279
        Right =665
        Bottom =423
        Top =0
        Name ="tblTeams_1"
        Name =""
    End
End
