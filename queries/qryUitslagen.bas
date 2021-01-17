Operation =1
Option =0
Begin InputTables
    Name ="tblUitslagen"
    Name ="tblTeams"
    Name ="tblTeams"
    Alias ="tblTeams_1"
End
Begin OutputColumns
    Expression ="tblUitslagen.SessieID"
    Alias ="TeamNrThuis"
    Expression ="tblTeams.Teamnr"
    Alias ="TeamNrUit"
    Expression ="tblTeams_1.Teamnr"
    Alias ="Thuis"
    Expression ="tblTeams.TeamNaam"
    Alias ="Uit"
    Expression ="tblTeams_1.TeamNaam"
    Expression ="tblUitslagen.ImpsThuis"
    Expression ="tblUitslagen.ImpsUit"
    Expression ="tblUitslagen.VpsThuis"
    Expression ="tblUitslagen.VpsUit"
End
Begin Joins
    LeftTable ="tblTeams"
    RightTable ="tblUitslagen"
    Expression ="tblTeams.id = tblUitslagen.TeamIDThuis"
    Flag =1
    LeftTable ="tblTeams"
    RightTable ="tblUitslagen"
    Expression ="tblTeams.id = tblUitslagen.TeamIDThuis"
    Flag =1
    LeftTable ="tblUitslagen"
    RightTable ="tblTeams_1"
    Expression ="tblUitslagen.TeamIDUit = tblTeams_1.id"
    Flag =1
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
        dbText "Name" ="TeamNrUit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.SessieID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.TeamNaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Thuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TeamNrThuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams_1.TeamNaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams_1.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Uit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.ImpsThuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.ImpsUit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.VpsThuis"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.VpsUit"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =955
    Bottom =796
    Left =-1
    Top =-1
    Right =939
    Bottom =551
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =182
        Top =147
        Right =421
        Bottom =442
        Top =0
        Name ="tblUitslagen"
        Name =""
    End
    Begin
        Left =721
        Top =93
        Right =865
        Bottom =237
        Top =0
        Name ="tblTeams"
        Name =""
    End
    Begin
        Left =699
        Top =324
        Right =843
        Bottom =468
        Top =0
        Name ="tblTeams_1"
        Name =""
    End
End
