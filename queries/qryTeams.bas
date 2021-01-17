Operation =1
Option =1
Begin InputTables
    Name ="tblTeams"
End
Begin OutputColumns
End
Begin OrderBy
    Expression ="tblTeams.ToernooiID"
    Flag =0
    Expression ="tblTeams.Teamnr"
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
        dbText "Name" ="tblTeams.ToernooiID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.ClubTeamsID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.TeamNaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Speler5"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1748
    Bottom =929
    Left =-1
    Top =-1
    Right =1732
    Bottom =453
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblTeams"
        Name =""
    End
End
