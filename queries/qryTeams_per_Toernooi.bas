Operation =1
Option =0
Begin InputTables
    Name ="tblTeams"
    Name ="tblToernooi"
End
Begin OutputColumns
    Expression ="tblToernooi.ID"
    Expression ="tblTeams.Teamnr"
    Expression ="tblTeams.TeamNaam"
End
Begin Joins
    LeftTable ="tblToernooi"
    RightTable ="tblTeams"
    Expression ="tblToernooi.ID = tblTeams.ToernooiID"
    Flag =1
End
Begin OrderBy
    Expression ="tblToernooi.ID"
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
        dbText "Name" ="tblToernooi.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.TeamNaam"
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
    Bottom =453
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =345
        Top =92
        Right =489
        Bottom =302
        Top =0
        Name ="tblTeams"
        Name =""
    End
    Begin
        Left =109
        Top =83
        Right =253
        Bottom =277
        Top =0
        Name ="tblToernooi"
        Name =""
    End
End
