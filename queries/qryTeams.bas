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
End
Begin
    State =0
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =-1
    Top =-1
    Right =1330
    Bottom =470
    Left =0
    Top =0
    ColumnsShown =521
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
