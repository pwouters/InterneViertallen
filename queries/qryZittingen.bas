Operation =1
Option =0
Begin InputTables
    Name ="tblToernooi"
    Name ="tblUitslagen"
    Name ="tblSessie"
End
Begin OutputColumns
    Expression ="tblToernooi.ID"
    Expression ="tblSessie.Sessienr"
    Expression ="tblUitslagen.Wedstrijdnr"
End
Begin Joins
    LeftTable ="tblToernooi"
    RightTable ="tblUitslagen"
    Expression ="tblToernooi.ID = tblUitslagen.ToernooiID"
    Flag =1
    LeftTable ="tblSessie"
    RightTable ="tblUitslagen"
    Expression ="tblSessie.id = tblUitslagen.SessieID"
    Flag =1
    LeftTable ="tblToernooi"
    RightTable ="tblSessie"
    Expression ="tblToernooi.ID = tblSessie.ToernooID"
    Flag =1
    LeftTable ="tblUitslagen"
    RightTable ="tblSessie"
    Expression ="tblUitslagen.SessieID = tblSessie.id"
    Flag =1
End
Begin OrderBy
    Expression ="tblToernooi.ID"
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
        dbText "Name" ="tblToernooi.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.Sessienr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.Wedstrijdnr"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1076
    Bottom =715
    Left =-1
    Top =-1
    Right =1060
    Bottom =470
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblToernooi"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblUitslagen"
        Name =""
    End
    Begin
        Left =607
        Top =55
        Right =751
        Bottom =199
        Top =0
        Name ="tblSessie"
        Name =""
    End
End
