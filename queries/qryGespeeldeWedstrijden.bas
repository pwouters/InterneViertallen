Operation =1
Option =2
Having ="(((Sum(tblUitslagen.VpsThuis))>0))"
Begin InputTables
    Name ="tblSessie"
    Name ="tblUitslagen"
End
Begin OutputColumns
    Expression ="tblUitslagen.ToernooiID"
    Expression ="tblSessie.Sessienr"
    Expression ="tblUitslagen.Wedstrijdnr"
End
Begin Joins
    LeftTable ="tblSessie"
    RightTable ="tblUitslagen"
    Expression ="tblSessie.id = tblUitslagen.SessieID"
    Flag =1
End
Begin Groups
    Expression ="tblUitslagen.ToernooiID"
    GroupLevel =0
    Expression ="tblSessie.Sessienr"
    GroupLevel =0
    Expression ="tblUitslagen.Wedstrijdnr"
    GroupLevel =0
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
        dbText "Name" ="tblUitslagen.ToernooiID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3510"
        dbBoolean "ColumnHidden" ="0"
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
    Right =1645
    Bottom =833
    Left =-1
    Top =-1
    Right =789
    Bottom =554
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =488
        Top =81
        Right =632
        Bottom =225
        Top =0
        Name ="tblSessie"
        Name =""
    End
    Begin
        Left =36
        Top =57
        Right =383
        Bottom =402
        Top =0
        Name ="tblUitslagen"
        Name =""
    End
End
