Operation =1
Option =2
Begin InputTables
    Name ="tblToernooi"
    Name ="tblSessie"
End
Begin OutputColumns
    Alias ="ID_Toernooi"
    Expression ="tblToernooi.ID"
    Expression ="tblToernooi.ToernooiNaam"
    Expression ="tblSessie.Sessienr"
    Alias ="ID_Sessie"
    Expression ="tblSessie.id"
    Expression ="tblSessie.AantalTeams"
    Expression ="tblSessie.wedstrijdvormID"
    Expression ="tblSessie.AantalWedstrijdenPerSessie"
End
Begin Joins
    LeftTable ="tblToernooi"
    RightTable ="tblSessie"
    Expression ="tblToernooi.ID = tblSessie.ToernooID"
    Flag =1
End
Begin OrderBy
    Expression ="tblToernooi.ID"
    Flag =0
    Expression ="tblSessie.Sessienr"
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
        dbText "Name" ="tblToernooi.ToernooiNaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.Sessienr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.AantalTeams"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.wedstrijdvormID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID_Toernooi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID_Sessie"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.AantalWedstrijdenPerSessie"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1038
    Bottom =833
    Left =-1
    Top =-1
    Right =1022
    Bottom =486
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =68
        Top =108
        Right =212
        Bottom =252
        Top =0
        Name ="tblToernooi"
        Name =""
    End
    Begin
        Left =299
        Top =58
        Right =553
        Bottom =430
        Top =0
        Name ="tblSessie"
        Name =""
    End
End
