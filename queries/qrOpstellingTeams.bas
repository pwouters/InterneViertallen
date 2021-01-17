Operation =1
Option =0
Begin InputTables
    Name ="tblOpstelling"
    Name ="tblTeams"
End
Begin OutputColumns
    Expression ="tblOpstelling.ToernooiID"
    Expression ="tblOpstelling.Teamnr"
    Expression ="tblTeams.TeamNaam"
    Expression ="tblOpstelling.SessieID"
End
Begin Joins
    LeftTable ="tblOpstelling"
    RightTable ="tblTeams"
    Expression ="tblOpstelling.TeamID = tblTeams.id"
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
        dbText "Name" ="tblOpstelling.ToernooiID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOpstelling.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.TeamNaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOpstelling.SessieID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =908
    Bottom =929
    Left =-1
    Top =-1
    Right =892
    Bottom =684
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =154
        Top =50
        Right =376
        Bottom =475
        Top =0
        Name ="tblOpstelling"
        Name =""
    End
    Begin
        Left =550
        Top =71
        Right =694
        Bottom =215
        Top =0
        Name ="tblTeams"
        Name =""
    End
End
