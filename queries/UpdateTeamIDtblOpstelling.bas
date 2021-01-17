Operation =4
Option =0
Where ="(((tblOpstelling.TeamID) Is Null))"
Begin InputTables
    Name ="tblOpstelling"
    Name ="tblTeams"
End
Begin OutputColumns
    Name ="tblOpstelling.TeamID"
    Expression ="[tblTeams]![id]"
End
Begin Joins
    LeftTable ="tblTeams"
    RightTable ="tblOpstelling"
    Expression ="tblTeams.ToernooiID = tblOpstelling.ToernooiID"
    Flag =1
    LeftTable ="tblTeams"
    RightTable ="tblOpstelling"
    Expression ="tblTeams.Teamnr = tblOpstelling.Teamnr"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="tblTeams.id"
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
    Begin
        dbText "Name" ="tblOpstelling.Speler1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOpstelling.Speler2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOpstelling.id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOpstelling.TeamID"
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
    ColumnsShown =579
    Begin
        Left =177
        Top =60
        Right =321
        Bottom =406
        Top =0
        Name ="tblOpstelling"
        Name =""
    End
    Begin
        Left =749
        Top =73
        Right =1035
        Bottom =411
        Top =0
        Name ="tblTeams"
        Name =""
    End
End
