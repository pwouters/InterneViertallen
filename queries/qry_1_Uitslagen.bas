Operation =1
Option =0
Where ="(((tblUitslagen.ToernooiID)=1))"
Begin InputTables
    Name ="tblUitslagen"
    Name ="tblTeams"
    Name ="tblTeams"
    Alias ="tblTeams_1"
    Name ="tblSessie"
End
Begin OutputColumns
    Expression ="tblSessie.Sessienr"
    Expression ="tblUitslagen.Wedstrijdnr"
    Alias ="Thuisnr"
    Expression ="tblTeams.Teamnr"
    Alias ="Uitnr"
    Expression ="tblTeams_1.Teamnr"
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
    LeftTable ="tblUitslagen"
    RightTable ="tblTeams_1"
    Expression ="tblUitslagen.TeamIDUit = tblTeams_1.id"
    Flag =1
    LeftTable ="tblSessie"
    RightTable ="tblUitslagen"
    Expression ="tblSessie.id = tblUitslagen.SessieID"
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
        dbText "Name" ="tblUitslagen.ToernooiID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.Sessienr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Uitnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.Wedstrijdnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Thuisnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTeams_1.Teamnr"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUitslagen.VpsUit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1006
    Bottom =833
    Left =-1
    Top =-1
    Right =990
    Bottom =571
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =10
        Top =30
        Right =417
        Bottom =429
        Top =0
        Name ="tblUitslagen"
        Name =""
    End
    Begin
        Left =884
        Top =121
        Right =1028
        Bottom =265
        Top =0
        Name ="tblTeams"
        Name =""
    End
    Begin
        Left =882
        Top =315
        Right =1026
        Bottom =459
        Top =0
        Name ="tblTeams_1"
        Name =""
    End
    Begin
        Left =631
        Top =22
        Right =775
        Bottom =166
        Top =0
        Name ="tblSessie"
        Name =""
    End
End
