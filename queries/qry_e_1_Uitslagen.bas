dbMemo "SQL" ="SELECT tbl_1_Uitslagen.Avond, tbl_1_Uitslagen.Wedstrijd, tbl_1_Uitslagen.Teamnr_"
    "thuis, tbl_1_Uitslagen.Teamnr_Uit, tbl_1_Uitslagen.TeamThuis, tbl_1_Uitslagen.Te"
    "amUit, tbl_1_Uitslagen.ImpsThuis, tbl_1_Uitslagen.ImpsUit, tbl_1_Uitslagen.VPThu"
    "is, tbl_1_Uitslagen.Vpuit, qry_1_Uitslagen.ImpsThuis, qry_1_Uitslagen.ImpsUit, q"
    "ry_1_Uitslagen.VpsThuis, qry_1_Uitslagen.VpsUit\015\012FROM tbl_1_Uitslagen INNE"
    "R JOIN qry_1_Uitslagen ON (tbl_1_Uitslagen.Avond = qry_1_Uitslagen.Sessienr) AND"
    " (tbl_1_Uitslagen.Wedstrijd = qry_1_Uitslagen.Wedstrijdnr) AND (tbl_1_Uitslagen."
    "Teamnr_thuis = qry_1_Uitslagen.Thuisnr) AND (tbl_1_Uitslagen.Teamnr_Uit = qry_1_"
    "Uitslagen.Uitnr);\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="tbl_1_Uitslagen.Avond"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.Wedstrijd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.Teamnr_thuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.ImpsUit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.Teamnr_Uit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.TeamThuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.TeamUit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.ImpsThuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.VPThuis"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_1_Uitslagen.Vpuit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_1_Uitslagen.ImpsThuis"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_1_Uitslagen.ImpsUit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_1_Uitslagen.VpsThuis"
        dbInteger "ColumnWidth" ="1260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_1_Uitslagen.VpsUit"
        dbLong "AggregateType" ="-1"
    End
End
