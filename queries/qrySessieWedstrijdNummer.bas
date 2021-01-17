dbMemo "SQL" ="SELECT DISTINCT tblSessie.ToernooID, tblSessie.id, tblSessie.Sessienr, tblUitsla"
    "gen.Wedstrijdnr\015\012FROM tblSessie INNER JOIN tblUitslagen ON tblSessie.id = "
    "tblUitslagen.SessieID\015\012ORDER BY tblSessie.ToernooID, tblSessie.Sessienr;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblSessie.ToernooID"
        dbInteger "ColumnWidth" ="4590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.id"
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
