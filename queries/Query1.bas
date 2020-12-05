dbMemo "SQL" ="SELECT tblSessie.id, tblSessie.Sessienaam, tblSessie.Sessienr\015\012FROM tblSes"
    "sie\015\012WHERE (((tblSessie.ToernooID) =1));\015\012"
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
        dbText "Name" ="tblSessie.id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.Sessienaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSessie.Sessienr"
        dbLong "AggregateType" ="-1"
    End
End
