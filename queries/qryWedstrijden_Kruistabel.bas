Operation =6
Option =0
Where ="(((qryWedstrijden.ToernooiID)=lngToernooiID()))"
Begin InputTables
    Name ="qryWedstrijden"
End
Begin OutputColumns
    Expression ="qryWedstrijden.ToernooiID"
    GroupLevel =2
    Expression ="qryWedstrijden.Teamnr"
    GroupLevel =2
    Expression ="qryWedstrijden.TeamNaam"
    GroupLevel =2
    Expression ="qryWedstrijden.ZittingNr"
    GroupLevel =1
    Alias ="SomVanVPS"
    Expression ="Sum(qryWedstrijden.VPS)"
    Alias ="Totaal VPS"
    Expression ="Sum(qryWedstrijden.VPS)"
    GroupLevel =2
    Alias ="GemVanVPS"
    Expression ="Avg(qryWedstrijden.VPS)"
    GroupLevel =2
End
Begin Groups
    Expression ="qryWedstrijden.ToernooiID"
    GroupLevel =2
    Expression ="qryWedstrijden.Teamnr"
    GroupLevel =2
    Expression ="qryWedstrijden.TeamNaam"
    GroupLevel =2
    Expression ="qryWedstrijden.ZittingNr"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Teamnr]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[TeamNaam]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Totaal VPS"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="#,#00"
        dbByte "DecimalPlaces" ="2"
    End
    Begin
        dbText "Name" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SomVanVPS"
    End
    Begin
        dbText "Name" ="162"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="63"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1510"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1010"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1410"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="123"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="101"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1210"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="93"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="112"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryWedstrijden.[Teamnr]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="22"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryWedstrijden.ToernooiID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="710"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="163"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryWedstrijden.[TeamNaam]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="42"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="143"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="110"
        dbInteger "ColumnWidth" ="1125"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wedstrijd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="102"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1310"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="11"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="12"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="13"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="21"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="210"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="810"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="31"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="41"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="32"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="72"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="103"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1110"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="113"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="121"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="131"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="132"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="141"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="43"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="151"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="52"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="152"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="61"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="153"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="610"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="161"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="62"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1610"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="510"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="73"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="81"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="82"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="83"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="91"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="#,#00"
        dbByte "DecimalPlaces" ="2"
    End
    Begin
        dbText "Name" ="qryWedstrijden.Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryWedstrijden.ZittingNr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryWedstrijden.TeamNaam"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryWedstrijden.VPS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GemVanVPS"
        dbInteger "ColumnWidth" ="1320"
        dbBoolean "ColumnHidden" ="0"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="#,#00"
    End
    Begin
        dbText "Name" ="Gem"
        dbInteger "ColumnWidth" ="3150"
        dbBoolean "ColumnHidden" ="0"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gem1"
        dbInteger "ColumnWidth" ="3150"
        dbBoolean "ColumnHidden" ="0"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ToernooiID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Teamnr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TeamNaam"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =950
    Bottom =715
    Left =-1
    Top =-1
    Right =934
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qryWedstrijden"
        Name =""
    End
End
