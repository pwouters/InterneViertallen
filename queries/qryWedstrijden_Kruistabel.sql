TRANSFORM Sum(qryWedstrijden.VPS) AS SomVanVPS
SELECT qryWedstrijden.ToernooiID, qryWedstrijden.Teamnr, qryWedstrijden.TeamNaam, Sum(qryWedstrijden.VPS) AS [Totaal VPS], Avg(qryWedstrijden.VPS) AS GemVanVPS
FROM qryWedstrijden
WHERE (((qryWedstrijden.ToernooiID)=lngToernooiID()))
GROUP BY qryWedstrijden.ToernooiID, qryWedstrijden.Teamnr, qryWedstrijden.TeamNaam
PIVOT qryWedstrijden.ZittingNr;
