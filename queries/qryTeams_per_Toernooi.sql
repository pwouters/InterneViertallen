SELECT tblToernooi.ID, tblTeams.Teamnr, tblTeams.TeamNaam
FROM tblToernooi INNER JOIN tblTeams ON tblToernooi.ID = tblTeams.ToernooiID
ORDER BY tblToernooi.ID, tblTeams.Teamnr;
