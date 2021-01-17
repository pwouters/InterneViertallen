SELECT tblOpstelling.ToernooiID, tblOpstelling.Teamnr, tblTeams.TeamNaam, tblOpstelling.SessieID
FROM tblOpstelling INNER JOIN tblTeams ON tblOpstelling.TeamID = tblTeams.id;
