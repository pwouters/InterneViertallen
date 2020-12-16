UPDATE tblTeams INNER JOIN tblOpstelling ON (tblTeams.Teamnr = tblOpstelling.Teamnr) AND (tblTeams.ToernooiID = tblOpstelling.ToernooiID) SET tblOpstelling.TeamID = [tblTeams]![id]
WHERE (((tblOpstelling.TeamID) Is Null));
