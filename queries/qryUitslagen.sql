SELECT tblUitslagen.SessieID, tblTeams.Teamnr AS TeamNrThuis, tblTeams_1.Teamnr AS TeamNrUit, tblTeams.TeamNaam AS Thuis, tblTeams_1.TeamNaam AS Uit, tblUitslagen.ImpsThuis, tblUitslagen.ImpsUit, tblUitslagen.VpsThuis, tblUitslagen.VpsUit
FROM (tblTeams INNER JOIN tblUitslagen ON (tblTeams.id = tblUitslagen.TeamIDThuis) AND (tblTeams.id = tblUitslagen.TeamIDThuis)) INNER JOIN tblTeams AS tblTeams_1 ON tblUitslagen.TeamIDUit = tblTeams_1.id;
