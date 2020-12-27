SELECT tblSessie.Sessienr, tblUitslagen.Wedstrijdnr, tblTeams.Teamnr AS Thuisnr, tblTeams_1.Teamnr AS Uitnr, tblUitslagen.ImpsThuis, tblUitslagen.ImpsUit, tblUitslagen.VpsThuis, tblUitslagen.VpsUit
FROM tblSessie INNER JOIN ((tblTeams INNER JOIN tblUitslagen ON tblTeams.id = tblUitslagen.TeamIDThuis) INNER JOIN tblTeams AS tblTeams_1 ON tblUitslagen.TeamIDUit = tblTeams_1.id) ON tblSessie.id = tblUitslagen.SessieID
WHERE (((tblUitslagen.ToernooiID)=1));
