SELECT DISTINCT tblSessie.ToernooID, tblSessie.id, tblSessie.Sessienr, tblUitslagen.Wedstrijdnr
FROM tblSessie INNER JOIN tblUitslagen ON tblSessie.id = tblUitslagen.SessieID
ORDER BY tblSessie.ToernooID, tblSessie.Sessienr;
