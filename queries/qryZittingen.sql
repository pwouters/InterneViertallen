SELECT tblToernooi.ID, tblSessie.Sessienr, tblUitslagen.Wedstrijdnr
FROM tblSessie INNER JOIN (tblToernooi INNER JOIN tblUitslagen ON tblToernooi.ID = tblUitslagen.ToernooiID) ON (tblUitslagen.SessieID = tblSessie.id) AND (tblToernooi.ID = tblSessie.ToernooID) AND (tblSessie.id = tblUitslagen.SessieID)
ORDER BY tblToernooi.ID, tblSessie.Sessienr, tblUitslagen.Wedstrijdnr;
