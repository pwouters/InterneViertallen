SELECT tblToernooi.ID, tblSessie.Sessienr, tblUitslagen.Wedstrijdnr
FROM (tblToernooi INNER JOIN tblUitslagen ON tblToernooi.ID = tblUitslagen.ToernooiID) INNER JOIN tblSessie ON (tblSessie.id = tblUitslagen.SessieID) AND (tblToernooi.ID = tblSessie.ToernooID) AND (tblUitslagen.SessieID = tblSessie.id)
ORDER BY tblToernooi.ID, tblSessie.Sessienr, tblUitslagen.Wedstrijdnr;
