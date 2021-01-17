SELECT DISTINCT tblUitslagen.ToernooiID, tblSessie.Sessienr, tblUitslagen.Wedstrijdnr
FROM tblSessie INNER JOIN tblUitslagen ON tblSessie.id = tblUitslagen.SessieID
GROUP BY tblUitslagen.ToernooiID, tblSessie.Sessienr, tblUitslagen.Wedstrijdnr
HAVING (((Sum(tblUitslagen.VpsThuis))>0));
