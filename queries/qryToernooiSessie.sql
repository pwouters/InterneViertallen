SELECT DISTINCT tblToernooi.ID AS ID_Toernooi, tblToernooi.ToernooiNaam, tblSessie.Sessienr, tblSessie.id AS ID_Sessie, tblSessie.AantalTeams, tblSessie.wedstrijdvormID, tblSessie.AantalWedstrijdenPerSessie
FROM tblToernooi INNER JOIN tblSessie ON tblToernooi.ID = tblSessie.ToernooID
ORDER BY tblToernooi.ID, tblSessie.Sessienr;
