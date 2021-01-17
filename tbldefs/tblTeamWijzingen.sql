CREATE TABLE [tblTeamWijzingen] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [TeamID] SHORT ,
  [Spelnr] SHORT ,
  [Speler1] VARCHAR (255),
  [Speler2] VARCHAR (255),
  [Speler3] VARCHAR (255),
  [Speler4] VARCHAR (255),
  [SessieID] LONG ,
  [ToernooiID] LONG 
)
