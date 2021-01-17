CREATE TABLE [tblTeams] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Teamnr] SHORT ,
  [TeamNaam] VARCHAR (255),
  [Speler1] VARCHAR (255),
  [Speler2] VARCHAR (255),
  [Speler3] VARCHAR (255),
  [Speler4] VARCHAR (255),
  [Speler5] VARCHAR (255),
  [Speler6] VARCHAR (255),
  [Speler7] VARCHAR (255),
  [Speler8] VARCHAR (255),
  [ToernooiID] LONG ,
  [ClubTeamsID] LONG 
)
