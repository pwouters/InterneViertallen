CREATE TABLE [tblSessie] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ToernooID] LONG ,
  [Sessienaam] VARCHAR (255),
  [Sessienr] SHORT ,
  [Aantalspellen] SHORT ,
  [Competitie] SHORT ,
  [Prefixkopjesscorestaat] VARCHAR (255),
  [PrefixKopjeuitslagen] VARCHAR (255),
  [Suffixkopjesscorestaat] VARCHAR (255),
  [SuffixKopjeuitslagen] VARCHAR (255),
  [Voettekst] VARCHAR (255),
  [Voetlink] VARCHAR (255),
  [wedstrijdvormID] SHORT ,
  [ActivityID] SHORT ,
  [AantalTeams] SHORT ,
  [ByeTeam] BIT ,
  [AantalWedstrijdenPerSessie] SHORT 
)