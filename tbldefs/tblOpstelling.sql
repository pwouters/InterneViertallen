CREATE TABLE [tblOpstelling] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ToernooiID] LONG ,
  [Sessie] UNSIGNED BYTE ,
  [Teamnr] SHORT ,
  [Speler1] VARCHAR (255),
  [Speler2] VARCHAR (255),
  [Speler3] VARCHAR (255),
  [Speler4] VARCHAR (255),
  [Wedstrijd1] SHORT ,
  [Wedstrijd2] SHORT ,
  [Wedstrijd3] SHORT ,
  [Wedstrijd4] SHORT ,
  [Wedstrijd5] SHORT ,
  [Wedstrijd6] SHORT ,
  [Wedstrijd7] SHORT ,
  [Wedstrijd8] SHORT ,
  [Wedstrijd9] SHORT ,
  [Wedstrijd10] SHORT ,
  [Wedstrijd11] SHORT 
)
