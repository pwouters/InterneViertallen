CREATE TABLE [tblIndeling] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SessieID] LONG ,
  [GroepsID] LONG ,
  [TeamThuis] SHORT ,
  [TeamUit] SHORT ,
  [Tafel1] VARCHAR (255),
  [Tafel2] VARCHAR (255),
  [Wedstrijdnr] LONG 
)
