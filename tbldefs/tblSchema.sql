CREATE TABLE [tblSchema] (
  [Id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ToernooiID] LONG ,
  [Wedstrijdronde] SHORT ,
  [Paring] SHORT ,
  [TeamThuis] SHORT ,
  [TeamUit] SHORT ,
  [Tafel1] VARCHAR (255),
  [Tafel2] VARCHAR (255)
)
