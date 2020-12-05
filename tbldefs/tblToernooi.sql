CREATE TABLE [tblToernooi] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ToernooiNaam] VARCHAR (255),
  [WORKFOLDER] LONGTEXT ,
  [WORKFILE] VARCHAR (255),
  [STEPDATA] LONGTEXT ,
  [STEPRESULTS] LONGTEXT ,
  [LOCALSITE] VARCHAR (255),
  [LOCALHTML] VARCHAR (255),
  [PREFIX] VARCHAR (255),
  [AANTALSESSIES] SHORT ,
  [WEDSTRIJDENPERSESSIE] SHORT ,
  [WORKTEMPLATE] VARCHAR (255)
)
