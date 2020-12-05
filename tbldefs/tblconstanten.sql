CREATE TABLE [tblconstanten] (
  [Id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [iToernooi] VARCHAR (255),
  [iSoort] VARCHAR (255),
  [iNaam] LONGTEXT ,
  [iFunctie] VARCHAR (255),
  [iToernooiID] LONG 
)
