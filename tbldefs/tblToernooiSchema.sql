CREATE TABLE [tblToernooiSchema] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ToernooiID] LONG ,
  [SessieID] VARCHAR (255),
  [Wedstrijdnr] SHORT ,
  [ThuisTeamnr] SHORT ,
  [UitTeamnr] SHORT 
)
