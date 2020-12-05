CREATE TABLE [tblGroepen] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SessieID] LONG ,
  [GroepsNaam] VARCHAR (255),
  [GroepsLetter] VARCHAR (5)
)
