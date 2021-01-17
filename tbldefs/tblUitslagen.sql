CREATE TABLE [tblUitslagen] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SessieID] LONG ,
  [TeamIDThuis] SHORT ,
  [TeamIDUit] SHORT ,
  [Wedstrijdnr] LONG ,
  [ImpsThuis] SHORT ,
  [ImpsUit] SHORT ,
  [VpsThuis] DOUBLE ,
  [VpsUit] DOUBLE ,
  [ToernooiID] LONG ,
  [Tafel] VARCHAR (255)
)
