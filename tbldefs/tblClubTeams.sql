CREATE TABLE [tblClubTeams] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ClubTeam] VARCHAR (255),
  [Beker] VARCHAR (255),
  [Clubnr] VARCHAR (255),
  [ToernooiID] LONG 
)