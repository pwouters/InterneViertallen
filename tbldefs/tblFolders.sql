CREATE TABLE [tblFolders] (
  [id] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ExcelFolder] VARCHAR (255),
  [HTMLFolder] VARCHAR (255),
  [TemplateFolder] VARCHAR (255),
  [TemplateFile] VARCHAR (255)
)
