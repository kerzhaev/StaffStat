CREATE TABLE [tbl_Import_Meta] (
  [ID] AUTOINCREMENT CONSTRAINT [PK_ImportMeta] PRIMARY KEY UNIQUE NOT NULL,
  [ExportFileDate] DATETIME,
  [ImportRunAt] DATETIME,
  [SourceFilePath] VARCHAR (255)
)
