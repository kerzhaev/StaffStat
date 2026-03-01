CREATE TABLE [tbl_Localization] (
  [MsgKey] VARCHAR (100) CONSTRAINT [PK_Localization] PRIMARY KEY UNIQUE NOT NULL,
  [LocalValue] LONGTEXT,
  [Category] VARCHAR (50)
)
