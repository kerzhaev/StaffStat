CREATE TABLE [tbl_Settings] (
  [SettingKey] VARCHAR (50) CONSTRAINT [PK_Settings] PRIMARY KEY UNIQUE NOT NULL,
  [SettingValue] VARCHAR (255),
  [SettingGroup] VARCHAR (50),
  [Description] VARCHAR (255)
)
