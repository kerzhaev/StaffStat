CREATE TABLE [tbl_Import_Buffer] (
  [ID] AUTOINCREMENT CONSTRAINT [PK_Buffer] PRIMARY KEY UNIQUE NOT NULL,
  [SourceID_Raw] VARCHAR (255),
  [PersonUID_Raw] VARCHAR (255),
  [Rank_Raw] VARCHAR (255),
  [FullName_Raw] VARCHAR (255),
  [BirthDate_Raw] VARCHAR (255),
  [WorkStatus_Raw] VARCHAR (255),
  [PosCode_Raw] VARCHAR (255),
  [PosName_Raw] LONGTEXT,
  [OrderDate_Raw] VARCHAR (255),
  [OrderNum_Raw] VARCHAR (255)
)
