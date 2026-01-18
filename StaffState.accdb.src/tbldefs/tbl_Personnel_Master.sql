CREATE TABLE [tbl_Personnel_Master] (
  [PersonUID] VARCHAR (50) CONSTRAINT [PK_Person] PRIMARY KEY UNIQUE NOT NULL,
  [SourceID] LONG,
  [FullName] VARCHAR (150),
  [RankName] VARCHAR (100),
  [BirthDate] DATETIME,
  [WorkStatus] VARCHAR (100),
  [PosCode] VARCHAR (50),
  [PosName] LONGTEXT,
  [OrderDate] DATETIME,
  [OrderNum] VARCHAR (50),
  [LastUpdated] DATETIME,
  [IsActive] BIT
)
