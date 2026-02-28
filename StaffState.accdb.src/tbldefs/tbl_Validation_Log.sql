CREATE TABLE [tbl_Validation_Log] (
  [LogID] AUTOINCREMENT CONSTRAINT [PK_ValidationLog] PRIMARY KEY UNIQUE NOT NULL,
  [RecordID] LONG,
  [TableName] VARCHAR (50),
  [ErrorType] VARCHAR (50),
  [ErrorMessage] VARCHAR (255),
  [CheckDate] DATETIME
)
