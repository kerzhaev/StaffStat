CREATE TABLE [tbl_Import_Mapping] (
  [MappingID] AUTOINCREMENT CONSTRAINT [PK_Import_Mapping] PRIMARY KEY UNIQUE NOT NULL,
  [ProfileID] LONG,
  [ExcelHeader] VARCHAR (255),
  [TargetField] VARCHAR (100),
  [Поле1] VARCHAR (255)
)
