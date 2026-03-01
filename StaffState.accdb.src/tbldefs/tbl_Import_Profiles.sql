CREATE TABLE [tbl_Import_Profiles] (
  [ProfileID] LONG CONSTRAINT [PK_Import_Profiles] PRIMARY KEY UNIQUE NOT NULL,
  [ProfileName] VARCHAR (100),
  [IdStrategy] VARCHAR (20)
)
