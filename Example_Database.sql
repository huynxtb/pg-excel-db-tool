USE [TestImportExcelToDatabase]
GO
/****** Object:  Table [dbo].[ExamList]    Script Date: 5/11/2023 9:39:54 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ExamList](
	[StudentName] [nvarchar](255) NULL,
	[ExamName] [nvarchar](255) NULL,
	[ExamDate] [nvarchar](255) NULL,
	[ExamPoints] [int] NULL
) ON [PRIMARY]
GO
