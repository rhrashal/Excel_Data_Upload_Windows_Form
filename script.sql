USE [testDB]
GO
/****** Object:  Table [dbo].[Employee]    Script Date: 6/3/2022 1:00:10 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Employee](
	[Id] [int] IDENTITY NOT NULL primary key,
	[EmpName] [nvarchar](max) NULL,
	[Phone] [nvarchar](50) NULL,
	[Email] [nvarchar](50) NULL,
	[Address] [nvarchar](max) NULL
 ) 
GO
