USE [master]
GO

IF  EXISTS (SELECT name FROM sys.databases WHERE name = N'Api-template-report')
BEGIN
 ALTER DATABASE [Api-template-report] SET SINGLE_USER WITH ROLLBACK IMMEDIATE
 DROP DATABASE [Api-template-report]
END
GO

CREATE DATABASE [Api-template-report]
GO

-- Create the FileLoader login to the server
--IF  EXISTS (SELECT * FROM sys.server_principals WHERE name = N'FileLoader')
--DROP LOGIN [FileLoader]
--GO
 
--EXEC sp_addlogin @loginame = 'FileLoader', @passwd  = 'Pd123456';
--GO
 
USE [Api-template-report]
GO

CREATE TABLE [dbo].[tblFileDetails](  
    [Id] [int] IDENTITY(1,1) NOT NULL,  
    [FILENAME] [nvarchar](100) NULL,  
    [FILEURL] [nvarchar](1500) NULL,
	[JSONNAME] [nvarchar](50) NULL,
	[JSONURL] [nvarchar](1500) NULL
)   

DROP Table [dbo].tblFileDetails

Select * from [dbo].tblFileDetails
Select * from [dbo].tblFileDetails where FILENAME='202104050951397805.Report.docx'
SELECT COUNT(*) FROM [dbo].tblFileDetails where FILENAME= '202104050951397805.Report.docx'
DELETE FROM [dbo].tblFileDetails