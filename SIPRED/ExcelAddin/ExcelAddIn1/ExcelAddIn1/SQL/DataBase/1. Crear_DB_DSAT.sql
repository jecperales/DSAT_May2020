/********************************\
  Proyect  : Deloitte ExcelAddIn
  Developer: Ing. Ivan Mu�oz
  Action   : Create DB DSAT
\********************************/
USE [master]
GO
/****** Object:  Database [DSAT] ******/
IF EXISTS(SELECT * FROM DBO.SYSDATABASES WHERE NAME = 'DSAT')
  BEGIN
    DROP DATABASE [DSAT]
  END
GO
/****** Object:  Database [DSAT] ******/
CREATE DATABASE [DSAT]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'DSAT', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\DATA\DSAT.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'DSAT_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\DATA\DSAT_log.ldf' , SIZE = 73728KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
GO

ALTER DATABASE [DSAT] SET COMPATIBILITY_LEVEL = 130
GO

IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [DSAT].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO

ALTER DATABASE [DSAT] SET ANSI_NULL_DEFAULT OFF 
GO

ALTER DATABASE [DSAT] SET ANSI_NULLS OFF 
GO

ALTER DATABASE [DSAT] SET ANSI_PADDING OFF 
GO

ALTER DATABASE [DSAT] SET ANSI_WARNINGS OFF 
GO

ALTER DATABASE [DSAT] SET ARITHABORT OFF 
GO

ALTER DATABASE [DSAT] SET AUTO_CLOSE OFF 
GO

ALTER DATABASE [DSAT] SET AUTO_SHRINK OFF 
GO

ALTER DATABASE [DSAT] SET AUTO_UPDATE_STATISTICS ON 
GO

ALTER DATABASE [DSAT] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO

ALTER DATABASE [DSAT] SET CURSOR_DEFAULT  GLOBAL 
GO

ALTER DATABASE [DSAT] SET CONCAT_NULL_YIELDS_NULL OFF 
GO

ALTER DATABASE [DSAT] SET NUMERIC_ROUNDABORT OFF 
GO

ALTER DATABASE [DSAT] SET QUOTED_IDENTIFIER OFF 
GO

ALTER DATABASE [DSAT] SET RECURSIVE_TRIGGERS OFF 
GO

ALTER DATABASE [DSAT] SET  DISABLE_BROKER 
GO

ALTER DATABASE [DSAT] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO

ALTER DATABASE [DSAT] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO

ALTER DATABASE [DSAT] SET TRUSTWORTHY OFF 
GO

ALTER DATABASE [DSAT] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO

ALTER DATABASE [DSAT] SET PARAMETERIZATION SIMPLE 
GO

ALTER DATABASE [DSAT] SET READ_COMMITTED_SNAPSHOT OFF 
GO

ALTER DATABASE [DSAT] SET HONOR_BROKER_PRIORITY OFF 
GO

ALTER DATABASE [DSAT] SET RECOVERY FULL 
GO

ALTER DATABASE [DSAT] SET  MULTI_USER 
GO

ALTER DATABASE [DSAT] SET PAGE_VERIFY CHECKSUM  
GO

ALTER DATABASE [DSAT] SET DB_CHAINING OFF 
GO

ALTER DATABASE [DSAT] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO

ALTER DATABASE [DSAT] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO

ALTER DATABASE [DSAT] SET DELAYED_DURABILITY = DISABLED 
GO

ALTER DATABASE [DSAT] SET QUERY_STORE = OFF
GO

USE [DSAT]
GO

ALTER DATABASE SCOPED CONFIGURATION SET LEGACY_CARDINALITY_ESTIMATION = OFF;
GO

ALTER DATABASE SCOPED CONFIGURATION SET MAXDOP = 0;
GO

ALTER DATABASE SCOPED CONFIGURATION SET PARAMETER_SNIFFING = ON;
GO

ALTER DATABASE SCOPED CONFIGURATION SET QUERY_OPTIMIZER_HOTFIXES = OFF;
GO

ALTER DATABASE [DSAT] SET  READ_WRITE 
GO