/********************************************\
  Proyect         : Deloitte ExcelAddIn
  Developer       : Ing. Ivan Mu�oz
  Stored Procedure: spObtenerIdTiposPlantillas
\********************************************/
USE [DSAT]
GO

/****** Object:  StoredProcedure [dbo].[spObtenerIdTiposPlantillas]    Script Date: 02/03/2019 06:04:43 p. m. ******/
IF  EXISTS (SELECT * FROM sysobjects WHERE name = 'spObtenerIdTiposPlantillas')
  BEGIN
    DROP PROCEDURE [dbo].[spObtenerIdTiposPlantillas]
  END
GO

/****** Object:  StoredProcedure [dbo].[spObtenerIdTiposPlantillas]    Script Date: 02/03/2019 06:04:43 p. m. ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spObtenerIdTiposPlantillas]
AS
 BEGIN
   SELECT IdTipoPlantilla, MAX(Fecha_Modificacion) AS Fecha_Modificacion FROM tbl_Plantillas WHERE Activo = 1 GROUP BY IdTipoPlantilla ORDER BY IdTipoPlantilla ASC
 END;
GO
