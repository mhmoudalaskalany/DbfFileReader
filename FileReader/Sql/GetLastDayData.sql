USE [OcsDb]
GO

/****** Object:  StoredProcedure [dbo].[GetLastDayData]    Script Date: 15/08/2021 01:14:18 ã ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[GetLastDayData] 
	
	@Day Date
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

   SELECT 
   AUFTRAGA.DOSANW FgMaterialCode
  ,AUFTRAGA.BEZEICH ProductVersion
  ,AUFTRAGA.AUFTRAG ProcessOrderNumber
  ,AUFTRAGA.CHARGE  Charge
  ,AUFTRAGA.SOLL    BatchCount
  ,AUFTRAGA.DATUM   EndDate
  ,AUFTRAGA.DATUM  StartDate
  ,AUFTRAGB.ZEIT_VON StartTime
  ,AUFTRAGB.ZEIT_BIS EndTime
  ,AUFTRAGA.FZELLE RawCodeMaterial
  ,AUFTRAGB.ISTKG CurrentQuantity
   
   FROM AUFTRAGA left join  AUFTRAGB

    on AUFTRAGA.AUFTRAG = AUFTRAGB.AUFTRAG

    where AUFTRAGA.DATUM >= @Day
END
GO


