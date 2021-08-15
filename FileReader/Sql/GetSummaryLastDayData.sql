USE [OcsDb]
GO

/****** Object:  StoredProcedure [dbo].[GetSummaryLastDayData]    Script Date: 15/08/2021 02:18:55 ã ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Mahmoud Alaskalany>
-- Create date: <15/8/2021>
-- Description:	<Get Summary Data Of Last Day>
-- =============================================
CREATE PROCEDURE [dbo].[GetSummaryLastDayData] 
	-- Add the parameters for the stored procedure here
	@Day Date
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

   select 
   AUFTRAGC.AUFTRAG  ProcessOrderNumber ,
   AUFTRAGC.CHARGENR  Chargener,
   AUFTRAGC.ROHSTOFF  RawMaterialCode,
   AUFTRAGC.SOLLKG  StartQty,
   AUFTRAGC.ISTKG  CurrentQty,
   AUFTRAGC.BEMERKUNG Comment,
   AUFTRAGC.DATUM  Date,
   AUFTRAGC.ZEIT_VON StartTime ,
   AUFTRAGC.ZEIT_BIS EndTime
   from AUFTRAGC   where DATUM >= @Day
END
GO


