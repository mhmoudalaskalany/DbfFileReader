USE [OcsDb]
GO

/****** Object:  StoredProcedure [dbo].[GetLfsLastDayData]    Script Date: 11/3/2021 12:11:47 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Mahmoud Alaskalany>
-- Create date: <15/8/2021>
-- Description:	<Get Summary Data Of Last Day>
-- =============================================
CREATE PROCEDURE [dbo].[GetLfsLastDayData] 
	-- Add the parameters for the stored procedure here
	@Day Date
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

   select 
   LNUMMER SrNo,
   WNUMMER MaterialCode,
   ARTNAME MaterialDescription,
   '' DeliveryType,
   '' DeliveryNo,
   INFO ReferenceDocNo,
   BEMERKUNG ContainerNo,
   '' TransporterName,
   FAHRZEUG  VehicleNo,
   DATUM InDate,
   ZEIT InTime,
   DATUM2 OutDate,
   ZEIT2 OutTime,
   TARA TareWeight,
   BRUTTO GrossWeight,
   NETTO NetWeight,
   WAEGER	Weigher
   from LFS   where DATUM >= @Day
END
GO


