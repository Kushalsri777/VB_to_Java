SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_clm_id_lu_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_clm_id_lu_select]
GO


/* Date: 01/02/2003
  * Author: Betsy Walker
  * 
  * Purpose: This stored procedure is written with the intent to select a Claim ID from the CLAIM_Ttable 
  *                based on a specified Claim Number. 
  * 
  * Logic flow: Perform the select of the Claim ID from the CLAIM_T table
  *                   Perform error checking
  *
  * Return: number of rows for success
  *             4028 for failure
  *
  *  Change Log
  *  08/21/07 K758  Changed to accommodate expansion of Claim# to varchar(20) per WRM6371.
  */

CREATE PROCEDURE proc_clm_id_lu_select (@clm_num VARCHAR(20),
					@clm_id dom_id OUTPUT) WITH RECOMPILE AS

BEGIN

	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	IF NOT EXISTS(SELECT *
			FROM claim_t
			WHERE clm_num = @clm_num)
	BEGIN
		SELECT @Error_Message = 'Claim Number: ' + ISNULL(CONVERT(VARCHAR(255), @clm_num), '<NULL>') + ' does not exist in the CLAIM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	SELECT @clm_id = clm_id
	FROM claim_t 
	WHERE clm_num = @clm_num

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error happened while selecting the Claim from the CLAIM_T table for the Claim Number: ' + CONVERT(VARCHAR(255), @clm_num)
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON [dbo].proc_clm_id_lu_select TO AppRoleClaims
GO
GRANT EXECUTE ON [dbo].proc_clm_id_lu_select TO Support
GO
