SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_clm_num_lu_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_clm_num_lu_select]
GO


CREATE PROCEDURE proc_clm_num_lu_select (@clm_id dom_id,
					@clm_num VARCHAR(20) OUTPUT) WITH RECOMPILE AS
BEGIN
 /* Date: 05/07/2003
  * Author: Betsy Walker
  * 
  * Purpose: This stored procedure is written with the intent to select a Claim Number from the CLAIM_Ttable 
  *          based on a specified Claim ID. 
  * 
  * Logic flow: Perform the select of the Claim Number from the CLAIM_T table
  *             Perform error checking
  *
  * Return: number of rows for success
  *             4028 for failure
  *
  *  Change Log
  *  08/21/07 K758  Changed to accommodate expansion of Claim# to varchar(20) per WRM6371.
  */
	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	IF NOT EXISTS(SELECT *
			FROM claim_t
			WHERE clm_id = @clm_id)
	BEGIN                                
		SELECT @Error_Message = 'Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>') + ' does not exist in the CLAIM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	SELECT @clm_num = clm_num
	FROM claim_t 
	WHERE clm_id = @clm_id

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error happened while selecting the Claim from the CLAIM_T table for the Claim ID: ' + CONVERT(VARCHAR(255), @clm_id)
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON [dbo].proc_clm_num_lu_select TO AppRoleClaims
GO
GRANT EXECUTE ON [dbo].proc_clm_num_lu_select TO Support
GO
