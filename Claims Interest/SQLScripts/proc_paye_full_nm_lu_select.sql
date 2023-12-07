SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_paye_full_nm_lu_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_paye_full_nm_lu_select]
GO



CREATE PROCEDURE [dbo].proc_paye_full_nm_lu_select (@paye_id dom_id,
					@paye_full_nm VARCHAR(40) OUTPUT) WITH RECOMPILE AS
BEGIN
 /* Date: 05/07/2003
  * Author: Betsy Walker
  * 
  * Purpose: This stored procedure is written with the intent to select a Payee Full Name 
  *          from the PAYEE_T table based on a specified Payee ID. 
  * 
  * Logic flow: Perform the select of the Payee Full Name from the PAYEE_T table
  *             Perform error checking
  *
  * Return: number of rows for success
  *             4028 for failure
  */
	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	IF NOT EXISTS(SELECT *
			FROM payee_t
			WHERE paye_id = @paye_id)
	BEGIN                                
		SELECT @Error_Message = 'Payee ID: ' + ISNULL(CONVERT(VARCHAR(255), @paye_id), '<NULL>') + ' does not exist in the PAYEE_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	SELECT @paye_full_nm = paye_full_nm
	FROM payee_t 
	WHERE paye_id = @paye_id

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error happened while selecting the Payee from the PAYEE_T table for the Payee ID: ' + CONVERT(VARCHAR(255), @paye_id)
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT  EXECUTE  ON [dbo].[proc_paye_full_nm_lu_select]  TO [Support]
GO
GRANT  EXECUTE  ON [dbo].[proc_paye_full_nm_lu_select]  TO [AppRoleClaims]
GO

