SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_verify_dependents]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_verify_dependents]
GO


/* Date: 10/14/2002
  * AUTHOR: Betsy Walker
  * 
  * Purpose: This sproc determines whether any rows exist in dependent tables for the specified PAYEE_T row.
  *                NOTE: At this time, there are no dependent tables for PAYEE_T.
  * 
  * Logic flow: Verify supplied @paye_id does exist in payee_t
  *
  * Return: number of rows affected for success
  *             4027 for failure
  */

CREATE PROCEDURE proc_payee_verify_dependents(@paye_id dom_id,
					@Dependent_Table VARCHAR(255) OUTPUT) WITH RECOMPILE AS

BEGIN

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

	RETURN 0

END
GO
setuser
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON dbo.proc_payee_verify_dependents TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_payee_verify_dependents TO Support
GO
