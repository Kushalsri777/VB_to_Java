SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_delete]
GO


/* Date: 10/14/2002
  * AUTHOR: Betsy Walker
  * 
  * Purpose: This sproc deletes a row from the payee_t table.
  * 
  * Logic flow: Verify supplied @paye_id does exist in payee_t
  *             Delete of the row from the payee_t table
  *
  * Return: number of rows affected for success
  *             4027, 4028 for failure
  */

CREATE PROCEDURE proc_payee_delete(@paye_id dom_id,
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

	DELETE payee_t
	WHERE paye_id = @paye_id

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error happened while deleting the PAYEE_T row for Payee ID: ' + ISNULL(CONVERT(VARCHAR(255), @paye_id), '<NULL>')
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END

GO
setuser
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON dbo.proc_payee_delete TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_payee_delete TO Support
GO
