SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_lu_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_lu_select]
GO


/* Date: 10/14/2002
  * Author: Betsy Walker
  *
  * Purpose: This stored procedure selects data from the PAYEE_T table
  *          in order to populate a Lookup combo box. This could return
  *          zero rows, if no Payees have yet been defined for the
  *          specified Claim Number.
  *
  * Logic flow: Perform the select of the desired columns from the PAYEE_T table
  *             Perform error checking
  *
  * Return: number of rows for success
  *             4028 for failure
  */

CREATE PROCEDURE proc_payee_lu_select(@clm_id dom_id) WITH RECOMPILE AS

BEGIN

    DECLARE @Error_Number INTEGER
    DECLARE @Row_Count INTEGER
    DECLARE @Error_Message VARCHAR(255)

    SELECT paye_full_nm
	,paye_id
	,clm_id
    FROM payee_t
    WHERE clm_id = @clm_id
    ORDER BY paye_full_nm

    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'An error occurred while selecting lookup info from the CLAIM_T table for the Claim ID: ' +  ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>')
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

GRANT EXECUTE ON dbo.proc_payee_lu_select TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_payee_lu_select TO Support
GO
