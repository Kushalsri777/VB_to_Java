SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_claim_delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_claim_delete]
GO


/* Date: 10/14/2002
  * AUTHOR: Betsy Walker
  *
  * Purpose: This sproc deletes a row from the CLAIM_T table.
  *
  * Logic flow: Verify supplied @clm_id does exist in CLAIM_T
  *             Verify suppled @clm_id does not exist in PAYEE_T
  *             Delete of the row from the CLAIM_T table
  *
  * Return: number of rows affected for success
  *             4027, 4028, 4029 for failure
  */
CREATE PROCEDURE proc_claim_delete(@clm_id dom_id,
                    @Dependent_Table VARCHAR(255) OUTPUT) WITH RECOMPILE AS

BEGIN

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

    IF EXISTS(SELECT *
        FROM payee_t
        WHERE clm_id = @clm_id)
    BEGIN
        SELECT @Error_Message = 'A row exists for Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>') + ' in the PAYEE_T table.',
            @Dependent_Table = 'payee_t'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4029
    END

    DELETE claim_t
    WHERE clm_id = @clm_id

    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'An error happened while deleting the CLAIM_T row for Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>')
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

GRANT EXECUTE ON dbo.proc_claim_delete TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_claim_delete TO Support
GO
