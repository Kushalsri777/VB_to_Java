SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_select2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_select2]
GO

setuser N'dbo'
GO

/* Date: 10/14/2002
  * Author: Betsy Walker
  *
  * Purpose: This stored procedure determines if any payees for the specified Claim have a Date of Payment
  *                earlier than the Date of Proof (from CLAIM_T)
  *
  * Logic flow:
  *             Verify the Claim ID specified is not Null
  *             Verify the Claim Proof Dt specified is not Null
  *             Perform the select
  *             Perform error checking
  *
  * Return: number of rows affected for success
  *             4027, 4028 for failure
  */
CREATE PROCEDURE proc_payee_select2(@clm_id dom_id, @clm_proof_dt datetime, @nbr_of_payees int OUTPUT) WITH RECOMPILE AS

BEGIN

    DECLARE @Error_Number INTEGER
    DECLARE @Row_Count INTEGER
    DECLARE @Error_Message VARCHAR(255)

    IF @clm_id IS NULL
    BEGIN
        SELECT @Error_Message = 'Claim ID cannot be <NULL>'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4027
    END

    IF @clm_proof_dt IS NULL
    BEGIN
        SELECT @Error_Message = 'Date of Proof cannot be <NULL>'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4027
    END

    SELECT @nbr_of_payees = (SELECT COUNT(paye_id)
    FROM payee_t
    WHERE paye_pmt_dt < @clm_proof_dt AND clm_id = @clm_id)

    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'Error number: ' + CONVERT(VARCHAR(255), @Error_Number) + ' - error while selecting Payees for Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>')
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

GRANT EXECUTE ON dbo.proc_payee_select2 TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_payee_select2 TO Support
GO
