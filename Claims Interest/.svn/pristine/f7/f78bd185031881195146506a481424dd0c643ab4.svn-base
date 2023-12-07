SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_totals_for_claim]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_totals_for_claim]
GO


/****************************************************************************************
** Created By/Date:		K758 - 05/30/2003
** 
** Purpose: 			This sproc calculates totals for each field accumulated across
**                      all Payees for a given Claim and stored in the CLAIM_T table.
**
** Date of Release: 	06/2003
** Current Version:		1.0
**
** Called by:			fnCalcTotalsForAllPayees( ) in frmInsured
**
** Calls:				N/A
**
** =================
** Inputs 
** =================
** - @clm_id			The Claim ID identifying the claim whose totals should be calculated
**
**
** =================
** Local Variables
** =================
** - @Error_Number		The return code
** - @Row_Count			The number of rows affected by a SQL command
** - @Error_Message		The error message text
**
**
** =================
** Outputs
** =================
** - @ClmTotDthbPmtAmt	The Death Benefit paid across all Payees on this claim
** - @ClmTotWthldAmt	The Amt of Interest Withheld across all Payees on this claim
** - @ClmTotIntAmt		The Amt of Claims Interest Paid across all Payees on this claim
** - @ClmTotClmPdAmt	The Total amount (DB - W/H Int + Claims Int) paid across all Payees on this claim

**
**
** =================
** Returns
** =================
** - 4027 if any of the input parameters are invalid.
** - 4028 if any other unexpected error was encountered.
**
**
** =================
** Additional Notes
** =================
**
**
** =================
** Revision history
** =================
**
** Date       Author   Tag      Purpose
** ---------- -------  -------- --------------------------------------------------------
** 05/30/2003 K758				Initial creation.
**
***************************************************************************************
*/
CREATE PROCEDURE dbo.proc_payee_totals_for_claim(@clm_id dom_id) WITH RECOMPILE AS
BEGIN

	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	-- Ensure something valid was input for each input parameter
	IF NOT EXISTS(SELECT *
			FROM claim_t
			WHERE clm_id = @clm_id)
	BEGIN
		SELECT @Error_Message = 'Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>') + ' does not exist in the CLAIM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END
    
    -- Calculate aggregate totals across all Payees for the specified Claim
    SELECT SUM(paye_dthb_pmt_amt)	AS ClmTotDthbPmtAmt
		,SUM(paye_wthld_amt)		AS ClmTotWthldAmt
		,SUM(paye_clm_int_amt)		AS ClmTotIntAmt
		,SUM(paye_clm_pd_amt)		AS ClmTotClmPdAmt
    FROM payee_t
    WHERE clm_id = @clm_id

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting info from the PAYEE_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END
GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

GRANT EXECUTE ON dbo.proc_payee_totals_for_claim TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_payee_totals_for_claim TO Support
GO
