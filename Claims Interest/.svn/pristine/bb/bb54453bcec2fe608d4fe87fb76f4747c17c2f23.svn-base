SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_select4]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_select4]
GO

/****************************************************************************************
** Created By/Date:		dbo - 04/16/2003
**
** Purpose: 			This stored procedure returns a recordset containing PAYEE_T data
**                      for all Payees associated with the specified CLM_ID.
**
** Date of Release: 	04/30/2003
** Current Version:		0.1
**
** Called by:			ctpyePayee.cls in Claims Interest front-end
**
** Calls:				N/A
**
** =================
** Inputs
** =================
** - CLM_ID, the identity value that identifies the Claim for which Payees should be retrieved
**
**
** =================
** Local Variables
** =================
** - @Error_Number, the current SQL Error code (@@ERROR) value
** - @Row_Count, the number of rows affected by the last SQL statement
** - @Error_Message, the text of the error message, if any, to display
**
**
** =================
** Outputs
** =================
** - Recordset containing subset of columns for all PAYEE_T rows for the specified CLM_ID
**
**
** =================
** Returns
** =================
** - 4027, indicating an invalid input parameter was supplied
** - 4028, indicating an unexpected error occured.
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
**
**
****************************************************************************************/
CREATE PROCEDURE dbo.proc_payee_select4(@clm_id dom_id) WITH RECOMPILE AS

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

    SELECT paye_full_nm
			,paye_addr_ln1_txt
			,paye_addr_ln2_txt
    		,calc_st_cd
            ,paye_pmt_dt
            ,paye_ssn_tin_num
            ,paye_clm_int_amt
            ,paye_clm_pd_amt
            ,paye_dthb_pmt_amt
            ,paye_clm_int_rt
            ,paye_wthld_rt
            ,paye_wthld_amt
            ,paye_id
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
        SELECT @Error_Message = 'Error number: ' + CONVERT(VARCHAR(255), @Error_Number) + ' - error while selecting Payees for Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>')
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4028
    END

END


GO

SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO

GRANT EXECUTE ON dbo.proc_payee_select4 TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_payee_select4 TO Support
GO
