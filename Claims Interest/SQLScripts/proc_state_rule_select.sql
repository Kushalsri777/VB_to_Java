SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_state_rule_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_state_rule_select]
GO


/****************************************************************************************
** Created By/Date:		K758 - 05/30/2003
** 
** Purpose: 			This sproc selects certain columns from the STATE_RULE_T table.
**                      It is used by the Payee screen to drive its calculations.
**
** Date of Release: 	06/2003
** Current Version:		1.0
**
** Called by:			fnGetStateRule( ) in modGeneral.bas
**
** Calls:				N/A
**
** =================
** Inputs 
** =================
** - @lob_cd			The line-of- business (I for Individual, G for Group)
** - @st_cd				The 2-character Postal Abbreviation representing the state whose rule
**						should be obtained
** - @paye_pmt_dt		The date whose rule should be obtained
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
** - An ADODB.Recordset containing the selected row. (There should only be 1 row.)
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
** mm/dd/yyyy User ID  xxxxxxxx xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx  
**
***************************************************************************************
*/
CREATE PROCEDURE dbo.proc_state_rule_select(
	@lob_cd dom_lob_cd, 
	@st_cd dom_state_cd, 
	@paye_pmt_dt datetime) WITH RECOMPILE AS
BEGIN

	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	-- Ensure something valid was input for each input parameter
	IF NOT EXISTS(SELECT *
			FROM line_of_business_t
			WHERE lob_cd = @lob_cd)
	BEGIN
		SELECT @Error_Message = 'LOB Code: ' + ISNULL(CONVERT(VARCHAR(255), @lob_cd), '<NULL>') + ' does not exist in the LINE_OF_BUSINESS_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	IF NOT EXISTS(SELECT *
			FROM state_t
			WHERE st_cd = @st_cd)
	BEGIN
		SELECT @Error_Message = 'State Code: ' + ISNULL(CONVERT(VARCHAR(255), @st_cd), '<NULL>') + ' does not exist in the STATE_T table'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	IF @paye_pmt_dt IS NULL
	BEGIN
		SELECT @Error_Message = 'The Payee''s Date of Payment cannot be <NULL>'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

    SELECT lob_cd
    	,st_cd
    	,strl_eff_dt
    	,calc_idtyp_cd
    	,reqd_idtyp_cd
    	,irule_cd
    	,strl_end_dt
    	,strl_int_rptg_flr_amt
    	,strl_int_calc_ofst_num
    	,strl_int_reqd_ofst_num
    	,strl_int_rule_amt
    	,strl_spcl_instr_txt
    FROM state_rule_t
    WHERE lob_cd = @lob_cd
        AND st_cd = @st_cd
        AND strl_eff_dt = (
        		SELECT MAX(strl_eff_dt) 
        		FROM state_rule_t 
    			WHERE lob_cd = @lob_cd
        			AND st_cd = @st_cd
        			AND strl_eff_dt <= @paye_pmt_dt
        		)

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting info from the STATE_RULE_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END
GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

GRANT EXECUTE ON dbo.proc_state_rule_select TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_state_rule_select TO Support
GO
