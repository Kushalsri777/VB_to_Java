SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_current_rate_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_current_rate_select]
GO


/****************************************************************************************
** Created By/Date:		dbo - 05/30/2003
**
** Purpose: 			This sproc selects certain columns from the CURRENT_RATE_T table.
**                      It is used by the Payee screen to drive its calculations.
**
** Date of Release: 	06/2003
** Current Version:		1.0
**
** Called by:			fnGetCurrentRate( ) in frmPayee.frm
**
** Calls:				N/A
**
** =================
** Inputs
** =================
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
** - A double containing the Current Rate.
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
CREATE PROCEDURE dbo.proc_current_rate_select(
    @paye_pmt_dt datetime
   ,@curr_int_rt decimal(11,5) OUTPUT) WITH RECOMPILE AS
BEGIN

	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	IF @paye_pmt_dt IS NULL
	BEGIN
		SELECT @Error_Message = 'The Payee''s Date of Payment cannot be <NULL>'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

    SELECT @curr_int_rt = curr_int_rt
    FROM CURRENT_RATE_T
    WHERE curr_rt_eff_dt = (
        		SELECT MAX(curr_rt_eff_dt)
        		FROM CURRENT_RATE_T
    			WHERE curr_rt_eff_dt <= @paye_pmt_dt
        		)

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting info from the CURRENT_RATE_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END
GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

GRANT EXECUTE ON dbo.proc_current_rate_select TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_current_rate_select TO Support
GO
