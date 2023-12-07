SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_claim_lu_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_claim_lu_select]
GO


/****************************************************************************************
** Created By/Date:		dbo - 10/14/2002
** 
** Purpose: 			This stored procedure selects data from the CLAIM_T table in 
**                      order to populate the Claim Lookup combo box on the Insured
**                      screen of the Claims Interest front-end
**
** Date of Release: 	04/11/2003
** Current Version:		0.3
**
** Called by:			ctclmClaim.cls VB class in Claims Interest front-end
**
**
** Calls:				N/A
**
** =================
** Inputs 
** =================
** - N/A
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
** - Recordset containing subset of columns for all CLAIM_T rows
**
**
** =================
** Returns
** =================
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
** 04/11/2003 dbo     v2.4     Split this sproc into 3 pieces (one for each Lookup
**                              combo box on the Insured screen of the Claims Interest
**                              front-end, so each has only the columns it needs.
**                              This was done to address performance issues when 
**                              initially displaying the Insured screen.                              
**
****************************************************************************************/
  

CREATE PROCEDURE dbo.proc_claim_lu_select WITH RECOMPILE AS

BEGIN

    DECLARE @Error_Number	INTEGER
    DECLARE @Row_Count		INTEGER
    DECLARE @Error_Message	VARCHAR(255)

    SELECT clm_num
        ,clm_id
    FROM claim_t
    ORDER BY clm_num

    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'An error occured while selecting lookup info from the CLAIM_T table.'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4028
    END

END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON dbo.proc_claim_lu_select TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_claim_lu_select TO Support
GO
