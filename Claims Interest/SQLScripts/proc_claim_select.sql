SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_claim_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_claim_select]
GO


/* Date: 10/14/2002
  * Author: Betsy Walker
  *
  * Purpose: This stored procedure selects a row on the CLAIM_T table.
  *
  * Logic flow:
  *             Verify supplied @clm_id exists in the CLAIM_T table
  *             Perform the select
  *             Perform error checking
  *
  * Return: number of rows affected for success
  *             4027, 4028 for failure
  * 
  * Modifications:
  *   03/06/03 BAW - Added support for new CLM_FOR_RES_DTH_IND column
  */
CREATE PROCEDURE proc_claim_select(@clm_id dom_id) WITH RECOMPILE AS

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


    SELECT admn_syst_cd
	,clm_id
	,clm_insd_dth_dt
	,clm_insd_first_nm
	,clm_insd_last_nm
	,clm_insd_ssn_num
	,clm_num
	,clm_pol_num
	,clm_proof_dt
	,clm_tot_clm_pd_amt
	,clm_tot_dthb_pmt_amt
	,clm_tot_int_amt
	,clm_tot_wthld_amt
	,insd_dth_res_st_cd
	,iss_st_cd
	,lst_updt_dtm
	,lst_updt_user_id
	,pyco_typ_cd
	,clm_for_res_dth_ind
    FROM claim_t
    WHERE clm_id = @clm_id

    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'An error happened while selecting the CLAIM_T for Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>')
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

GRANT EXECUTE ON dbo.proc_claim_select TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_claim_select TO Support
GO
