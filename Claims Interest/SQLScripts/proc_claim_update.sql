SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_claim_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_claim_update]
GO


/* Date: 01/02/2003
  * Author: Betsy Walker
  *
  * Purpose: This stored procedure is written with the intent to update a row in the CLAIM_T table
  *
  * Logic flow: Verify supplied @clm_id parameter does exist in the claim_t table
  *             Verify supplied @admn_syst_cd is valid
  *             Verify suppled @insd_dth_res_st_cd is valid
  *             Verify suppled @iss_st_cd is valid
  *             Verify suppled @pyco_typ_cd is valid
  *             Update the row in the claim_t table
  *
  * Return: number of rows affected for success
  *             4027, 4028 for failure
  *
  * Modifications:
  *   08/21/07 BAW - Per WRM6371, changed to accommodate new size of CLM_NUM (varchar 20)
  *   08/09/07 BAW - Per WRM5630, changed to accommodate new size of CLM_POL_NUM (char 15)
  *   03/06/03 BAW - Added support for new CLM_FOR_RES_DTH_IND column and for nullability of
  * 		         iss_st_cd and insd_dth_res_st_cd
  */

CREATE PROCEDURE proc_claim_update(@clm_id dom_id
                   ,@admn_syst_cd dom_admin_cd
                   ,@clm_insd_dth_dt datetime
                   ,@clm_insd_first_nm varchar(50)
                   ,@clm_insd_last_nm varchar(50)
                   ,@clm_insd_ssn_num char(9)
                   ,@clm_num varchar(20)
                   ,@clm_pol_num char(15)
                   ,@clm_proof_dt datetime
                   ,@clm_tot_clm_pd_amt decimal(13,2)
                   ,@clm_tot_dthb_pmt_amt decimal(11,2)
                   ,@clm_tot_int_amt decimal(11,2)
                   ,@clm_tot_wthld_amt decimal(11,2)
                   ,@insd_dth_res_st_cd dom_state_cd
                   ,@iss_st_cd dom_state_cd
                   ,@pyco_typ_cd dom_payr_co_cd
                   ,@clm_for_res_dth_ind dom_ind
                   ,@Invalid_Key VARCHAR(255) OUTPUT) WITH RECOMPILE AS

BEGIN

    DECLARE @Error_Number INTEGER
    DECLARE @Row_Count INTEGER
    DECLARE @Error_Message VARCHAR(255)

    IF NOT EXISTS(SELECT *
            FROM claim_t
            WHERE clm_id = @clm_id)
    BEGIN
        SELECT @Error_Message = 'Claim ID: ' + ISNULL(CONVERT(VARCHAR(255), @clm_id), '<NULL>') + ' does not exist in the CLAIM_T table'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4027
    END

    IF NOT EXISTS(SELECT *
            FROM admin_system_t
            WHERE admn_syst_cd = @admn_syst_cd)
    BEGIN
        SELECT @Error_Message = 'Admin System Code: ' + ISNULL(@admn_syst_cd, '<NULL>') + ' is invalid.',
            @Invalid_Key = 'admn_syst_cd'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4032
    END


    IF NOT (@insd_dth_res_st_cd IS NULL) AND
	NOT EXISTS(SELECT *
		FROM state_t
		WHERE st_cd = @insd_dth_res_st_cd)
    BEGIN
        SELECT @Error_Message = 'Insured Death Residence State Code: ' + ISNULL(@insd_dth_res_st_cd, '<NULL>') + ' is invalid.',
            @Invalid_Key = 'insd_dth_res_st_cd'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4032
    END


    IF NOT (@iss_st_cd IS NULL) AND
	NOT EXISTS(SELECT *
		FROM state_t
		WHERE st_cd = @iss_st_cd)
    BEGIN
        SELECT @Error_Message = 'Issue State Code: ' + ISNULL(@iss_st_cd, '<NULL>') + ' is invalid.',
            @Invalid_Key = 'iss_st_cd'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4032
    END


    IF NOT EXISTS(SELECT *
            FROM payor_company_type_t
            WHERE pyco_typ_cd = @pyco_typ_cd)
    BEGIN
        SELECT @Error_Message = 'Payor Company Type Code: ' + ISNULL(@pyco_typ_cd, '<NULL>') + ' is invalid.',
            @Invalid_Key = 'pyco_typ_cd'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4032
    END


    UPDATE claim_t
    SET admn_syst_cd = @admn_syst_cd
            ,clm_insd_dth_dt = @clm_insd_dth_dt
            ,clm_insd_first_nm = @clm_insd_first_nm
            ,clm_insd_last_nm = @clm_insd_last_nm
            ,clm_insd_ssn_num = @clm_insd_ssn_num
            ,clm_num = @clm_num
            ,clm_pol_num = @clm_pol_num
            ,clm_proof_dt = @clm_proof_dt
            ,clm_tot_clm_pd_amt = @clm_tot_clm_pd_amt
            ,clm_tot_dthb_pmt_amt = @clm_tot_dthb_pmt_amt
            ,clm_tot_int_amt = @clm_tot_int_amt
            ,clm_tot_wthld_amt = @clm_tot_wthld_amt
            ,insd_dth_res_st_cd = @insd_dth_res_st_cd
            ,iss_st_cd = @iss_st_cd
            ,lst_updt_dtm = GETDATE()
            ,lst_updt_user_id = SUSER_SNAME()
            ,pyco_typ_cd = @pyco_typ_cd
            ,clm_for_res_dth_ind = @clm_for_res_dth_ind
    WHERE clm_id = @clm_id

    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'An error happened while updating the CLAIM_T for the Claim ID: ' + ISNULL(@clm_id, '<NULL>')
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4028
    END

END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON [dbo].proc_claim_update TO AppRoleClaims
GO
GRANT EXECUTE ON [dbo].proc_claim_update TO Support
GO
