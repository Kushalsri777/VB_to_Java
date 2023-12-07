SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_update]
GO


/* Date: 01/02/2003
  * Author: Betsy Walker
  *
  * Purpose: This stored procedure is written with the intent to update a row in the CLAIM_T table
  *
  * Logic flow:
  *             Verify supplied @paye_id does exist in the PAYEE_T table
  *             Verify supplied @calc_st_cd is valid
  *             Verify suppled @paye_st_cd is valid
  *             Verify suppled @clm_id is valid
  *             Update the row in the PAYEE_T table
  *
  * Return: number of rows affected for success
  *             4027, 4028 for failure
  *
  * Modifications:
  */

CREATE PROCEDURE proc_payee_update(@paye_id dom_id
                   ,@calc_st_cd dom_state_cd
                   ,@clm_id dom_id
                   ,@paye_addr_ln1_txt varchar(40)
                   ,@paye_addr_ln2_txt varchar(40)
                   ,@paye_care_of_txt varchar(40)
                   ,@paye_city_nm_txt varchar(25)
                   ,@paye_clm_int_amt decimal(11,2)
                   ,@paye_clm_int_rt decimal(11,5)
                   ,@paye_clm_pd_amt decimal(11,2)
                   ,@paye_dflt_ovrd_ind dom_ind
                   ,@paye_dthb_pmt_amt decimal(11,2)
                   ,@paye_full_nm varchar(40)
                   ,@paye_int_days_pd_num int
                   ,@paye_pmt_dt datetime
                   ,@paye_ssn_tin_num char(9)
                   ,@paye_ssn_tin_typ_cd char(1)
                   ,@paye_st_cd dom_state_cd
                   ,@paye_wthld_amt decimal(11,2)
                   ,@paye_wthld_rt decimal(11,5)
                   ,@paye_zip4_cd char(4)
                   ,@paye_zip_cd char(5)
                   ,@Invalid_Key VARCHAR(255) OUTPUT) WITH RECOMPILE AS

BEGIN

    DECLARE @Error_Number INTEGER
    DECLARE @Row_Count INTEGER
    DECLARE @Error_Message VARCHAR(255)


    IF NOT EXISTS(SELECT *
            FROM payee_t
            WHERE paye_id = @paye_id)
    BEGIN
        SELECT @Error_Message = 'Payee ID: ' + ISNULL(CONVERT(VARCHAR(255), @paye_id), '<NULL>') + ' does not exist in the PAYEE_T table'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4027
    END


    IF NOT EXISTS(SELECT *
            FROM state_t
            WHERE st_cd = @calc_st_cd)
    BEGIN
        SELECT @Error_Message = 'Calc State Code: ' + ISNULL(@calc_st_cd, '<NULL>') + ' is invalid.',
            @Invalid_Key = 'calc_st_cd'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4032
    END


    IF NOT EXISTS(SELECT *
            FROM state_t
            WHERE st_cd = @paye_st_cd)
    BEGIN
        SELECT @Error_Message = 'Payee State Code: ' + ISNULL(@paye_st_cd, '<NULL>') + ' is invalid.',
            @Invalid_Key = 'paye_st_cd'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4032
    END


    IF NOT(@clm_id IS NULL) AND
        NOT EXISTS(SELECT *
                FROM claim_t
                WHERE clm_id = @clm_id)
    BEGIN
        SELECT @Error_Message = 'Claim ID: ' + CONVERT(VARCHAR(255), @clm_id) + ' is invalid',
            @Invalid_Key = 'clm_id'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4032
    END


    UPDATE payee_t
	SET calc_st_cd = @calc_st_cd
       		,clm_id = @clm_id
		,lst_updt_dtm = GETDATE()
		,lst_updt_user_id = SUSER_SNAME()
		,paye_addr_ln1_txt = @paye_addr_ln1_txt
		,paye_addr_ln2_txt = @paye_addr_ln2_txt
		,paye_care_of_txt = @paye_care_of_txt
		,paye_city_nm_txt = @paye_city_nm_txt
		,paye_clm_int_amt = @paye_clm_int_amt
		,paye_clm_int_rt = @paye_clm_int_rt
		,paye_clm_pd_amt = @paye_clm_pd_amt
		,paye_dflt_ovrd_ind = @paye_dflt_ovrd_ind
		,paye_dthb_pmt_amt = @paye_dthb_pmt_amt
		,paye_full_nm = @paye_full_nm
		,paye_int_days_pd_num = @paye_int_days_pd_num
		,paye_pmt_dt = @paye_pmt_dt
		,paye_ssn_tin_num = @paye_ssn_tin_num
		,paye_ssn_tin_typ_cd = @paye_ssn_tin_typ_cd
		,paye_st_cd = @paye_st_cd
		,paye_wthld_amt = @paye_wthld_amt
		,paye_wthld_rt = @paye_wthld_rt
		,paye_zip4_cd = @paye_zip4_cd
		,paye_zip_cd = @paye_zip_cd
    WHERE paye_id = @paye_id

    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'An error happened while updating the PAYEE_T for the Payee ID: ' + ISNULL(@paye_id, '<NULL>')
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4028
    END

END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON dbo.proc_payee_update TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_payee_update TO Support
GO
