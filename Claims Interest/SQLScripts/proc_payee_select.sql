SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_payee_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_payee_select]
GO


CREATE PROCEDURE [dbo].proc_payee_select(@paye_id dom_id) WITH RECOMPILE AS
BEGIN
 /* Date: 10/14/2002
  * Author: Betsy Walker
  *
  * Purpose: This stored procedure selects a row on the PAYEE_T table.
  *
  * Logic flow:
  *             Verify supplied @paye_id exists in the PAYEE_T table
  *             Perform the select
  *             Perform error checking
  *
  * Return: number of rows affected for success
  *             4027, 4028 for failure
  * 
  * Modification History:
  */
    DECLARE @Error_Number INTEGER
    DECLARE @Row_Count INTEGER
    DECLARE @Error_Message VARCHAR(255)

    IF NOT EXISTS(SELECT *
        FROM payee_t
        WHERE paye_id = @paye_id)
    BEGIN
        SELECT @Error_Message = 'Payee ID: ' + ISNULL(CONVERT(VARCHAR(255), @paye_id), '<NULL>') + ' does not exist in the PAYEE_T table.'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4027
    END


    SELECT calc_st_cd
            ,clm_id
            ,lst_updt_dtm
            ,lst_updt_user_id
            ,paye_addr_ln1_txt
            ,paye_addr_ln2_txt
            ,paye_care_of_txt
            ,paye_city_nm_txt
            ,paye_clm_int_amt
            ,paye_clm_int_rt
            ,paye_clm_pd_amt
            ,paye_dflt_ovrd_ind
            ,paye_dthb_pmt_amt
            ,paye_full_nm
            ,paye_id
            ,paye_int_days_pd_num
            ,paye_pmt_dt
            ,paye_ssn_tin_num
            ,paye_ssn_tin_typ_cd
            ,paye_st_cd
            ,paye_wthld_amt
            ,paye_wthld_rt
            ,paye_zip4_cd
            ,paye_zip_cd
    FROM payee_t
    WHERE P.paye_id = @paye_id


    SELECT @Error_Number = @@ERROR,
        @Row_Count = @@ROWCOUNT
    IF @Error_Number = 0
    BEGIN
        RETURN @Row_Count
    END
    ELSE
    BEGIN
        SELECT @Error_Message = 'An error happened while selecting the PAYEE_T for Payee ID: ' + ISNULL(CONVERT(VARCHAR(255), @paye_id), '<NULL>')
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4028
    END

END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON [dbo].proc_payee_select TO AppRoleClaims
GO
GRANT EXECUTE ON [dbo].proc_payee_select TO Support
GO
