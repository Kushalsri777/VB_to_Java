if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_admin_system_select2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_admin_system_select2]
GO


/* Date: 01/02/2003
  * Author: Betsy Walker
  * 
  * Purpose: This stored procedure is written with the intent to select the metadata-related columns from the ADMIN_SYSTEM_T table, 
  *          so it can be used to "drive" validations/processing performed by the Insured screen.
  * 
  * Logic flow: Perform the select
  *                   Perform error checking
  *
  * Return: number of rows for success
  *				4028 for failure
  *             4028 for failure
  */

CREATE PROCEDURE proc_admin_system_select2(@admn_syst_cd dom_admin_cd,
	@MinLenth		SMALLINT OUTPUT, 	
	@MaxLength		SMALLINT OUTPUT,
	@DfltPycoTypDsc	dom_dsc OUTPUT,
	@TaxRptgInd		dom_ind OUTPUT) WITH RECOMPILE AS
BEGIN

	DECLARE @Error_Number 	INTEGER
	DECLARE @Row_Count 		INTEGER
	DECLARE @Error_Message 	VARCHAR(255)

    IF @admn_syst_cd IS NULL
    BEGIN
        SELECT @Error_Message = 'Admin System Cd cannot be <NULL>'
        RAISERROR(@Error_Message, 16, 1)
        RETURN 4027
    END

	SELECT @MinLenth = admn_syst_id_min_lgth_num FROM ADMIN_SYSTEM_T
							WHERE admn_syst_cd = @admn_syst_cd

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number != 0
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting Min Length from the ADMIN_SYSTEM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

	SELECT @MaxLength = admn_syst_id_max_lgth_num FROM ADMIN_SYSTEM_T
							WHERE admn_syst_cd = @admn_syst_cd

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number != 0
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting Max Length from the ADMIN_SYSTEM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

	SELECT @DfltPycoTypDsc = PCT.pyco_typ_dsc 
							FROM ADMIN_SYSTEM_T 			AST
							INNER JOIN payor_company_type_t PCT  ON AST.dflt_pyco_typ_cd = PCT.pyco_typ_cd
							WHERE admn_syst_cd = @admn_syst_cd

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number != 0
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting Payor Company Type from the ADMIN_SYSTEM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

	SELECT @TaxRptgInd = admn_syst_tax_rptg_ind FROM ADMIN_SYSTEM_T
							WHERE admn_syst_cd = @admn_syst_cd

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number != 0
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting Tax Rptg Ind from the ADMIN_SYSTEM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

	RETURN 0
END
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON [dbo].proc_admin_system_select2 TO AppRoleClaims
GO
GRANT EXECUTE ON [dbo].proc_admin_system_select2 TO Support
GO
