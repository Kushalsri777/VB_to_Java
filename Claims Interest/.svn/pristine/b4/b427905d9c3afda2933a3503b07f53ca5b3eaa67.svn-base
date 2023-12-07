SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_tax_file_layout_generate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_tax_file_layout_generate]
GO


/* Date: 03/11/2003
  * Author: Betsy Walker
  *
  * Purpose: This stored procedure builds a recordset of Payee & Claim information in anticipation
  *                 of using it to prepare a tax file in the necessary I.R.S. TVTAXFORM layout for the PC.
  *
  * Logic flow:
  *             Verify supplied @paye_pmt_dt_from_date is valid
  *             Verify supplied @paye_pmt_dt_to_date is valid
  *             Perform the SELECT
  *
  * Return: number of rows affected for success
  *             4028, 4031, 4032 for failure
  *
  * Modifications:
  *  03/14/03 BAW - Per 03/14/03 e-mail from Michelle Wilkosky, changed the SELECT 
  *                 statement so that the *current* STATE_RULE_T row is always used.
  *  10/22/03 BAW - Per 10/22/03 e-mail from Michelle Wilkosky, changed so that the
  *                 recordset reflects the 2003 TVTAXFORM layout in its entirety.
  *  11/03/03 BAW - Per 11/03/03 telephone conversation with Michelle Wilkosky, add a 
  *                 temporary "AND (AST.lob_cd = 'I'" condition to the WHERE clause
  *                 to prevent tax reporting on Group claims. The retention of
  *                 this logic will be reassessed on a year-to-year basis.
  *  12/06/03 BAW   Changed a comment line so it correctly states that the app
  *                 expects the Payee Residence State to be in Field9, not Field11.
  *  12/13/05 BAW   Per BZ1, added support for Keyport policies which have a LOB_CD = 'I',
  *                 so no tax reporting is done.
  *  10/20/06 BAW   Changed the SQL to use a UNION to ensure that duplicate rows don't 
  *                 appear if a given state has more than 1 rule in effect for the
  *                 reporting period.
  *  08/09/07 BAW   Changed to use new metadata column (admn_syst_tax_rptg_ind of 
  *                 admin_system_t) to determine whether tax reporting should be done.
  *                 Changed to put first part of Claim # in field 16 (the Account
  *                 Number field) with residual, if any, placed into Fld18.
  */
CREATE PROCEDURE [dbo].proc_tax_file_layout_generate(@paye_pmt_dt_from_date datetime
	,@paye_pmt_dt_to_date datetime) WITH RECOMPILE AS

BEGIN

	DECLARE @Error_Number       INTEGER
	DECLARE @Row_Count          INTEGER
	DECLARE @Error_Message      VARCHAR(255)

	DECLARE @ConstSpace         CHAR(1)
	DECLARE @ConstZero			CHAR(1)
	DECLARE @ConstUnused		CHAR(1)

	-- Make sure the following fields are defined with the same width as the data
	-- used to populate it in the subsequent SELECT statement
	DECLARE @Fld8LiteralValue	CHAR(2)
	DECLARE @Fld10LiteralValue  CHAR(2)
	DECLARE @Fld12LiteralValue  CHAR(1)
	DECLARE @Fld16LiteralValue1	CHAR(5)
	DECLARE @Fld16LiteralValue2	CHAR(1)
	DECLARE @Fld17LiteralValue  CHAR(2)
	DECLARE @Fld18LiteralValue  CHAR(27)
	DECLARE @Fld20LiteralValue  CHAR(11)
	DECLARE @Fld21LiteralValue  CHAR(11)
	DECLARE @Fld23LiteralValue  CHAR(11)
	DECLARE @Fld24LiteralValue  CHAR(11)
	DECLARE @Fld25LiteralValue	CHAR(48)
	DECLARE @Fld26LiteralValue	CHAR(188)
	DECLARE @Fld27Unused		CHAR(1)
	DECLARE @Fld28Unused		CHAR(1)

	 -- Initialize constants. The lengths must jive with PC (vs. mainframe) layout of the 
	-- TVTAXFORM as set by the I.R.S.
	SELECT 
		@ConstSpace = ' '		    	 -- literal space
		,@ConstZero = '0'		    	 -- literal zero
		,@ConstUnused = '!'				 -- literal used to identify an unused field
		,@Fld8LiteralValue = '  '		 -- Spaces as filler in this unused area of TVTAXFORM
		,@Fld10LiteralValue = '  '		 -- Spaces as filler in this unused area of TVTAXFORM
		,@Fld12LiteralValue = ' '		 -- Address Type (blank = U.S.)
		,@Fld16LiteralValue1 = 'S1271'	 -- Sun Code to which I.R.S. inquiries s/b sent
		,@Fld16LiteralValue2 = 'S'       -- literal S  (purpose unknown)
		,@Fld17LiteralValue = 'US'		 -- Currency Type (US)
		,@Fld18LiteralValue = SPACE(27)	 -- Spaces as filler in this unused area of TVTAXFORM
		,@Fld20LiteralValue = '00000000000' -- Box 2 amount
		,@Fld21LiteralValue = '00000000000' -- Box 3 amount
		,@Fld23LiteralValue = '00000000000' -- Box 5 amount
		,@Fld24LiteralValue = '00000000000' -- Box 6 amount
		,@Fld25LiteralValue = SPACE(48)	 -- Spaces as filler in this unused area of TVTAXFORM
		,@Fld26LiteralValue = SPACE(188) -- Unused area

	-- These fields aren't actually used. They're here as placeholders...so if the TVTAXFORM
	-- is ever expanded, a few fields can be added with only the sproc (not the app itself)
	-- requiring changes. (The app, by the way, will just write each Field in the ADODB.Recordset
	-- to the file, as long as it doesn't contain '!' (the @ConstUnused value).
	SELECT
		@Fld27Unused = @ConstUnused
		,@Fld28Unused = @ConstUnused


	-- Ensure the From/To Date input parameters were supplied as non-Null values
	IF @paye_pmt_dt_from_date IS NULL
	BEGIN
		SELECT @Error_Message = 'Date of Payment From Date parameter cannot be <NULL>'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	IF @paye_pmt_dt_to_date IS NULL
	BEGIN
		SELECT @Error_Message = 'Date of Payment To Date parameter cannot be <NULL>'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	-- Use the dbo.func_RPad and dbo.func_LPad user-defined functions (UDFs) to do
	-- right and left padding of the source column so it is formatted as the output
	-- column must be.
	--
	-- The syntax of these functions is:
	--		dbo.func_RPad(<expression>, <size of output field>, <pad char>)
	--		dbo.func_LPad(<expression>, <size of output field>, <pad char>, <UseImplicitDecimals>)
	-- where <expression> is a variable or column name or expression
	--       <size of output field> is an Integer which represents the width of the
	--              column as required by the I.R.S.' TVTAXFORM layout for PC files
	--       <pad char> is a CHAR(1) constant or literal to indicate what to
	--              pad the <expression> with, e.g., @ConstZero with func_LPad() to pad
	--              with leading zeroes --or-- @ConstSpace with the func_RPad() to pad
	--              with trailing spaces.
	--       <UseImplicitDecimals> is a CHAR(1) constant or literal = 'Y' if the input
	--              value is an amount that should have its decimal point removed.
	-- 
	-- You can skip using the RPad and LPad UDFs if the source column is defined as
	-- CHAR *and* Not Null. (VARCHARS and numeric fields will need the padding, even if Not Null.)
	--
	-- NOTE: The app expects to find the Payee Residence State in Fld9, the Box 1 Amount in Fld19,
	--       and the Box 4 Amount in Fld22, so if this ever shifts to another location, the app will require changes.
	-- Find rows that use a rule with a matching LOB (see ON clause re: STATE_RULE_T)
	SELECT C.pyco_typ_cd                                        AS Fld01 --  1. Payor Company
		,C.admn_syst_cd                                         AS Fld02 --  2. Admin System
		,dbo.func_RPad(P.paye_full_nm, 40, @ConstSpace)         AS Fld03 --  3. Payee Name         
		,dbo.func_RPad(P.paye_care_of_txt, 40, @ConstSpace)		AS Fld04 --  4. Care Of portion of Payee's residence address
		,dbo.func_RPad(P.paye_addr_ln1_txt, 40, @ConstSpace)	AS Fld05 --  5. Line 1 of Payee's residence address
		,dbo.func_RPad(P.paye_addr_ln2_txt, 40, @ConstSpace)	AS Fld06 --  6. Line 2 of Payee's residence address
		,dbo.func_RPad(P.paye_city_nm_txt, 25, @ConstSpace)		AS Fld07 --  7. City of Payee's residence address
		,@Fld8LiteralValue										AS Fld08 --  8. Unused  (2 bytes)
		,dbo.func_RPad(P.paye_st_cd, 2, @ConstSpace)			AS Fld09 --  9. State of Payee's residence address
		,@Fld10LiteralValue										AS Fld10 -- 10. Unused  (2 bytes)
		,dbo.func_RPad(P.paye_zip_cd + ISNULL(P.paye_zip4_cd,''), 9, @ConstSpace)	AS Fld11 -- 11. Zip/Zip4 of Payee's residence address
		,@Fld12LiteralValue										AS Fld12 -- 12. Address Type (blank = U.S.)
		,dbo.func_RPad(P.paye_st_cd, 2, @ConstSpace)			AS Fld13 -- 13. State of Withholding
		,dbo.func_RPad(P.paye_ssn_tin_num, 9, @ConstSpace)		AS Fld14 -- 14. Payee SSN/TIN
		,dbo.func_RPad(P.paye_ssn_tin_typ_cd, 1, @ConstSpace)	AS Fld15 -- 15. Payee SSN/TIN Type (P=person, B=business)
		,Fld16 =
			CASE 														 
				WHEN AST.lob_cd = 'I' THEN								 -- 16. Indiv LOB: SunCode, first 10 chars of Pol#
					@Fld16LiteralValue1
					+ LEFT(C.clm_num,10)
				WHEN AST.lob_cd = 'G' AND LEN(C.clm_num) < 16 THEN		 -- 16. Group LOB: 15-char claim number (no spillover)
					dbo.func_RPad(C.clm_num, 15, @ConstSpace)
				ELSE													 -- 16. Group LOB: 15-char claim number (spillover in Fld18)
					LEFT(C.clm_num, 15)
			END
		,@Fld17LiteralValue										AS Fld17 -- 17. Currency Type (US), then blank for TIN2 Notice
		,Fld18 =
			CASE 														 
				WHEN AST.lob_cd = 'I' THEN								 -- 18. Indiv LOB: Spillover of Claim# and literal 'S'
					dbo.func_RPad(SUBSTRING(C.clm_num, 11, 5) + @Fld16LiteralValue2, 27, @ConstSpace)
				WHEN AST.lob_cd = 'G' AND LEN(C.clm_num) < 16 THEN	-- 18. Group LOB: Spaces (no spillover of Claim#)
					@Fld18LiteralValue
				ELSE												-- 18. Group LOB: Spillover of Claim #
					dbo.func_RPad(SUBSTRING(C.clm_num, 16, LEN(C.clm_num)-15), 27, @ConstSpace)
			END
		,dbo.func_LPad(P.paye_clm_int_amt, 11, @ConstZero, 'Y') AS Fld19 -- 19. Box 1 amount (Claim Interest Paid to this Payee)
		,@Fld20LiteralValue										AS Fld20 -- 20. Box 2 amount
		,@Fld21LiteralValue										AS Fld21 -- 21. Box 3 amount
		,dbo.func_LPad(P.paye_wthld_amt, 11, @ConstZero, 'Y')	AS Fld22 -- 22. Box 4 amount (Interest withheld from Claim Interest paid to this Payee)
		,@Fld23LiteralValue										AS Fld23 -- 23. Box 5 amount
		,@Fld24LiteralValue										AS Fld24 -- 24. Box 6 amount
		,@Fld25LiteralValue										AS Fld25
		,@Fld26LiteralValue										AS Fld26
		,@Fld27Unused											AS Fld29
		,@Fld28Unused											AS Fld30
	FROM payee_t					P
	INNER JOIN claim_t				C		ON	P.clm_id = C.clm_id
	INNER JOIN state_t				S		ON	P.paye_st_cd = S.st_cd
	INNER JOIN admin_system_t		AST		ON	C.admn_syst_cd = AST.admn_syst_cd
	INNER JOIN state_rule_t			SR1		ON	P.paye_st_cd = SR1.st_cd AND AST.lob_cd = SR1.lob_cd
	WHERE (P.paye_pmt_dt BETWEEN @paye_pmt_dt_from_date AND @paye_pmt_dt_to_date)
		AND (P.paye_clm_int_amt >= SR1.strl_int_rptg_flr_amt)
		AND (AST.admn_syst_tax_rptg_ind = 'Y')
		AND (P.paye_pmt_dt >= SR1.strl_eff_dt)
		AND (SR1.strl_end_dt IS NULL OR P.paye_pmt_dt <= SR1.strl_end_dt)


	UNION

	-- Find rows that use the default rule (LOB=I) (see ON clause re: STATE_RULE_T)
	SELECT C.pyco_typ_cd                                        AS Fld01 --  1. Payor Company
		,C.admn_syst_cd                                         AS Fld02 --  2. Admin System
		,dbo.func_RPad(P.paye_full_nm, 40, @ConstSpace)         AS Fld03 --  3. Payee Name         
		,dbo.func_RPad(P.paye_care_of_txt, 40, @ConstSpace)		AS Fld04 --  4. Care Of portion of Payee's residence address
		,dbo.func_RPad(P.paye_addr_ln1_txt, 40, @ConstSpace)	AS Fld05 --  5. Line 1 of Payee's residence address
		,dbo.func_RPad(P.paye_addr_ln2_txt, 40, @ConstSpace)	AS Fld06 --  6. Line 2 of Payee's residence address
		,dbo.func_RPad(P.paye_city_nm_txt, 25, @ConstSpace)		AS Fld07 --  7. City of Payee's residence address
		,@Fld8LiteralValue										AS Fld08 --  8. Unused  (2 bytes)
		,dbo.func_RPad(P.paye_st_cd, 2, @ConstSpace)			AS Fld09 --  9. State of Payee's residence address
		,@Fld10LiteralValue										AS Fld10 -- 10. Unused  (2 bytes)
		,dbo.func_RPad(P.paye_zip_cd + ISNULL(P.paye_zip4_cd,''), 9, @ConstSpace)	AS Fld11 -- 11. Zip/Zip4 of Payee's residence address
		,@Fld12LiteralValue										AS Fld12 -- 12. Address Type (blank = U.S.)
		,dbo.func_RPad(P.paye_st_cd, 2, @ConstSpace)			AS Fld13 -- 13. State of Withholding
		,dbo.func_RPad(P.paye_ssn_tin_num, 9, @ConstSpace)		AS Fld14 -- 14. Payee SSN/TIN
		,dbo.func_RPad(P.paye_ssn_tin_typ_cd, 1, @ConstSpace)	AS Fld15 -- 15. Payee SSN/TIN Type (P=person, B=business)
		,Fld16 =
			CASE 														 
				WHEN AST.lob_cd = 'I' THEN								 -- 16. Indiv LOB: SunCode, Pol#, then S
					@Fld16LiteralValue1
					+ dbo.func_RPad(C.clm_pol_num, 9, @ConstSpace)
					+ @Fld16LiteralValue2
				ELSE													 -- 16. Group LOB: 15-char claim number					
					dbo.func_RPad(C.clm_num, 15, @ConstSpace)
			END
		,@Fld17LiteralValue										AS Fld17 -- 17. Currency Type (US), then blank for TIN2 Notice
		,@Fld18LiteralValue										AS Fld18 -- 18. Unused (27 bytes)
		,dbo.func_LPad(P.paye_clm_int_amt, 11, @ConstZero, 'Y') AS Fld19 -- 19. Box 1 amount (Claim Interest Paid to this Payee)
		,@Fld20LiteralValue										AS Fld20 -- 20. Box 2 amount
		,@Fld21LiteralValue										AS Fld21 -- 21. Box 3 amount
		,dbo.func_LPad(P.paye_wthld_amt, 11, @ConstZero, 'Y')	AS Fld22 -- 22. Box 4 amount (Interest withheld from Claim Interest paid to this Payee)
		,@Fld23LiteralValue										AS Fld23 -- 23. Box 5 amount
		,@Fld24LiteralValue										AS Fld24 -- 24. Box 6 amount
		,@Fld25LiteralValue										AS Fld25
		,@Fld26LiteralValue										AS Fld26
		,@Fld27Unused											AS Fld29
		,@Fld28Unused											AS Fld30
	FROM payee_t					P
	INNER JOIN claim_t				C		ON	P.clm_id = C.clm_id
	INNER JOIN state_t				S		ON	P.paye_st_cd = S.st_cd
	INNER JOIN admin_system_t		AST		ON	C.admn_syst_cd = AST.admn_syst_cd
	INNER JOIN state_rule_t			SR1		ON  P.paye_st_cd = SR1.st_cd AND SR1.lob_cd <> 'I'
	WHERE (P.paye_pmt_dt BETWEEN @paye_pmt_dt_from_date AND @paye_pmt_dt_to_date)
		AND (P.paye_clm_int_amt >= SR1.strl_int_rptg_flr_amt)
		AND (AST.admn_syst_tax_rptg_ind = 'Y')
		AND (P.paye_pmt_dt >= SR1.strl_eff_dt)
		AND (SR1.strl_end_dt IS NULL OR P.paye_pmt_dt <= SR1.strl_end_dt)

	-- Order by ClmPolNum (Fld16) and PayeeFullNm (Fld3)
	ORDER BY 16, 3



	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error happened while selecting the payee data with which to prepare the tax file. From: ' + CONVERT(VARCHAR(255), @paye_pmt_dt_from_date, 121) + ' and To: ' + CONVERT(VARCHAR(255), @paye_pmt_dt_to_date, 121)
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END


END
GO



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON [dbo].proc_tax_file_layout_generate TO AppRoleClaims
GO
GRANT EXECUTE ON [dbo].proc_tax_file_layout_generate TO Support
GO
