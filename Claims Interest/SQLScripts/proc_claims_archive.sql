SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


/* DBA: DANNY KHOURY
 * Date: 10/21/2002
 *
 * Purpose: This stored procedure is written with the intent to archive the claim_t, and the payee_t tables into the INDClaims1_ar database.
 *                As a requirement for this archival mechanism, this stored procedure keeps the lookup tables for these three tables in the archive database up-to-date with the information
 *                in the active database. This archive process is written with the intent to run once a year on May 1. All the rows for the years before the previous year are archived.
 *                In the process, the stored procedure keeps a log in the archive_log_t table of all the transaction activities that it had completed along with the
 *                number of records that are affected by those activities...Only seven years worth of log information is kept in the archive_log_t table...
 *
 * Lookup tables: current_rate_t
 *                          line_of_business_t
 *                          admin_system_t
 *                          payor_company_t
 *                          interest_rate_rule_t
 *                          interest_date_type_t
 *                          state_t
 *                          state_rule_t
 */

ALTER  PROCEDURE proc_claims_archive WITH RECOMPILE AS

BEGIN

	DECLARE @Error_No INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Stored_Procedure_Name VARCHAR(30)
	DECLARE @Log_Message VARCHAR(255)
	DECLARE @Archive_Cutoff_Datetime DATETIME
	DECLARE @Num_Of_Years_Cutoff SMALLINT

	SELECT @Stored_Procedure_Name = 'proc_claims_archive',
		@Archive_Cutoff_Datetime = GETDATE(),
		@Num_Of_Years_Cutoff = 7

	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			0,
			'Stored procedure started...')

	DELETE INDClaims1_pr..archive_log_t
	WHERE DATEADD(yy, 7, INDClaims1_pr..archive_log_t.Entry_Datetime) <= @Archive_Cutoff_Datetime
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records deleted in table INDClaims1_pr..archive_log_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		SELECT @Log_Message = 'An error happened when deleting table INDClaims1_pr..archive_log_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	BEGIN TRANSACTION Transaction_Lookups

	UPDATE INDClaims1_ar..current_rate_t
	SET INDClaims1_ar..current_rate_t.curr_int_rt = INDClaims1_pr..current_rate_t.curr_int_rt,
		INDClaims1_ar..current_rate_t.curr_rt_end_dt = INDClaims1_pr..current_rate_t.curr_rt_end_dt,
		INDClaims1_ar..current_rate_t.lst_updt_dtm = INDClaims1_pr..current_rate_t.lst_updt_dtm,
		INDClaims1_ar..current_rate_t.lst_updt_user_id = INDClaims1_pr..current_rate_t.lst_updt_user_id
	FROM INDClaims1_pr..current_rate_t
	WHERE INDClaims1_ar..current_rate_t.curr_rt_eff_dt = INDClaims1_pr..current_rate_t.curr_rt_eff_dt

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..current_rate_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..current_rate_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..current_rate_t(curr_rt_eff_dt,
						curr_int_rt,
						curr_rt_end_dt,
						lst_updt_dtm,
						lst_updt_user_id)
		SELECT INDClaims1_pr..current_rate_t.curr_rt_eff_dt,
			INDClaims1_pr..current_rate_t.curr_int_rt,
			INDClaims1_pr..current_rate_t.curr_rt_end_dt,
			INDClaims1_pr..current_rate_t.lst_updt_dtm,
			INDClaims1_pr..current_rate_t.lst_updt_user_id
		FROM INDClaims1_pr..current_rate_t LEFT OUTER JOIN INDClaims1_ar..current_rate_t ON
			INDClaims1_pr..current_rate_t.curr_rt_eff_dt  = INDClaims1_ar..current_rate_t.curr_rt_eff_dt
		WHERE INDClaims1_ar..current_rate_t.curr_rt_eff_dt IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..current_rate_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..current_rate_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	UPDATE INDClaims1_ar..line_of_business_t
	SET INDClaims1_ar..line_of_business_t.lob_dsc = INDClaims1_pr..line_of_business_t.lob_dsc,
		INDClaims1_ar..line_of_business_t.lst_updt_dtm = INDClaims1_pr..line_of_business_t.lst_updt_dtm,
		INDClaims1_ar..line_of_business_t.lst_updt_user_id = INDClaims1_pr..line_of_business_t.lst_updt_user_id
	FROM INDClaims1_pr..line_of_business_t
	WHERE INDClaims1_ar..line_of_business_t.lob_cd = INDClaims1_pr..line_of_business_t.lob_cd

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..line_of_business_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..line_of_business_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..line_of_business_t(lob_cd,
						lob_dsc,
						lst_updt_dtm,
						lst_updt_user_id)
		SELECT INDClaims1_pr..line_of_business_t.lob_cd,
			INDClaims1_pr..line_of_business_t.lob_dsc,
			INDClaims1_pr..line_of_business_t.lst_updt_dtm,
			INDClaims1_pr..line_of_business_t.lst_updt_user_id
		FROM INDClaims1_pr..line_of_business_t LEFT OUTER JOIN INDClaims1_ar..line_of_business_t ON
			INDClaims1_pr..line_of_business_t.lob_cd  = INDClaims1_ar..line_of_business_t.lob_cd
		WHERE INDClaims1_ar..line_of_business_t.lob_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..line_of_business_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..line_of_business_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)
/*
	UPDATE INDClaims1_ar..tax_file_layout_t
	SET INDClaims1_ar..tax_file_layout_t.tfcol_dsc = INDClaims1_pr..tax_file_layout_t.tfcol_dsc,
		INDClaims1_ar..tax_file_layout_t.tfcol_src_nm = INDClaims1_pr..tax_file_layout_t.tfcol_src_nm,
		INDClaims1_ar..tax_file_layout_t.tfcol_litr_vlu_txt = INDClaims1_pr..tax_file_layout_t.tfcol_litr_vlu_txt,
		INDClaims1_ar..tax_file_layout_t.tfcol_lgth_num = INDClaims1_pr..tax_file_layout_t.tfcol_lgth_num,
		INDClaims1_ar..tax_file_layout_t.tfcol_pad_ldg_zero_ind = INDClaims1_pr..tax_file_layout_t.tfcol_pad_ldg_zero_ind,
		INDClaims1_ar..tax_file_layout_t.tfcol_pad_trlg_sp_ind = INDClaims1_pr..tax_file_layout_t.tfcol_pad_trlg_sp_ind,
		INDClaims1_ar..tax_file_layout_t.lst_updt_dtm = INDClaims1_pr..tax_file_layout_t.lst_updt_dtm,
		INDClaims1_ar..tax_file_layout_t.lst_updt_user_id = INDClaims1_pr..tax_file_layout_t.lst_updt_user_id
	FROM INDClaims1_pr..tax_file_layout_t
	WHERE INDClaims1_ar..tax_file_layout_t.lob_cd = INDClaims1_pr..tax_file_layout_t.lob_cd AND
		INDClaims1_ar..tax_file_layout_t.tfcol_seq_num = INDClaims1_pr..tax_file_layout_t.tfcol_seq_num

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..tax_file_layout_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..tax_file_layout_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..tax_file_layout_t(lob_cd,
						tfcol_seq_num,
						tfcol_dsc,
						tfcol_src_nm,
						tfcol_litr_vlu_txt,
						tfcol_lgth_num,
						tfcol_pad_ldg_zero_ind,
						tfcol_pad_trlg_sp_ind,
						lst_updt_dtm,
						lst_updt_user_id)
		SELECT INDClaims1_pr..tax_file_layout_t.lob_cd,
			INDClaims1_pr..tax_file_layout_t.tfcol_seq_num,
			INDClaims1_pr..tax_file_layout_t.tfcol_dsc,
			INDClaims1_pr..tax_file_layout_t.tfcol_src_nm,
			INDClaims1_pr..tax_file_layout_t.tfcol_litr_vlu_txt,
			INDClaims1_pr..tax_file_layout_t.tfcol_lgth_num,
			INDClaims1_pr..tax_file_layout_t.tfcol_pad_ldg_zero_ind,
			INDClaims1_pr..tax_file_layout_t.tfcol_pad_trlg_sp_ind,
			INDClaims1_pr..tax_file_layout_t.lst_updt_dtm,
			INDClaims1_pr..tax_file_layout_t.lst_updt_user_id
		FROM INDClaims1_pr..tax_file_layout_t LEFT OUTER JOIN INDClaims1_ar..tax_file_layout_t ON
			INDClaims1_pr..tax_file_layout_t.lob_cd  = INDClaims1_ar..tax_file_layout_t.lob_cd AND
			INDClaims1_pr..tax_file_layout_t.tfcol_seq_num  = INDClaims1_ar..tax_file_layout_t.tfcol_seq_num
		WHERE INDClaims1_ar..tax_file_layout_t.lob_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..tax_file_layout_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..tax_file_layout_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)
*/
	UPDATE INDClaims1_ar..admin_system_t
	SET INDClaims1_ar..admin_system_t.lob_cd = INDClaims1_pr..admin_system_t.lob_cd,
		INDClaims1_ar..admin_system_t.admn_syst_dsc = INDClaims1_pr..admin_system_t.admn_syst_dsc,
		INDClaims1_ar..admin_system_t.lst_updt_dtm = INDClaims1_pr..admin_system_t.lst_updt_dtm,
		INDClaims1_ar..admin_system_t.lst_updt_user_id = INDClaims1_pr..admin_system_t.lst_updt_user_id
	FROM INDClaims1_pr..admin_system_t
	WHERE INDClaims1_ar..admin_system_t.admn_syst_cd = INDClaims1_pr..admin_system_t.admn_syst_cd

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table indppvul2_ar..admin_system_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table indppvul2_ar..admin_system_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..admin_system_t(admn_syst_cd,
						lob_cd,
						admn_syst_dsc,
						lst_updt_dtm,
						lst_updt_user_id)
		SELECT INDClaims1_pr..admin_system_t.admn_syst_cd,
			INDClaims1_pr..admin_system_t.lob_cd,
			INDClaims1_pr..admin_system_t.admn_syst_dsc,
			INDClaims1_pr..admin_system_t.lst_updt_dtm,
			INDClaims1_pr..admin_system_t.lst_updt_user_id
		FROM INDClaims1_pr..admin_system_t LEFT OUTER JOIN INDClaims1_ar..admin_system_t ON
			INDClaims1_pr..admin_system_t.admn_syst_cd  = INDClaims1_ar..admin_system_t.admn_syst_cd
		WHERE INDClaims1_ar..admin_system_t.admn_syst_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..admin_system_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..admin_system_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	UPDATE INDClaims1_ar..payor_company_type_t
	SET INDClaims1_ar..payor_company_type_t.pyco_typ_dsc = INDClaims1_pr..payor_company_type_t.pyco_typ_dsc,
		INDClaims1_ar..payor_company_type_t.lst_updt_dtm = INDClaims1_pr..payor_company_type_t.lst_updt_dtm,
		INDClaims1_ar..payor_company_type_t.lst_updt_user_id = INDClaims1_pr..payor_company_type_t.lst_updt_user_id
	FROM INDClaims1_pr..payor_company_type_t
	WHERE INDClaims1_ar..payor_company_type_t.pyco_typ_cd = INDClaims1_pr..payor_company_type_t.pyco_typ_cd

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..payor_company_type_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..payor_company_type_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..payor_company_type_t(pyco_typ_cd,
						pyco_typ_dsc,
						lst_updt_dtm,
						lst_updt_user_id)
		SELECT INDClaims1_pr..payor_company_type_t.pyco_typ_cd,
			INDClaims1_pr..payor_company_type_t.pyco_typ_dsc,
			INDClaims1_pr..payor_company_type_t.lst_updt_dtm,
			INDClaims1_pr..payor_company_type_t.lst_updt_user_id
		FROM INDClaims1_pr..payor_company_type_t LEFT OUTER JOIN INDClaims1_ar..payor_company_type_t ON
			INDClaims1_pr..payor_company_type_t.pyco_typ_cd  = INDClaims1_ar..payor_company_type_t.pyco_typ_cd
		WHERE INDClaims1_ar..payor_company_type_t.pyco_typ_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..payor_company_type_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..payor_company_type_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	UPDATE INDClaims1_ar..interest_rate_rule_t
	SET INDClaims1_ar..interest_rate_rule_t.irule_dsc = INDClaims1_pr..interest_rate_rule_t.irule_dsc,
		INDClaims1_ar..interest_rate_rule_t.lst_updt_dtm = INDClaims1_pr..interest_rate_rule_t.lst_updt_dtm,
		INDClaims1_ar..interest_rate_rule_t.lst_updt_user_id = INDClaims1_pr..interest_rate_rule_t.lst_updt_user_id
	FROM INDClaims1_pr..interest_rate_rule_t
	WHERE INDClaims1_ar..interest_rate_rule_t.irule_cd = INDClaims1_pr..interest_rate_rule_t.irule_cd

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..interest_rate_rule_t = ' + CONVERT(VARCHAR, @Row_Count)
	END

	ELSE

	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..interest_rate_rule_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..interest_rate_rule_t(irule_cd,
						irule_dsc,
						lst_updt_dtm,
						lst_updt_user_id)
		SELECT INDClaims1_pr..interest_rate_rule_t.irule_cd,
			INDClaims1_pr..interest_rate_rule_t.irule_dsc,
			INDClaims1_pr..interest_rate_rule_t.lst_updt_dtm,
			INDClaims1_pr..interest_rate_rule_t.lst_updt_user_id
		FROM INDClaims1_pr..interest_rate_rule_t LEFT OUTER JOIN INDClaims1_ar..interest_rate_rule_t ON
			INDClaims1_pr..interest_rate_rule_t.irule_cd  = INDClaims1_ar..interest_rate_rule_t.irule_cd
		WHERE INDClaims1_ar..interest_rate_rule_t.irule_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..interest_rate_rule_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..interest_rate_rule_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	UPDATE INDClaims1_ar..interest_date_type_t
	SET INDClaims1_ar..interest_date_type_t.idtyp_dsc = INDClaims1_pr..interest_date_type_t.idtyp_dsc,
		INDClaims1_ar..interest_date_type_t.lst_updt_dtm = INDClaims1_pr..interest_date_type_t.lst_updt_dtm,
		INDClaims1_ar..interest_date_type_t.lst_updt_user_id = INDClaims1_pr..interest_date_type_t.lst_updt_user_id
	FROM INDClaims1_pr..interest_date_type_t
	WHERE INDClaims1_ar..interest_date_type_t.idtyp_cd = INDClaims1_pr..interest_date_type_t.idtyp_cd

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..interest_date_type_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..interest_date_type_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..interest_date_type_t(idtyp_cd,
						idtyp_dsc,
						lst_updt_dtm,
						lst_updt_user_id)
		SELECT INDClaims1_pr..interest_date_type_t.idtyp_cd,
			INDClaims1_pr..interest_date_type_t.idtyp_dsc,
			INDClaims1_pr..interest_date_type_t.lst_updt_dtm,
			INDClaims1_pr..interest_date_type_t.lst_updt_user_id
		FROM INDClaims1_pr..interest_date_type_t LEFT OUTER JOIN INDClaims1_ar..interest_date_type_t ON
			INDClaims1_pr..interest_date_type_t.idtyp_cd  = INDClaims1_ar..interest_date_type_t.idtyp_cd
		WHERE INDClaims1_ar..interest_date_type_t.idtyp_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..interest_date_type_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..interest_date_type_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	UPDATE INDClaims1_ar..state_t
	SET INDClaims1_ar..state_t.st_domcl_cd = INDClaims1_pr..state_t.st_domcl_cd,
		INDClaims1_ar..state_t.st_dsc = INDClaims1_pr..state_t.st_dsc,
		INDClaims1_ar..state_t.lst_updt_dtm = INDClaims1_pr..state_t.lst_updt_dtm,
		INDClaims1_ar..state_t.lst_updt_user_id = INDClaims1_pr..state_t.lst_updt_user_id
	FROM INDClaims1_pr..state_t
	WHERE INDClaims1_ar..state_t.st_cd = INDClaims1_pr..state_t.st_cd

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..state_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..state_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..state_t(st_cd,
				st_domcl_cd,
				st_dsc,
				lst_updt_dtm,
				lst_updt_user_id)
			SELECT INDClaims1_pr..state_t.st_cd,
				INDClaims1_pr..state_t.st_domcl_cd,
				INDClaims1_pr..state_t.st_dsc,
				INDClaims1_pr..state_t.lst_updt_dtm,
				INDClaims1_pr..state_t.lst_updt_user_id
			FROM INDClaims1_pr..state_t LEFT OUTER JOIN INDClaims1_ar..state_t ON
				INDClaims1_pr..state_t.st_cd  = INDClaims1_ar..state_t.st_cd
			WHERE INDClaims1_ar..state_t.st_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..state_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..state_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	UPDATE INDClaims1_ar..state_rule_t
	SET INDClaims1_ar..state_rule_t.calc_idtyp_cd = INDClaims1_pr..state_rule_t.calc_idtyp_cd,
		INDClaims1_ar..state_rule_t.reqd_idtyp_cd = INDClaims1_pr..state_rule_t.reqd_idtyp_cd,
		INDClaims1_ar..state_rule_t.irule_cd = INDClaims1_pr..state_rule_t.irule_cd,
		INDClaims1_ar..state_rule_t.strl_end_dt = INDClaims1_pr..state_rule_t.strl_end_dt,
		INDClaims1_ar..state_rule_t.strl_int_rptg_flr_amt = INDClaims1_pr..state_rule_t.strl_int_rptg_flr_amt,
		INDClaims1_ar..state_rule_t.strl_int_calc_ofst_num = INDClaims1_pr..state_rule_t.strl_int_calc_ofst_num,
		INDClaims1_ar..state_rule_t.strl_int_reqd_ofst_num = INDClaims1_pr..state_rule_t.strl_int_reqd_ofst_num,
		INDClaims1_ar..state_rule_t.strl_int_rule_amt = INDClaims1_pr..state_rule_t.strl_int_rule_amt,
		INDClaims1_ar..state_rule_t.strl_spcl_instr_txt = INDClaims1_pr..state_rule_t.strl_spcl_instr_txt,
		INDClaims1_ar..state_rule_t.lst_updt_dtm = INDClaims1_pr..state_rule_t.lst_updt_dtm,
		INDClaims1_ar..state_rule_t.lst_updt_user_id = INDClaims1_pr..state_rule_t.lst_updt_user_id
	FROM INDClaims1_pr..state_rule_t
	WHERE INDClaims1_ar..state_rule_t.lob_cd = INDClaims1_pr..state_rule_t.lob_cd AND
		INDClaims1_ar..state_rule_t.st_cd = INDClaims1_pr..state_rule_t.st_cd AND
		INDClaims1_ar..state_rule_t.strl_eff_dt = INDClaims1_pr..state_rule_t.strl_eff_dt

	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in table INDClaims1_ar..state_rule_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when updating table INDClaims1_ar..state_rule_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO INDClaims1_ar..state_rule_t(lob_cd,
					st_cd,
					strl_eff_dt,
					calc_idtyp_cd,
					reqd_idtyp_cd,
					irule_cd,
					strl_end_dt,
					strl_int_rptg_flr_amt,
					strl_int_calc_ofst_num,
					strl_int_reqd_ofst_num,
					strl_int_rule_amt,
					strl_spcl_instr_txt,
					lst_updt_dtm,
					lst_updt_user_id)
		SELECT INDClaims1_pr..state_rule_t.lob_cd,
			INDClaims1_pr..state_rule_t.st_cd,
			INDClaims1_pr..state_rule_t.strl_eff_dt,
			INDClaims1_pr..state_rule_t.calc_idtyp_cd,
			INDClaims1_pr..state_rule_t.reqd_idtyp_cd,
			INDClaims1_pr..state_rule_t.irule_cd,
			INDClaims1_pr..state_rule_t.strl_end_dt,
			INDClaims1_pr..state_rule_t.strl_int_rptg_flr_amt,
			INDClaims1_pr..state_rule_t.strl_int_calc_ofst_num,
			INDClaims1_pr..state_rule_t.strl_int_reqd_ofst_num,
			INDClaims1_pr..state_rule_t.strl_int_rule_amt,
			INDClaims1_pr..state_rule_t.strl_spcl_instr_txt,
			INDClaims1_pr..state_rule_t.lst_updt_dtm,
			INDClaims1_pr..state_rule_t.lst_updt_user_id
		FROM INDClaims1_pr..state_rule_t LEFT OUTER JOIN INDClaims1_ar..state_rule_t ON
			INDClaims1_pr..state_rule_t.lob_cd  = INDClaims1_ar..state_rule_t.lob_cd AND
			INDClaims1_pr..state_rule_t.st_cd  = INDClaims1_ar..state_rule_t.st_cd AND
			INDClaims1_pr..state_rule_t.strl_eff_dt  = INDClaims1_ar..state_rule_t.strl_eff_dt
		WHERE INDClaims1_ar..state_rule_t.lob_cd IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..state_rule_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Lookups
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..state_rule_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	COMMIT TRANSACTION Transaction_Lookups

	BEGIN TRANSACTION Transaction_Claim_Payee

	CREATE TABLE #claims_archive_temp_t(clm_id INTEGER)
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Created temp table #claims_archive_temp_t successfully'
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened while creating temp table #claims_archive_temp_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO #claims_archive_temp_t(clm_id)
		SELECT DISTINCT clm_id
		FROM INDClaims1_pr..payee_t
		WHERE YEAR(INDClaims1_pr..payee_t.paye_pmt_dt) < (YEAR(@Archive_Cutoff_Datetime) - 1)
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into temp table #claims_archive_temp_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when inserting into temp table #claims_archive_temp_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	DELETE #claims_archive_temp_t
	FROM INDClaims1_pr..payee_t
	WHERE #claims_archive_temp_t.clm_id = INDClaims1_pr..payee_t.clm_id AND
		YEAR(INDClaims1_pr..payee_t.paye_pmt_dt) >= (YEAR(@Archive_Cutoff_Datetime) - 1)
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records deleted from temp table #claims_archive_temp_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened while deleting from temp table #claims_archive_temp_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	UPDATE INDClaims1_ar..claim_t
	SET INDClaims1_ar..claim_t.clm_pol_num = INDClaims1_pr..claim_t.clm_pol_num,
		INDClaims1_ar..claim_t.clm_proof_dt = INDClaims1_pr..claim_t.clm_proof_dt,
		INDClaims1_ar..claim_t.clm_tot_dthb_pmt_amt = INDClaims1_ar..claim_t.clm_tot_dthb_pmt_amt + INDClaims1_pr..claim_t.clm_tot_dthb_pmt_amt,
		INDClaims1_ar..claim_t.clm_tot_int_amt = INDClaims1_ar..claim_t.clm_tot_int_amt + INDClaims1_pr..claim_t.clm_tot_int_amt,
		INDClaims1_ar..claim_t.clm_tot_clm_pd_amt = INDClaims1_ar..claim_t.clm_tot_clm_pd_amt + INDClaims1_pr..claim_t.clm_tot_clm_pd_amt,
		INDClaims1_ar..claim_t.clm_tot_wthld_amt = INDClaims1_ar..claim_t.clm_tot_wthld_amt + INDClaims1_pr..claim_t.clm_tot_wthld_amt,
		INDClaims1_ar..claim_t.clm_insd_last_nm = INDClaims1_pr..claim_t.clm_insd_last_nm,
		INDClaims1_ar..claim_t.clm_insd_first_nm = INDClaims1_pr..claim_t.clm_insd_first_nm,
		INDClaims1_ar..claim_t.clm_insd_dth_dt = INDClaims1_pr..claim_t.clm_insd_dth_dt,
		INDClaims1_ar..claim_t.clm_insd_ssn_num = INDClaims1_pr..claim_t.clm_insd_ssn_num,
		INDClaims1_ar..claim_t.iss_st_cd = INDClaims1_pr..claim_t.iss_st_cd,
		INDClaims1_ar..claim_t.insd_dth_res_st_cd = INDClaims1_pr..claim_t.insd_dth_res_st_cd,
		INDClaims1_ar..claim_t.admn_syst_cd = INDClaims1_pr..claim_t.admn_syst_cd,
		INDClaims1_ar..claim_t.pyco_typ_cd = INDClaims1_pr..claim_t.pyco_typ_cd,
		INDClaims1_ar..claim_t.clm_for_res_dth_ind = INDClaims1_pr..claim_t.clm_for_res_dth_ind,
		INDClaims1_ar..claim_t.lst_updt_dtm = INDClaims1_pr..claim_t.lst_updt_dtm,
		INDClaims1_ar..claim_t.lst_updt_user_id = INDClaims1_pr..claim_t.lst_updt_user_id
	FROM INDClaims1_pr..claim_t INNER JOIN #claims_archive_temp_t ON
		INDClaims1_pr..claim_t.clm_id = #claims_archive_temp_t.clm_id
	WHERE INDClaims1_ar..claim_t.clm_num = INDClaims1_pr..claim_t.clm_num
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records updated in the table INDClaims1_ar..claim_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened while updating the table INDClaims1_ar..claim_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	SET IDENTITY_INSERT INDClaims1_ar..claim_t ON
	INSERT INTO INDClaims1_ar..claim_t(clm_id,
					clm_num,
					clm_pol_num,
					clm_proof_dt,
					clm_tot_dthb_pmt_amt,
					clm_tot_int_amt,
					clm_tot_clm_pd_amt,
					clm_tot_wthld_amt,
					clm_insd_last_nm,
					clm_insd_first_nm,
					clm_insd_dth_dt,
					clm_insd_ssn_num,
					iss_st_cd,
					insd_dth_res_st_cd,
					admn_syst_cd,
					pyco_typ_cd,
					clm_for_res_dth_ind,
					lst_updt_dtm,
					lst_updt_user_id)
			SELECT INDClaims1_pr..claim_t.clm_id,
				INDClaims1_pr..claim_t.clm_num,
				INDClaims1_pr..claim_t.clm_pol_num,
				INDClaims1_pr..claim_t.clm_proof_dt,
				INDClaims1_pr..claim_t.clm_tot_dthb_pmt_amt,
				INDClaims1_pr..claim_t.clm_tot_int_amt,
				INDClaims1_pr..claim_t.clm_tot_clm_pd_amt,
				INDClaims1_pr..claim_t.clm_tot_wthld_amt,
				INDClaims1_pr..claim_t.clm_insd_last_nm,
				INDClaims1_pr..claim_t.clm_insd_first_nm,
				INDClaims1_pr..claim_t.clm_insd_dth_dt,
				INDClaims1_pr..claim_t.clm_insd_ssn_num,
				INDClaims1_pr..claim_t.iss_st_cd,
				INDClaims1_pr..claim_t.insd_dth_res_st_cd,
				INDClaims1_pr..claim_t.admn_syst_cd,
				INDClaims1_pr..claim_t.pyco_typ_cd,
				INDClaims1_pr..claim_t.clm_for_res_dth_ind,
				INDClaims1_pr..claim_t.lst_updt_dtm,
				INDClaims1_pr..claim_t.lst_updt_user_id
			FROM INDClaims1_pr..claim_t INNER JOIN #claims_archive_temp_t ON
				INDClaims1_pr..claim_t.clm_id = #claims_archive_temp_t.clm_id
			LEFT OUTER JOIN INDClaims1_ar..claim_t ON
				INDClaims1_pr..claim_t.clm_num = INDClaims1_ar..claim_t.clm_num
			WHERE INDClaims1_ar..claim_t.clm_num IS NULL
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..claim_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..claim_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)
	SET IDENTITY_INSERT INDClaims1_ar..claim_t OFF

	SET IDENTITY_INSERT INDClaims1_ar..payee_t ON
	INSERT INTO INDClaims1_ar..payee_t(paye_id,
					clm_id,
					calc_st_cd,
					paye_full_nm,
					paye_care_of_txt,
					paye_addr_ln1_txt,
					paye_addr_ln2_txt,
					paye_city_nm_txt,
					paye_st_cd,
					paye_zip_cd,
					paye_zip4_cd,
					paye_ssn_tin_num,
					paye_ssn_tin_typ_cd,
					paye_clm_int_amt,
					paye_clm_pd_amt,
					paye_dthb_pmt_amt,
					paye_wthld_amt,
					paye_clm_int_rt,
					paye_wthld_rt,
					paye_int_days_pd_num,
					paye_pmt_dt,
					paye_dflt_ovrd_ind,
					lst_updt_dtm,
					lst_updt_user_id)
			SELECT INDClaims1_pr..payee_t.paye_id,
				ISNULL(INDClaims1_ar..claim_t.clm_id, INDClaims1_pr..claim_t.clm_id) clm_id,
				INDClaims1_pr..payee_t.calc_st_cd,
				INDClaims1_pr..payee_t.paye_full_nm,
				INDClaims1_pr..payee_t.paye_care_of_txt,
				INDClaims1_pr..payee_t.paye_addr_ln1_txt,
				INDClaims1_pr..payee_t.paye_addr_ln2_txt,
				INDClaims1_pr..payee_t.paye_city_nm_txt,
				INDClaims1_pr..payee_t.paye_st_cd,
				INDClaims1_pr..payee_t.paye_zip_cd,
				INDClaims1_pr..payee_t.paye_zip4_cd,
				INDClaims1_pr..payee_t.paye_ssn_tin_num,
				INDClaims1_pr..payee_t.paye_ssn_tin_typ_cd,
				INDClaims1_pr..payee_t.paye_clm_int_amt,
				INDClaims1_pr..payee_t.paye_clm_pd_amt,
				INDClaims1_pr..payee_t.paye_dthb_pmt_amt,
				INDClaims1_pr..payee_t.paye_wthld_amt,
				INDClaims1_pr..payee_t.paye_clm_int_rt,
				INDClaims1_pr..payee_t.paye_wthld_rt,
				INDClaims1_pr..payee_t.paye_int_days_pd_num,
				INDClaims1_pr..payee_t.paye_pmt_dt,
				INDClaims1_pr..payee_t.paye_dflt_ovrd_ind,
				INDClaims1_pr..payee_t.lst_updt_dtm,
				INDClaims1_pr..payee_t.lst_updt_user_id
			FROM INDClaims1_pr..claim_t INNER JOIN #claims_archive_temp_t ON
				INDClaims1_pr..claim_t.clm_id = #claims_archive_temp_t.clm_id
			INNER JOIN INDClaims1_pr..payee_t ON
				INDClaims1_pr..claim_t.clm_id = INDClaims1_pr..payee_t.clm_id
			LEFT OUTER JOIN INDClaims1_ar..claim_t ON
				INDClaims1_pr..claim_t.clm_num = INDClaims1_ar..claim_t.clm_num
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into table INDClaims1_ar..payee_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when inserting into table INDClaims1_ar..payee_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)
	SET IDENTITY_INSERT INDClaims1_ar..payee_t OFF

	DELETE INDClaims1_pr..payee_t
	FROM #claims_archive_temp_t
	WHERE INDClaims1_pr..payee_t.clm_id = #claims_archive_temp_t.clm_id
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records deleted from table INDClaims1_pr..payee_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when deleting from table INDClaims1_pr..payee_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	DELETE INDClaims1_pr..claim_t
	FROM #claims_archive_temp_t
	WHERE INDClaims1_pr..claim_t.clm_id = #claims_archive_temp_t.clm_id
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records deleted from table INDClaims1_pr..claim_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when deleting from table INDClaims1_pr..claim_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	DELETE #claims_archive_temp_t
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records deleted from temp table #claims_archive_temp_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened while deleting from temp table #claims_archive_temp_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	INSERT INTO #claims_archive_temp_t(clm_id)
		SELECT DISTINCT clm_id
		FROM INDClaims1_ar..payee_t
		WHERE YEAR(INDClaims1_ar..payee_t.paye_pmt_dt) <= YEAR(@Archive_Cutoff_Datetime) - @Num_Of_Years_Cutoff
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records inserted into temp table #claims_archive_temp_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when inserting into temp table #claims_archive_temp_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	DELETE #claims_archive_temp_t
	FROM INDClaims1_ar..payee_t
	WHERE #claims_archive_temp_t.clm_id = INDClaims1_ar..payee_t.clm_id AND
		YEAR(INDClaims1_ar..payee_t.paye_pmt_dt) > YEAR(@Archive_Cutoff_Datetime) - @Num_Of_Years_Cutoff
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of records deleted from temp table #claims_archive_temp_t = ' + CONVERT(VARCHAR, @Row_Count)
	END

	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened while deleting from temp table #claims_archive_temp_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	DELETE INDClaims1_ar..payee_t
	FROM #claims_archive_temp_t
	WHERE INDClaims1_ar..payee_t.clm_id = #claims_archive_temp_t.clm_id
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of quarterly records deleted from table INDClaims1_ar..payee_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when deleting quarterly records from table INDClaims1_ar..payee_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	DELETE INDClaims1_ar..claim_t
	FROM #claims_archive_temp_t
	WHERE INDClaims1_ar..claim_t.clm_id = #claims_archive_temp_t.clm_id
	SELECT @Error_No = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Number of quarterly records deleted from table INDClaims1_ar..claim_t = ' + CONVERT(VARCHAR, @Row_Count)
	END
	ELSE
	BEGIN
		ROLLBACK TRANSACTION Transaction_Claim_Payee
		SELECT @Log_Message = 'An error happened when deleting quarterly records from table INDClaims1_ar..claim_t'
		RAISERROR(@Log_Message, 16, 1)
		GOTO Exit_Stored_Procedure
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

	COMMIT TRANSACTION Transaction_Claim_Payee

Exit_Stored_Procedure:

	If @Error_No = 0
	BEGIN
		SELECT @Log_Message = 'Stored procedure completed successfully...'
	END
	ELSE
	BEGIN
		SELECT @Log_Message = 'Stored procedure completed unsuccessfully...'
	END
	INSERT INTO archive_log_t(entry_datetime,
				stored_procedure_name,
				error_number,
				message)
		VALUES(GETDATE(),
			@Stored_Procedure_Name,
			@Error_No,
			@Log_Message)

END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

