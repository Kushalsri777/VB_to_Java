SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IndividualReport_v]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[IndividualReport_v]
GO



CREATE VIEW dbo.IndividualReport_v
AS
/****************************************************************************************
** Created By/Date:		K758 - 05/06/2003
** 
** Purpose: 			Gathers data to be shown on the Individual Report which is printed
**                      from the Insured screen.
**
** Date of Release: 	05/06/2003
** Current Version:		2.0
**
** Used by:				frmInsured in Claims Interest application
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
** 05/06/2003 K758 	   v2.4	    Changed join to PAYEE_T table to be a LEFT OUTER join
**                              rather than INNER JOIN, to allow report to show Insureds
**                              with no Payees (yet). 
** 05/08/2003 K758     v2.4     Added Insured SSN to resultset.
**
****************************************************************************************/
SELECT  C.clm_id
	,C.clm_num
	,AST.admn_syst_dsc
	,PCT.pyco_typ_dsc
	,calcInsuredFullName = convert(varchar(102),
		CASE
			WHEN ISNULL(C.clm_insd_first_nm, '') <> '' THEN C.clm_insd_first_nm + ' ' + C.clm_insd_last_nm
			ELSE C.clm_insd_last_nm
		END)
	,C.clm_insd_dth_dt
	,C.clm_proof_dt
	,C.clm_tot_dthb_pmt_amt
	,C.clm_tot_int_amt
	,C.clm_tot_clm_pd_amt
	,C.clm_tot_wthld_amt
	,calcInsdDthResStCd = convert(char(17),
		CASE
			WHEN (C.CLM_FOR_RES_DTH_IND = 'Y')			THEN '(foreign address)'
			ELSE C.insd_dth_res_st_cd
		END)
	,C.iss_st_cd
    ,C.clm_insd_ssn_num	
	,P.paye_full_nm
	,P.paye_addr_ln1_txt
	,P.paye_addr_ln2_txt
	,P.paye_city_nm_txt
	,P.paye_st_cd
	,P.calc_st_cd
	,calcPayeZipCd = convert(char(10),
		CASE
			WHEN (ISNULL(P.paye_zip4_cd,'') <> '')		THEN P.paye_zip_cd + '-' + P.paye_zip4_cd
			ELSE P.paye_zip_cd
		END)
	,P.paye_care_of_txt
	,P.paye_ssn_tin_num
	,P.paye_clm_int_amt
	,P.paye_clm_pd_amt
	,P.paye_dthb_pmt_amt
	,P.paye_clm_int_rt
	,P.paye_wthld_rt
	,P.paye_wthld_amt
	,P.paye_pmt_dt
	,P.paye_int_days_pd_num
	,calcPayeDfltOvrdInd = convert(char(12),
		CASE
			WHEN (P.PAYE_DFLT_OVRD_IND = 'Y')			THEN '(overridden)'
			ELSE ''
		END)
	
FROM claim_t					C
LEFT OUTER JOIN payee_t			P		ON	C.clm_id = P.clm_id
INNER JOIN admin_system_t		AST		ON	C.admn_syst_cd = AST.admn_syst_cd
INNER JOIN payor_company_type_t	PCT		ON	C.pyco_typ_cd = PCT.pyco_typ_cd




GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

GRANT SELECT ON dbo.IndividualReport_v TO AppRoleClaims
GO
GRANT SELECT ON dbo.IndividualReport_v TO Support
GO
GRANT SELECT ON dbo.IndividualReport_v TO UserAdmin
GO
GRANT SELECT ON dbo.IndividualReport_v TO UserStd
GO