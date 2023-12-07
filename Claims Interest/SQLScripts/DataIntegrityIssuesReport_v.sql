SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DataIntegrityIssuesReport_v]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[DataIntegrityIssuesReport_v]
GO

/****************************************************************************************
** Created By/Date:		dbo - 06/17/2003
** 
** Purpose: 			Identifies PAYEE_T rows that could result in incorrect or missing
**                      1099-INT reporting.
**
** Date of Release: 	06/30/2003
** Current Version:		1
**
** Used by:				Payee screen
**
** =================
** Additional Notes
** =================
** Since a given policy will only appear once on the report, the logic for setting calcReason
** should be done in DESCENDING Priority order, e.g., most important reasons done ahead of 
** less important reasons, as identified by the business area.
**
** =================
** Revision history
** =================
**
** Date       Author   Tag      Purpose
** ---------- -------  -------- --------------------------------------------------------
** 07/17/03   K758     Bug2456  Enhanced to list Payees whose # of Days of Interest To 
**                              Be Paid --or-- its Claims Interest Amount is negative.
**
****************************************************************************************/
CREATE VIEW dbo.DataIntegrityIssuesReport_v
AS
SELECT  C.clm_num
	,P.paye_full_nm
	,C.clm_insd_dth_dt
	,P.paye_pmt_dt
	,P.paye_clm_int_rt
	,UPPER(C.lst_updt_user_id) As lst_upd_user_id
	,calcReason = CONVERT(VARCHAR(100), 
		CASE
			WHEN ((P.paye_int_days_pd_num < 0) OR (P.paye_clm_int_amt < 0)) 
				THEN 'Claims Interest Amount or Days of Interest Paid is negative. Verify the calculation is correct.'

			WHEN (P.paye_int_days_pd_num > CONVERT(INTEGER, DATEDIFF(day, C.clm_insd_dth_dt, P.paye_pmt_dt)))
				THEN 'Days of Interest paid(' + CONVERT(VARCHAR(60),P.paye_int_days_pd_num) + 
					 ') exceeds # of days between Payment and Death(' + CONVERT(VARCHAR(60), DATEDIFF(day, C.clm_insd_dth_dt, P.paye_pmt_dt)) + ').'
			
			WHEN (P.paye_pmt_dt > DATEADD (d, 5, GetDate()))
				THEN 'Payment Date is more than 5 days into the future.'

			WHEN ((P.paye_clm_int_rt > 12.00000) AND (P.paye_clm_int_rt <> 18.00000) AND (P.calc_st_cd <> 'ME'))
				THEN 'Only Maine should have an interest rate in excess of 12%.'

			WHEN (P.paye_zip_cd IS NULL) OR (P.paye_zip_cd = '00000')
				THEN 'The Zip Code must be supplied.'

			WHEN (P.paye_pmt_dt < C.clm_insd_dth_dt)
				THEN 'The Date Of Payment is prior to Date Of Death.'

			ELSE ''
		END)
FROM claim_t            C
INNER JOIN payee_t      P   ON C.clm_id = P.clm_id



GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

GRANT SELECT ON dbo.DataIntegrityIssuesReport_v TO AppRoleClaims
GO
GRANT SELECT ON dbo.DataIntegrityIssuesReport_v TO Support
GO
GRANT SELECT ON dbo.DataIntegrityIssuesReport_v TO UserAdmin
GO
GRANT SELECT ON dbo.DataIntegrityIssuesReport_v TO UserStd
GO