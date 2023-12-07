SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CustomClaimPaymentReport_v]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[CustomClaimPaymentReport_v]
GO

CREATE VIEW dbo.CustomClaimPaymentReport_v
AS
SELECT  C.clm_num
	,calcInsuredFullName = convert(varchar(102),
		CASE
			WHEN ISNULL(C.clm_insd_first_nm, '') <> '' THEN C.clm_insd_first_nm + ' ' + C.clm_insd_last_nm
			ELSE C.clm_insd_last_nm
		END)
	,UPPER(C.lst_updt_user_id) As lst_upd_user_id
	,P.paye_full_nm
	,P.paye_clm_int_amt
	,P.paye_clm_pd_amt
	,P.paye_wthld_amt
	,P.paye_pmt_dt
	,AST.lob_cd
FROM		claim_t				C
INNER JOIN 	payee_t				P	ON C.clm_id = P.clm_id
INNER JOIN dbo.admin_system_t 	AST	ON C.admn_syst_cd = AST.admn_syst_cd


GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

GRANT SELECT ON dbo.CustomClaimPaymentReport_v TO AppRoleClaims
GO
GRANT SELECT ON dbo.CustomClaimPaymentReport_v TO Support
GO
GRANT SELECT ON dbo.CustomClaimPaymentReport_v TO UserAdmin
GO
GRANT SELECT ON dbo.CustomClaimPaymentReport_v TO UserStd
GO