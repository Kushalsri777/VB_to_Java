SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StateInterestReport_v]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[StateInterestReport_v]
GO


CREATE VIEW dbo.StateInterestReport_v
AS
SELECT  C.clm_num
	,P.paye_full_nm
	,P.paye_st_cd
	,P.paye_clm_int_amt
	,P.paye_clm_pd_amt
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


GRANT SELECT ON dbo.StateInterestReport_v TO AppRoleClaims
GO
GRANT SELECT ON dbo.StateInterestReport_v TO Support
GO
GRANT SELECT ON dbo.StateInterestReport_v TO UserAdmin
GO
GRANT SELECT ON dbo.StateInterestReport_v TO UserStd
GO