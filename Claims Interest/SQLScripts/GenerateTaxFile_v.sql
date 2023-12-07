SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GenerateTaxFile_v]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[GenerateTaxFile_v]
GO


CREATE VIEW dbo.GenerateTaxFile_v
AS
SELECT TOP 100 lob_cd
	,tfcol_seq_num
	,tfcol_src_nm
	,tfcol_litr_vlu_txt
	,tfcol_lgth_num
	,tfcol_pad_ldg_zero_ind
	,tfcol_pad_trlg_sp_ind
FROM tax_file_layout_t
ORDER BY lob_cd, tfcol_seq_num


GO

SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS ON
GO


GRANT SELECT ON dbo.GenerateTaxFile_v TO AppRoleClaims
GO
GRANT SELECT ON dbo.GenerateTaxFile_v TO Support
GO
GRANT SELECT ON dbo.GenerateTaxFile_v TO UserAdmin
GO
GRANT SELECT ON dbo.GenerateTaxFile_v TO UserStd
GO