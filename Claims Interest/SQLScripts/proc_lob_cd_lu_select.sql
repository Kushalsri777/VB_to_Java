SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_lob_cd_lu_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_lob_cd_lu_select]
GO


/* Date: 01/02/2003
  * Author: Betsy Walker
  * 
  * Purpose: This stored procedure is written with the intent to retrieve the LOB_CD value
  *                 associated with the specified ADMN_SYST_CD, from the ADMIN_SYSTEM_T table.
  * 
  * Logic flow: Perform the select
  *                   Perform error checking
  *
  * Returns: LOB_CD for success
  *               4027, 4028 for failure
  */

CREATE PROCEDURE proc_lob_cd_lu_select (@admn_syst_cd dom_admin_cd,
			@lob_cd dom_lob_cd OUTPUT) WITH RECOMPILE AS

BEGIN

	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	IF NOT EXISTS(SELECT *
			FROM admin_system_t
			WHERE admn_syst_cd = @admn_syst_cd)
	BEGIN
		SELECT @Error_Message = 'Admin System Code: ' + ISNULL(CONVERT(VARCHAR(255), @admn_syst_cd), '<NULL>') + ' does not exist in the ADMIN_SYSTEM_T table.'
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4027
	END

	SELECT @lob_cd = lob_cd
	FROM admin_system_t 
	WHERE admn_syst_cd = @admn_syst_cd

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error happened while selecting the line-of-business code from the ADMIN_SYSTEM_T table for the Admin System Code: ' + ISNULL(CONVERT(VARCHAR(255), @admn_syst_cd), '<NULL>')
		RAISERROR(@Error_Message, 16, 1)
		RETURN 4028
	END

END
GO
setuser
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

GRANT EXECUTE ON dbo.proc_lob_cd_lu_select TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_lob_cd_lu_select TO Support
GO
