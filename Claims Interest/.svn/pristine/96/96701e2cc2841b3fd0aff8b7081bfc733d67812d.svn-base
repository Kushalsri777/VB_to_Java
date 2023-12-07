SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[proc_admin_system_lu_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[proc_admin_system_lu_select]
GO


/* Date: 01/02/2003
  * Author: Betsy Walker
  * 
  * Purpose: This stored procedure is written with the intent to select the certain columns from the ADMIN_SYSTEM_T table, e.g., to
  *                 populate a combobox.
  * 
  * Logic flow: Perform the select
  *                   Perform error checking
  *
  * Return: number of rows for success
  *             4028 for failure
  */

CREATE PROCEDURE proc_admin_system_lu_select WITH RECOMPILE AS

BEGIN

	DECLARE @Error_Number INTEGER
	DECLARE @Row_Count INTEGER
	DECLARE @Error_Message VARCHAR(255)

	SELECT admn_syst_dsc, admn_syst_cd
	FROM ADMIN_SYSTEM_T
	ORDER BY admn_syst_dsc

	SELECT @Error_Number = @@ERROR,
		@Row_Count = @@ROWCOUNT
	IF @Error_Number = 0
	BEGIN
		RETURN @Row_Count
	END
	ELSE
	BEGIN
		SELECT @Error_Message = 'An error occurred while selecting info from the ADMIN_SYSTEM_T table.'
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

GRANT EXECUTE ON dbo.proc_admin_system_lu_select TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.proc_admin_system_lu_select TO Support
GO
