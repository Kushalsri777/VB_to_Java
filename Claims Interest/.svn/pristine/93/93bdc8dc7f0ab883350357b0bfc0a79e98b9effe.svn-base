SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[func_LPad]') and OBJECTPROPERTY(id, N'IsFunction') = 1)
drop function [dbo].[func_LPad]
GO

CREATE  FUNCTION [dbo].[func_LPad] (@InputValue SQL_VARIANT
	,@OutputLength INTEGER
	,@PadChar CHAR(1)
	,@UseImplicitDecimals CHAR(1))
-- Date: 03/11/2003
-- Author: Betsy Walker
--
-- Purpose: It returns a string containing the input value padded with leading whatevers, 
--          based on the specified desired length and pad character.
--          Accommodates null values and output values up to 100 characters
--
--          If the @UseImplicitDecimals = 'Y', then the input value will be massaged
--          to remove its decimal point.
--  
--   Modifications:
--  
RETURNS VARCHAR(100) AS  
BEGIN
	DECLARE @Result    AS VARCHAR(100)
    DECLARE @Interim  AS VARCHAR(100) 

	SELECT @Interim = 
		CASE
			WHEN @UseImplicitDecimals = 'Y'
				THEN REPLACE(CONVERT(VARCHAR(100), ISNULL(@InputValue,'')), '.', '')
			ELSE
				CONVERT(VARCHAR(100), ISNULL(@InputValue,''))
		END
    
	SELECT @Result = 
	    CASE
    	    WHEN DATALENGTH(@Interim) < @OutputLength 
        	    THEN REPLICATE(@PadChar, @OutputLength - DATALENGTH(@Interim)) +  @Interim
	        ELSE
    	        @Interim
	    END
    
	RETURN ( @Result )
END
GO



SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


GRANT EXECUTE ON dbo.func_LPad TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.func_LPad TO Support
GO