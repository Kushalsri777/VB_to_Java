SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[func_RPad]') and OBJECTPROPERTY(id, N'IsFunction') = 1)
drop function [dbo].[func_RPad]
GO

CREATE  FUNCTION [dbo].[func_RPad] (@InputValue SQL_VARIANT, @OutputLength INTEGER, @PadChar CHAR(1))  
-- Date: 03/11/2003
-- Author: Betsy Walker
--
-- Purpose: It returns a string containing the input value padded with trailing whatevers, 
--          based on the specified desired length and pad character.
--          Accommodates null values and output values up to 100 characters.
--  
--   Modifications:
--  
RETURNS VARCHAR(100) AS  
BEGIN
	DECLARE @Result    AS VARCHAR(100)
    DECLARE @Interim   AS VARCHAR(100) 

	SELECT @Interim = CONVERT(VARCHAR(100), ISNULL(@InputValue,''))

    SELECT @Result = 
		CASE
        	WHEN DATALENGTH(@Interim) < @OutputLength
            	THEN @Interim + REPLICATE(@PadChar, @OutputLength - DATALENGTH(@Interim))
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


GRANT EXECUTE ON dbo.func_RPad TO AppRoleClaims
GO
GRANT EXECUTE ON dbo.func_RPad TO Support
GO