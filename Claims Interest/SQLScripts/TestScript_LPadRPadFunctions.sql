DECLARE @E1 AS CHAR(25)
DECLARE @E2 AS VARCHAR(25)
DECLARE @E3 AS DECIMAL(11,2)
DECLARE @E4 AS INTEGER
DECLARE @F1 AS CHAR(25)
DECLARE @F2 AS VARCHAR(25)
DECLARE @F3 AS DECIMAL(11,2)
DECLARE @F4 AS INTEGER
DECLARE @SV AS SQL_VARIANT
DECLARE @VC AS VARCHAR(100)
DECLARE @MSG AS VARCHAR(200)

SELECT @E1 = NULL
	,@E2 = NULL
	,@E3 = NULL
	,@E4 = NULL
	,@F1 = 'ABC'
	,@F2 = 'ABC'
	,@F3 = 12345.67
	,@F4 = 123
	,@SV = NULL

PRINT 'E1 = [' + dbo.func_LPad(@E1, 14, '0', 'N') + ']'
PRINT 'E2 = [' + dbo.func_LPad(@E2, 14, '0', 'N') + ']'
PRINT 'E3 = [' + dbo.func_LPad(@E3, 14, '0', 'N') + ']'
PRINT 'E4 = [' + dbo.func_LPad(@E4, 14, '0', 'N') + ']'
PRINT 'F1 = [' + dbo.func_LPad(@F1, 14, '0', 'N') + ']'
PRINT 'F2 = [' + dbo.func_LPad(@F2, 14, '0', 'N') + ']'
PRINT 'F3 = [' + dbo.func_LPad(@F3, 14, '0', 'N') + ']'
PRINT 'F4 = [' + dbo.func_LPad(@F4, 14, '0', 'N') + ']'

