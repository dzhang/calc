IF OBJECT_ID('dbo.sigfig', 'FN') IS NOT NULL
  DROP FUNCTION dbo.sigfig;
GO

/* 
  this handles odd/even conversion also
  e.g sigfig(0.125, 2) convert to 0.12
  
  notice the conversion of int to decimal, this is needed to make sure the calculated number does not be converted to int, thus 
  may give wrong result or cause divide by 0 error
*/
CREATE FUNCTION sigfig(@number FLOAT, @sf SMALLINT) RETURNS FLOAT 
AS
BEGIN

DECLARE @sign SMALLINT, @magnitude INT;
DECLARE @conversion DECIMAL(38,17), @converted DECIMAL(38,17), @frac DECIMAL(38,17), @factor DECIMAL(38,17), @epa DECIMAL(38,17);

IF @number = 0 RETURN 0;
    
SELECT @sign = SIGN(@number);
SELECT @number = ABS(@number);

SELECT @magnitude = FLOOR(LOG10(@number)) + 1;
-- need to cast 10 to decimal, otherwise conversion would be calculated as integer, and could be 0
SELECT @conversion = POWER(CAST(10 AS DECIMAL(38,17)), (@sf - @magnitude));
SELECT @converted = @number * @conversion;

SELECT @frac = @converted - CAST(FLOOR(@converted) AS DECIMAL);
SELECT @factor = POWER(10.0, 10.0);
SELECT @frac = FLOOR(@frac * @factor) / @factor;

-- Calculate the epa rounding only if fractional part is exactly 0.5 and number is even
If @frac = 0.5 And FLOOR(@converted) % 2.00 = 0 
    SELECT @epa = FLOOR(@converted);
Else   -- Just use regular rounding
    SELECT @epa = FLOOR(@converted + 0.5);

SELECT @number = @sign * (@epa / @conversion);

RETURN @number;
     
END
GO

GRANT EXECUTE ON dbo.sigfig TO Admins;
GO
GRANT EXECUTE ON dbo.sigfig TO Analysts;
GO
GRANT EXECUTE ON dbo.sigfig TO Mgmt;
GO