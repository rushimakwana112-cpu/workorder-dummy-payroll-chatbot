-- SQL_SCALAR_FUNCTION copied from the original database.
CREATE   FUNCTION dbo.fn_TDS_OldRegime_FY2627 (@TaxableIncome DECIMAL(18,2))
RETURNS DECIMAL(18,2)
AS
BEGIN
    DECLARE @Tax DECIMAL(18,2) = 0;
    IF @TaxableIncome > 1000000
        SET @Tax = 112500 + (@TaxableIncome - 1000000) * 0.30;
    ELSE IF @TaxableIncome > 500000
        SET @Tax = 12500 + (@TaxableIncome - 500000) * 0.20;
    ELSE IF @TaxableIncome > 250000
        SET @Tax = (@TaxableIncome - 250000) * 0.05;
    ELSE
        SET @Tax = 0;

    IF @TaxableIncome <= 500000
        SET @Tax = 0; -- Section 87A rebate

    RETURN ROUND(@Tax * 1.04, 2); -- + 4% health & education cess
END;
GO
