-- SQL_SCALAR_FUNCTION copied from the original database.
CREATE   FUNCTION dbo.fn_TDS_NewRegime_FY2627 (@Taxable DECIMAL(18,2))
RETURNS DECIMAL(18,2)
AS
BEGIN
    DECLARE @Tax DECIMAL(18,2) = 0;

    IF @Taxable IS NULL OR @Taxable <= 0 RETURN 0;

    -- Sec 288A: taxable income rounded to nearest Rs 10
    SET @Taxable = ROUND(@Taxable / 10.0, 0) * 10;

    SET @Tax =
          CASE WHEN @Taxable > 400000  THEN (CASE WHEN @Taxable < 800000  THEN @Taxable ELSE 800000  END - 400000)  * 0.05 ELSE 0 END
        + CASE WHEN @Taxable > 800000  THEN (CASE WHEN @Taxable < 1200000 THEN @Taxable ELSE 1200000 END - 800000)  * 0.10 ELSE 0 END
        + CASE WHEN @Taxable > 1200000 THEN (CASE WHEN @Taxable < 1600000 THEN @Taxable ELSE 1600000 END - 1200000) * 0.15 ELSE 0 END
        + CASE WHEN @Taxable > 1600000 THEN (CASE WHEN @Taxable < 2000000 THEN @Taxable ELSE 2000000 END - 1600000) * 0.20 ELSE 0 END
        + CASE WHEN @Taxable > 2000000 THEN (CASE WHEN @Taxable < 2400000 THEN @Taxable ELSE 2400000 END - 2000000) * 0.25 ELSE 0 END
        + CASE WHEN @Taxable > 2400000 THEN (@Taxable - 2400000) * 0.30 ELSE 0 END;

    -- Sec 87A rebate
    IF @Taxable <= 1200000
        SET @Tax = 0;
    -- Marginal relief just above 12L
    ELSE IF @Tax > (@Taxable - 1200000)
        SET @Tax = @Taxable - 1200000;

    -- 4% cess, each rounded to the rupee (matches IT calculator)
    SET @Tax = ROUND(@Tax, 0);
    RETURN @Tax + ROUND(@Tax * 0.04, 0);
END;
GO
