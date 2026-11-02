-- SQL_STORED_PROCEDURE copied from the original database.
/* =====================================================================================
   dbo.usp_RefreshPayrollall  -  for table dbo.Payrollall1   (v2, Oct-2026)
   Same name as before, so the chatbot's "Refresh data" button needs no change.

   Steps:
     1) FinaluniqCode = own UniqueNo, else the next later UniqueNo of the same WorkOrderNo + Employee
     2) month_year    = date inside FinaluniqCode (chars 3-10, ddMMyyyy) minus 1 month
     3) delete unwanted work orders and the 2024 / 2025 month_year rows
     4) optional: work order-wise margin result (one row per WorkOrderNo + Year + Month), now with InvoiceNo

   Run manually in SSMS:   EXEC dbo.usp_RefreshPayrollall;                         -- with margin result
                           EXEC dbo.usp_RefreshPayrollall @ShowMarginReport = 0;   -- only refresh
   ===================================================================================== */
CREATE   PROCEDURE dbo.usp_RefreshPayrollall
    @ShowMarginReport BIT = 1
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;      -- any error stops the procedure (the chatbot then rolls everything back)

    -- 3a) optional: remove work orders you never want in reports (less work for the updates below), e.g.
    -- DELETE FROM dbo.Payrollall1 WHERE WorkOrderNo IN ('<WO-1>', '<WO-2>');

    -- 1) FinaluniqCode (blank UniqueNo is treated the same as NULL)
    UPDATE p
    SET FinaluniqCode = COALESCE(NULLIF(LTRIM(RTRIM(p.UniqueNo)), ''), x.next_uniqcode)
    FROM dbo.Payrollall1 p
    OUTER APPLY
    (
        SELECT TOP 1 p2.UniqueNo AS next_uniqcode
        FROM dbo.Payrollall1 p2
        WHERE p2.WorkOrderNo = p.WorkOrderNo
          AND p2.Employee    = p.Employee
          AND NULLIF(LTRIM(RTRIM(p2.UniqueNo)), '') IS NOT NULL
          AND (    p2.[Year] > p.[Year]
               OR (p2.[Year] = p.[Year] AND p2.[Month] >= p.[Month]))
        ORDER BY p2.[Year], p2.[Month]
    ) x;

    -- 2) month_year (rows without a usable code get NULL, not the 1900-01-01 left by the INSERT's '')
    UPDATE dbo.Payrollall1
    SET month_year =
        DATEADD(MONTH, -1,
            TRY_CONVERT(DATE,
                STUFF(STUFF(SUBSTRING(FinaluniqCode, 3, 8), 5, 0, '/'), 3, 0, '/'),
                103));

    -- 3b) remove 2024 and 2025 processing months (same as the old LIKE '2024-%' / '2025-%')
    DELETE FROM dbo.Payrollall1
    WHERE month_year >= '2024-01-01' AND month_year < '2026-01-01';

    -- 4) work order-wise margin (base sales once per WorkOrderNo + Year + Month, CTC = sum of its employees)
    IF @ShowMarginReport = 1
    BEGIN
        WITH CTE AS
        (
            SELECT
                Category,
                WorkOrderNo,
                InvoiceNo,
                [Year],
                [Month],
                InvoiceAmount,
                FinaluniqCode,
                FORMAT(month_year, 'MMM-yyyy') AS month_year,
                Department,
                WOStartDate,
                WOEndDate,
                UniqueNo,
                ROUND(InvoiceAmount * 100 / 118.0, 2)                    AS BaseSalesAmount,
                SUM(CTC) OVER (PARTITION BY WorkOrderNo, [Year], [Month]) AS CTC,
                ROW_NUMBER() OVER (PARTITION BY WorkOrderNo, [Year], [Month]
                                   ORDER BY Employee)                     AS rn
            FROM dbo.Payrollall1
        )
        SELECT
            Category, WorkOrderNo, InvoiceNo, [Year], [Month], month_year,
            InvoiceAmount, BaseSalesAmount, CTC,
            ROUND(BaseSalesAmount - CTC, 2)                                       AS Margin,
            ROUND((BaseSalesAmount - CTC) * 100.0 / NULLIF(CTC, 0), 2)            AS [Margin % on CTC],
            FinaluniqCode, Department, WOStartDate, WOEndDate, UniqueNo
        FROM CTE
        WHERE rn = 1
        ORDER BY WorkOrderNo, [Year], [Month];
    END
END;
GO
