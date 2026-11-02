-- SQL_STORED_PROCEDURE copied from the original database.
CREATE   PROCEDURE dbo.usp_CalculateTDS_FY2627
    @EmployeeCode NVARCHAR(50) = NULL      -- NULL = all employees
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @FYStart DATE = '2026-04-01';
    DECLARE @FYEnd   DATE = '2027-03-31';

    ;WITH SrcClean AS (
        SELECT
            Employee,
            LTRIM(RTRIM(WorkOrderNo)) AS WorkOrderNo,
            Name, Month, Year, Salary,
            -- v16: calculated gross = CTC - PF_Employer (no arrears)
            CASE WHEN CTC IS NULL THEN NULL
                 ELSE CTC - ISNULL(PF_Employer, 0) END AS GrossAmount,
            IncomeTax, WOStartDate, WOEndDate,
            resigndate, tdsopt,
            DATEFROMPARTS(Year, Month, 1) AS PeriodDate,
            COALESCE(DATEADD(MONTH, 1, month_year), DATEFROMPARTS(Year, Month, 1)) AS PaymentDate
        FROM Payrollall1
    ),
    SrcEffective AS (
        SELECT *,
            DATEDIFF(MONTH, PeriodDate, PaymentDate) AS LagMonths,
            CASE WHEN DATEDIFF(MONTH, PeriodDate, PaymentDate) > 1
                 THEN PaymentDate
                 ELSE PeriodDate
            END AS EffectiveDate
        FROM SrcClean
    ),
    FYMonths AS (
        SELECT 4  AS [Month], 2026 AS [Year] UNION ALL
        SELECT 5,  2026 UNION ALL SELECT 6,  2026 UNION ALL SELECT 7,  2026 UNION ALL
        SELECT 8,  2026 UNION ALL SELECT 9,  2026 UNION ALL SELECT 10, 2026 UNION ALL
        SELECT 11, 2026 UNION ALL SELECT 12, 2026 UNION ALL SELECT 1,  2027 UNION ALL
        SELECT 2,  2027 UNION ALL SELECT 3,  2027
    ),
    EmpResign AS (
        -- one resignation date per employee (NULL if still active)
        SELECT Employee, MAX(resigndate) AS ResignDate
        FROM SrcEffective
        GROUP BY Employee
    ),
    EmpWO AS (
        SELECT Employee, WorkOrderNo, MAX(Name) AS Name, MAX(tdsopt) AS TDSOpt,
               MAX(WOStartDate) AS WOStartDate, MAX(WOEndDate) AS WOEndDate
        FROM SrcEffective
        GROUP BY Employee, WorkOrderNo
    ),
    LatestSalary AS (
        -- v16: whole latest salary month (all its rows added), not one random row
        SELECT Employee, WorkOrderNo, MonthlySalary,
               ROW_NUMBER() OVER (PARTITION BY Employee, WorkOrderNo ORDER BY Year DESC, Month DESC) AS rn
        FROM (
            SELECT Employee, WorkOrderNo, Year, Month, SUM(GrossAmount) AS MonthlySalary
            FROM SrcEffective
            WHERE GrossAmount IS NOT NULL
            GROUP BY Employee, WorkOrderNo, Year, Month
        ) m
    ),
    Base AS (
        SELECT
            e.Employee, e.WorkOrderNo, e.Name, e.TDSOpt, e.WOStartDate, e.WOEndDate,
            ls.MonthlySalary, er.ResignDate
        FROM EmpWO e
        JOIN LatestSalary ls
          ON ls.Employee = e.Employee AND ls.WorkOrderNo = e.WorkOrderNo AND ls.rn = 1
        LEFT JOIN EmpResign er
          ON er.Employee = e.Employee
        WHERE (e.WOStartDate <= @FYEnd AND e.WOEndDate >= @FYStart)
           OR EXISTS (
                SELECT 1 FROM SrcEffective s2
                WHERE s2.Employee = e.Employee
                  AND s2.WorkOrderNo = e.WorkOrderNo
                  AND s2.GrossAmount IS NOT NULL
                  AND s2.LagMonths > 1
                  AND s2.EffectiveDate >= @FYStart AND s2.EffectiveDate <= @FYEnd
           )
    ),
    ActualByFYMonth AS (
        SELECT
            b.Employee, b.WorkOrderNo,
            YEAR(s.EffectiveDate)  AS EffYear,
            MONTH(s.EffectiveDate) AS EffMonth,
            SUM(s.GrossAmount)          AS ActualAmount,
            SUM(ISNULL(s.IncomeTax, 0)) AS ActualIncomeTax
        FROM Base b
        JOIN SrcEffective s
          ON s.Employee = b.Employee AND s.WorkOrderNo = b.WorkOrderNo AND s.GrossAmount IS NOT NULL
        GROUP BY b.Employee, b.WorkOrderNo, YEAR(s.EffectiveDate), MONTH(s.EffectiveDate)
    ),
    PeriodRecorded AS (
        SELECT DISTINCT
            b.Employee, b.WorkOrderNo, s.Year AS PYear, s.Month AS PMonth
        FROM Base b
        JOIN SrcEffective s
          ON s.Employee = b.Employee AND s.WorkOrderNo = b.WorkOrderNo AND s.GrossAmount IS NOT NULL
    ),
    MonthlyCalc AS (
        SELECT
            b.Employee, b.WorkOrderNo,
            CASE
                WHEN a.ActualAmount IS NOT NULL THEN a.ActualAmount
                WHEN pr.PYear IS NULL
                     AND DATEFROMPARTS(fm.Year, fm.Month, 1) BETWEEN b.WOStartDate AND b.WOEndDate
                     AND (b.ResignDate IS NULL OR DATEFROMPARTS(fm.Year, fm.Month, 1) <= b.ResignDate)
                     THEN b.MonthlySalary
                ELSE 0
            END AS MonthAmount,
            CASE WHEN a.ActualAmount IS NOT NULL THEN 1 ELSE 0 END AS IsActual,
            ISNULL(a.ActualIncomeTax, 0) AS MonthIncomeTax
        FROM Base b
        CROSS JOIN FYMonths fm
        LEFT JOIN ActualByFYMonth a
               ON a.Employee = b.Employee AND a.WorkOrderNo = b.WorkOrderNo
              AND a.EffYear = fm.Year AND a.EffMonth = fm.Month
        LEFT JOIN PeriodRecorded pr
               ON pr.Employee = b.Employee AND pr.WorkOrderNo = b.WorkOrderNo
              AND pr.PYear = fm.Year AND pr.PMonth = fm.Month
    ),
    EmployeeTotals AS (
        SELECT
            Employee,
            SUM(CASE WHEN IsActual = 1 THEN MonthAmount ELSE 0 END)  AS ActualSalaryPaid_FY2627,
            SUM(CASE WHEN IsActual = 0 THEN MonthAmount ELSE 0 END)  AS ProposedSalary_FY2627,
            SUM(MonthIncomeTax)                                      AS TDSAlreadyPaid_FY2627,
            SUM(CASE WHEN IsActual = 0 AND MonthAmount > 0 THEN 1 ELSE 0 END) AS PendingMonths
        FROM MonthlyCalc
        GROUP BY Employee
    ),
    LatestWO AS (
        SELECT Employee, WorkOrderNo, Name, TDSOpt, WOStartDate, WOEndDate, MonthlySalary, ResignDate,
               ROW_NUMBER() OVER (PARTITION BY Employee ORDER BY WOStartDate DESC) AS rn
        FROM Base
    )
    SELECT
        l.Employee                                                   AS EmployeeCode,
        l.Name                                                        AS EmployeeName,
        l.TDSOpt,
        l.WorkOrderNo,
        l.WOStartDate,
        l.WOEndDate,
        l.ResignDate,
        l.MonthlySalary,
        et.ActualSalaryPaid_FY2627,
        et.ProposedSalary_FY2627,
        et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627         AS TotalSalary_FY2627,
        et.TDSAlreadyPaid_FY2627,

        -- New regime
        (et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 75000            AS TaxableIncome_New,
        dbo.fn_TDS_NewRegime_FY2627((et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 75000) AS AnnualTaxWithCess_New,
        dbo.fn_TDS_NewRegime_FY2627((et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 75000) - et.TDSAlreadyPaid_FY2627 AS RemainingTDS_New,
        CASE WHEN et.PendingMonths > 0 THEN
            (dbo.fn_TDS_NewRegime_FY2627((et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 75000) - et.TDSAlreadyPaid_FY2627) / et.PendingMonths
        ELSE NULL END                                                              AS MonthlyTDS_New,

        -- Old regime
        (et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 50000            AS TaxableIncome_Old,
        dbo.fn_TDS_OldRegime_FY2627((et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 50000) AS AnnualTaxWithCess_Old,
        dbo.fn_TDS_OldRegime_FY2627((et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 50000) - et.TDSAlreadyPaid_FY2627 AS RemainingTDS_Old,
        CASE WHEN et.PendingMonths > 0 THEN
            (dbo.fn_TDS_OldRegime_FY2627((et.ActualSalaryPaid_FY2627 + et.ProposedSalary_FY2627) - 50000) - et.TDSAlreadyPaid_FY2627) / et.PendingMonths
        ELSE NULL END                                                              AS MonthlyTDS_Old,

        et.PendingMonths
    INTO #TDS
    FROM LatestWO l
    JOIN EmployeeTotals et ON et.Employee = l.Employee
    WHERE l.rn = 1
      AND (@EmployeeCode IS NULL OR CAST(l.Employee AS NVARCHAR(50)) = @EmployeeCode);

    -- Final output: all v14 columns + the employee's own regime picked from TDSOpt
    SELECT t.*,
        CASE WHEN t.TDSOpt = 1 THEN 'Old' ELSE 'New' END                    AS Regime,
        CASE WHEN t.TDSOpt = 1 THEN t.AnnualTaxWithCess_Old ELSE t.AnnualTaxWithCess_New END AS AnnualTax,
        CASE WHEN t.TDSOpt = 1 THEN t.RemainingTDS_Old      ELSE t.RemainingTDS_New      END AS RemainingTDS,
        CASE
            WHEN t.PendingMonths IS NULL OR t.PendingMonths = 0 THEN NULL
            WHEN (CASE WHEN t.TDSOpt = 1 THEN t.MonthlyTDS_Old ELSE t.MonthlyTDS_New END) < 0 THEN 0
            ELSE ROUND(CASE WHEN t.TDSOpt = 1 THEN t.MonthlyTDS_Old ELSE t.MonthlyTDS_New END, 0)
        END                                                                    AS MonthlyTDS
    FROM #TDS t
    ORDER BY t.EmployeeCode;
END;
GO
