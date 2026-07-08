-- Test queries for the SSMS Results View extension.
-- Run each batch in SSMS 22 with the extension installed and verify the
-- "Results View" tab captures, filters, and sorts correctly.

-- 1. Small result set: basic capture, filtering, and sorting.
SELECT TOP (50)
    o.object_id      AS ObjectId,
    o.name           AS ObjectName,
    o.type_desc      AS TypeDescription,
    o.create_date    AS CreateDate,
    o.modify_date    AS ModifyDate
FROM sys.objects AS o
ORDER BY o.name;
GO

-- 2. Large result set (1,000,000 rows): progressive load, UI responsiveness,
--    filter/sort latency, and the row cap + truncation indicator.
WITH n AS (
    SELECT TOP (1000000)
        ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS RowNum
    FROM sys.all_columns AS a
    CROSS JOIN sys.all_columns AS b
)
SELECT
    RowNum,
    CONCAT('Customer-', RowNum % 10000)                    AS CustomerName,
    CONCAT('Region-', RowNum % 12)                         AS Region,
    CAST(RowNum % 100000 AS decimal(12, 2)) / 100          AS Amount,
    DATEADD(SECOND, RowNum % 86400, '2026-01-01')          AS CreatedAt,
    CASE WHEN RowNum % 7 = 0 THEN NULL
         ELSE CONCAT('note for row ', RowNum) END          AS Notes
FROM n;
GO

-- 3. Multiple result sets: selector should list all three; sets 2 and 3
--    load lazily on first selection.
SELECT TOP (10) name, database_id, create_date FROM sys.databases;
SELECT TOP (20) name, object_id, type_desc FROM sys.objects ORDER BY name;
SELECT TOP (5) name, schema_id FROM sys.schemas;
GO

-- 4. Edge cases: NULLs, empty strings, unicode, quotes/commas (CSV escaping),
--    long text, and numeric-vs-text sort behavior.
SELECT *
FROM (VALUES
    (1,  NULL,          N'',                    N'plain'),
    (2,  N'has,comma',  N'has "quotes"',        N'ünïcødé ✓'),
    (10, N'ten',        REPLICATE(N'x', 4000),  N'long text'),
    (3,  N'multi
line',                          N'tab	separated',    N'control chars')
) AS v (SortKey, A, B, C);
GO

-- 5. No grid results: Results View should report "no grid results" gracefully.
PRINT 'messages only — nothing to capture';
GO

-- 6. Empty result set: headers only, zero rows.
SELECT name, object_id FROM sys.objects WHERE 1 = 0;
GO
