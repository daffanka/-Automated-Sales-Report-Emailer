WITH UNITED_DATA AS (
    -- Revenue from NAVDB
    SELECT
        entity AS office,
        salesperson_name,
        amount_in_sgd AS amount_SGD,
        amount_local AS amount_LCY,
        posting_date AS Date,
        is_invoiced,
        CAST('NAVDB' AS NVARCHAR(10)) AS Source,
        CAST('Revenue' AS NVARCHAR(10)) AS Type
    FROM NAVDBCONSO.dbo.datamart_sales
    WHERE 
        YEAR(posting_date) = 2025 AND is_invoiced = 1
        AND table_name NOT IN (
            'Esco Philippines Inc',
            'Esco Micro Pte Ltd',
            'Esco Lifesciences (Thailand)',
            'PT Esco Utama'
        )

    UNION ALL

    -- Revenue from BCDB
    SELECT
        entity AS office,
        salesperson_name,
        amount_in_sgd AS amount_SGD,
        amount_local AS amount_LCY,
        posting_date AS Date,
        is_invoiced,
        CAST('BCDB' AS NVARCHAR(10)) AS Source,
        CAST('Revenue' AS NVARCHAR(10)) AS Type
    FROM BCDB.dbo.datamart_sales_bcdb
    WHERE 
        YEAR(posting_date) = 2025 AND is_invoiced = 1

    UNION ALL

    -- Bookings from NAVDB
    SELECT
        entity AS office,
        salesperson_name,
        amount_bookings_in_sgd AS amount_SGD,
        amount_local_bookings AS amount_LCY,
        order_date_ AS Date,
        is_invoiced,
        CAST('NAVDB' AS NVARCHAR(10)) AS Source,
        CAST('Bookings' AS NVARCHAR(10)) AS Type
    FROM NAVDBCONSO.dbo.datamart_sales
    WHERE 
        YEAR(order_date_) = 2025
        AND table_name NOT IN (
            'Esco Philippines Inc',
            'Esco Micro Pte Ltd',
            'Esco Lifesciences (Thailand)',
            'PT Esco Utama'
        )

    UNION ALL

    -- Bookings from BCDB
    SELECT
        entity AS office,
        salesperson_name,
        amount_bookings_in_sgd AS amount_SGD,
        amount_local_bookings AS amount_LCY,
        order_date_ AS Date,
        is_invoiced,
        CAST('BCDB' AS NVARCHAR(10)) AS Source,
        CAST('Bookings' AS NVARCHAR(10)) AS Type
    FROM BCDB.dbo.datamart_sales_bcdb
    WHERE 
        YEAR(order_date_) = 2025
)

SELECT
    DATENAME(MONTH, Date) AS [Month],
	MONTH(Date) AS [MonthNumber],
	salesperson_name,
	office,
    ROUND(SUM(CASE WHEN Type = 'Revenue' THEN amount_SGD ELSE 0 END), 0) AS [Revenue (SGD)],
    ROUND(SUM(CASE WHEN Type = 'Revenue' THEN amount_LCY ELSE 0 END), 0) AS [Revenue (LCY)],

    ROUND(SUM(CASE WHEN Type = 'Bookings' THEN amount_SGD ELSE 0 END), 0) AS [Bookings (SGD)],
    ROUND(SUM(CASE WHEN Type = 'Bookings' THEN amount_LCY ELSE 0 END), 0) AS [Bookings (LCY)]

FROM UNITED_DATA

GROUP BY
    DATENAME(MONTH, Date),
	MONTH(Date),
	salesperson_name,
	office




