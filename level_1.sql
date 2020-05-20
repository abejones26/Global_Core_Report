-- Privacy of Randa Accessories, I am only able to show you the frame work of this database and datatable.
-- See screenshot for visualizations.

-- Datatable and descriptions
    -- ALP: Report for Inventory Management and Warehouse
        -- ID: Key
        -- Cut Number: Purchase order number
        -- ALP: Alpha size label

    -- Buying Group: Report of Sales, Forecasting, and Inventory Management
        -- ID: Key
        -- Style: Product style serial number
        -- Color: Product color code
        -- Buying Groups: Sales customer buying groups and ID of buying groups

    -- CORPO2WK is the report I am re-developing.
    -- CORPO2WK: Original 30 years old version of Randa's high level global product production core report.
        -- ID: Key
        -- Company: Code to determine which category a product belongs such as Belt, Shoe or luggage.
        -- Division: Code to determine which location of the company category belongs
        -- Core Code: High level category of product style and color that determines the amount of inventory productions and purchase orders.
        -- Sort Supplier: Suppliers' ID number
        -- Supplier Number: Suppliers' ID number
        -- Style: Product style number
        -- Color: Product color code
        -- Status: Empty
        -- Tier: Empty
        -- Effective Date: Empty
        -- Beginning Balance: Purchase Order amount at the beginning of sale
        -- Open Balance: Purchase Order amount after customer order picked
        -- Shipped Quantity: Amount of pieces shipped
        -- Heading Month 1 to Month 6: Placeholder of month 1 to 6
        -- Forecast Month 1 to Month 8: Product Orders forecast for month 1 to Month 8
        -- Receipts Month 1 to Month 6: Amount of pieces being received at the warehouse for month 1 to Month 6
        -- Balance Month 1 to Month 6: Amount of pieces after shipping and forecast, divided by next month forecast

    -- Diver: Report for Supply Planners, Inventory Management, Sourcing, and Sales
        -- ID: Key
        -- Style: Product style serial number
        -- Plant Code: Factory name and name code
        -- Inventory Type: Product caterogy
        -- Royalty Code: Customer name and name code

    -- Due to report size i'll only give description of featured columns
    -- WMAWIPSE: Product report for Warehouse, Supply Planner, Inventory Management, Sourcing, Sales,Forecasting and Merchadising
        -- ID:Key
        -- Shop Code table #26: Factory name code
        -- Ship Via Table #01: Shipping method freight, air or boat.
        -- Cut number: Purchase order number
        -- Cut sequence number: Split of purchase order number
        -- Shade code - table #45: Determine the location of products in warehouse
        -- Open Qty: Purchase Order amount after customer order picked
        -- Original Ship date: Ship date when purchase order is cut
        -- Requested Ship date: Date change requested by factory or Randa company date change
        -- Shipment date: Date order shipped from factory
        -- Warehouse Date: Date order arrived at the warehouse
        -- Revised reason code: Shipping delayed reason code
        -- Comments: company comments about products or order

SELECT "" AS Duplicate,
    [CORP02WK]![Style] & [CORP02WK]![Color] AS [KEY], 
    CORP02WK.[Core Code], 
    CORP02WK.[Supplier Number], 
    Diver.[Plant Code], 
    Diver.[Inventory Type], 
    Diver.[Royalty Code], 
    CORP02WK.Style, 
    CORP02WK.Color, 
    CORP02WK.[Beginning Balance], 
    CORP02WK.[Shipped Quantity], 
    CORP02WK.[Open Balance], 
    CORP02WK.[Heading Month 1], 
    CORP02WK.[Forecast Month 1], 
    "" AS [PO 1], 
    "" AS [Mo 1, Wk 1], 
    "" AS [Mo 1, Wk 2], 
    "" AS [Mo 1, Wk 3], 
    "" AS [Mo 1, Wk 4], 
    "" AS [Mo 1, Wk 5], 
    CORP02WK.[Receipts Month 1], 
    CORP02WK.[Balance Month 1], 
    [CORP02WK]![Balance Month 1]/[CORP02WK]![Forecast Month 2] AS [Percentage Month 1], 
    CORP02WK.[Forecast Month 2], 
    "" AS [PO 2], 
    "" AS [Mo 2, Wk 1], 
    "" AS [Mo 2, Wk 2], 
    "" AS [Mo 2, Wk 3], 
    "" AS [Mo 2, Wk 4], 
    "" AS [Mo 2, Wk 5], 
    CORP02WK.[Receipts Month 2], 
    CORP02WK.[Balance Month 2], 
    [CORP02WK]![Balance Month 2]/[CORP02WK]![Forecast Month 3] AS [Percentage Month 2], 
    CORP02WK.[Forecast Month 3], 
    "" AS [PO 3], 
    "" AS [Mo 3, Wk 1], 
    "" AS [Mo 3, Wk 2], 
    "" AS [Mo 3, Wk 3], 
    "" AS [Mo 3, Wk 4], 
    "" AS [Mo 3, Wk 5], 
    CORP02WK.[Receipts Month 3], 
    CORP02WK.[Balance Month 3], 
    [CORP02WK]![Balance Month 3]/[CORP02WK]![Forecast Month 4] AS [Percentage Month 3], 
    CORP02WK.[Forecast Month 4], 
    "" AS [PO 4], 
    "" AS [Mo 4, Wk 1], 
    "" AS [Mo 4, Wk 2], 
    "" AS [Mo 4, Wk 3], 
    "" AS [Mo 4, Wk 4], 
    "" AS [Mo 4, Wk 5], 
    CORP02WK.[Receipts Month 4], 
    CORP02WK.[Balance Month 4], 
    [CORP02WK]![Balance Month 4]/[CORP02WK]![Forecast Month 5] AS [Percentage Month 4], 
    CORP02WK.[Forecast Month 5], 
    "" AS [PO 5], 
    "" AS [Mo 5 Wk 1], 
    "" AS [Mo 5 Wk 2], 
    "" AS [Mo 5 Wk 3], 
    "" AS [Mo 5 Wk 4], 
    "" AS [Mo 5 Wk 5], 
    CORP02WK.[Receipts Month 5], 
    CORP02WK.[Balance Month 5], 
    [CORP02WK]![Balance Month 5]/[CORP02WK]![Forecast Month 6] AS [Percentage Month 5], 
    CORP02WK.[Forecast Month 6], 
    "" AS [PO 6], 
    "" AS [Mo 6 Wk 1], 
    "" AS [Mo 6 Wk 2], 
    "" AS [Mo 6 Wk 3], 
    "" AS [Mo 6 Wk 4], 
    "" AS [Mo 6 Wk 5], 
    CORP02WK.[Receipts Month 6], 
    CORP02WK.[Balance Month 6], 
    [CORP02WK]![Balance Month 6]/[CORP02WK]![Forecast Month 7] AS [Percentage Month 6], 
    CORP02WK.[Forecast Month 7], 
    CORP02WK.[Forecast Month 8], 
    WMAWIPSE.[Shade code - table #45], 
    WMAWIPSE.[Cut number], 
    WMAWIPSE.[Cut sequence number], 
    ALP.ALP, 
    WMAWIPSE.[Shop Code Table #26], 
    WMAWIPSE.[Open Qty], 
    "" AS [Rcpt Month], 
    WMAWIPSE.[Original Ship date], 
    WMAWIPSE.[Requested Ship date], 
    WMAWIPSE.[Shipment date], 
    "" AS [Curr Ship], 
    WMAWIPSE.[Warehouse Date], 
    WMAWIPSE.[Ship Via Table #01], 
    WMAWIPSE.[Revised reason code], 
    WMAWIPSE.Comments, 
    [Buying Group].[Buying Groups]
FROM (((((((((CORP02WK 
INNER JOIN WMAWIPSE ON (CORP02WK.Style = WMAWIPSE.Style) AND (CORP02WK.Color = WMAWIPSE.[Color Code Table #25])) 
INNER JOIN Month_Table_2 ON CORP02WK.[Heading Month 2] = Month_Table_2.[Heading Month]) 
INNER JOIN Month_Table_1 ON CORP02WK.[Heading Month 1] = Month_Table_1.[Heading Month]) 
INNER JOIN Month_Table_3 ON CORP02WK.[Heading Month 3] = Month_Table_3.[Heading Month]) 
INNER JOIN Month_Table_4 ON CORP02WK.[Heading Month 4] = Month_Table_4.[Heading Month]) 
INNER JOIN Month_Table_5 ON CORP02WK.[Heading Month 5] = Month_Table_5.[Heading Month]) 
INNER JOIN Month_Table_6 ON CORP02WK.[Heading Month 6] = Month_Table_6.[Heading Month]) 
INNER JOIN Diver ON CORP02WK.Style = Diver.Style) 
LEFT JOIN ALP ON WMAWIPSE.[Cut number] = ALP.[Cut Number]) 
INNER JOIN [Buying Group] ON (CORP02WK.Color = [Buying Group].Color) AND (CORP02WK.Style = [Buying Group].Style)
ORDER BY CORP02WK.[Core Code], 
    CORP02WK.Style, 
    CORP02WK.Color, 
    WMAWIPSE.[Warehouse Date];
