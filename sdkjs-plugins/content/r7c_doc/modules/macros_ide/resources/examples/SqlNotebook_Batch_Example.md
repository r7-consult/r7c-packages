# SQLite Notebook - Batch Operations Example

This example demonstrates how to execute multiple SQL queries at once using `Api.SqlNotebook.batch()`.

## Prerequisites
1. Open SQLite Notebook tab
2. Load MOCK_DATA.csv file
3. Run this macro

## Code

```javascript
(function() {
    var oWorksheet = Api.GetActiveSheet();

    // Batch multiple SQL queries
    var batchResult = Api.SqlNotebook.batch([
        {
            sql: 'SELECT COUNT(*) as total FROM "MOCK_DATA.csv"',
            filename: 'MOCK_DATA.csv'
        },
        {
            sql: 'SELECT gender, COUNT(*) as count FROM "MOCK_DATA.csv" GROUP BY gender',
            filename: 'MOCK_DATA.csv'
        },
        {
            sql: 'SELECT * FROM "MOCK_DATA.csv" WHERE gender = "Male" LIMIT 5',
            filename: 'MOCK_DATA.csv'
        },
        {
            sql: 'SELECT * FROM "MOCK_DATA.csv" WHERE gender = "Female" LIMIT 5',
            filename: 'MOCK_DATA.csv'
        }
    ]);

    if (!batchResult.ok) {
        oWorksheet.GetRange("A1").SetValue("ERROR: " + batchResult.error.message);
        return;
    }

    var row = 1;

    // Header
    oWorksheet.GetRange("A" + row).SetValue("Batch Query Results");
    oWorksheet.GetRange("A" + row).SetBold(true);
    oWorksheet.GetRange("A" + row).SetFontSize(14);
    row += 2;

    // Process each query result
    batchResult.results.forEach(function(result, index) {
        // Query header
        oWorksheet.GetRange("A" + row).SetValue("Query " + (index + 1) + ":");
        oWorksheet.GetRange("A" + row).SetBold(true);
        var queryHeaderColor = Api.CreateColorFromRGB(100, 150, 200);
        oWorksheet.GetRange("A" + row).SetFontColor(queryHeaderColor);
        row++;

        if (!result.ok) {
            oWorksheet.GetRange("A" + row).SetValue("ERROR: " + result.error.message);
            var errorColor = Api.CreateColorFromRGB(255, 0, 0);
            oWorksheet.GetRange("A" + row).SetFontColor(errorColor);
            row += 2;
            return;
        }

        var data = result.result;

        // Column headers
        for (var col = 0; col < data.columns.length; col++) {
            var cell = oWorksheet.GetRange(String.fromCharCode(65 + col) + row);
            cell.SetValue(data.columns[col]);
            cell.SetBold(true);
            var headerColor = Api.CreateColorFromRGB(220, 230, 240);
            cell.SetFillColor(headerColor);
        }
        row++;

        // Data rows
        for (var i = 0; i < data.rows.length; i++) {
            for (var col = 0; col < data.rows[i].length; col++) {
                var value = data.rows[i][col];
                oWorksheet.GetRange(String.fromCharCode(65 + col) + row).SetValue(value);
            }
            row++;
        }

        // Summary
        oWorksheet.GetRange("A" + row).SetValue("Rows returned: " + data.rows.length);
        oWorksheet.GetRange("A" + row).SetItalic(true);
        row += 2;
    });

    // Final summary
    row++;
    oWorksheet.GetRange("A" + row).SetValue("Total queries executed: " + batchResult.results.length);
    oWorksheet.GetRange("A" + row).SetBold(true);

    var successCount = 0;
    var errorCount = 0;
    batchResult.results.forEach(function(result) {
        if (result.ok) successCount++;
        else errorCount++;
    });

    row++;
    oWorksheet.GetRange("A" + row).SetValue("Successful: " + successCount);
    var successColor = Api.CreateColorFromRGB(0, 200, 0);
    oWorksheet.GetRange("A" + row).SetFontColor(successColor);

    row++;
    oWorksheet.GetRange("A" + row).SetValue("Failed: " + errorCount);
    if (errorCount > 0) {
        var failColor = Api.CreateColorFromRGB(255, 0, 0);
        oWorksheet.GetRange("A" + row).SetFontColor(failColor);
    }

    row += 2;
    oWorksheet.GetRange("A" + row).SetValue("✅ Batch execution completed!");
    var doneColor = Api.CreateColorFromRGB(0, 150, 0);
    oWorksheet.GetRange("A" + row).SetFontColor(doneColor);

})();
```

## Expected Output

The macro will display:
- Results from all 4 queries in sequence
- Total count of records
- Gender distribution
- Top 5 male records
- Top 5 female records
- Summary statistics (successful/failed queries)

## Benefits of Batch Operations

1. **Performance**: Execute multiple queries in one API call
2. **Error handling**: Each query result contains individual ok/error status
3. **Atomic operations**: All queries are processed together
4. **Simplified code**: No need for multiple Api.SqlNotebook.query() calls
