# SQLite Notebook - Get Sources Example

This example demonstrates:
1. How to get list of loaded data sources using `Api.SqlNotebook.getSources()`
2. How to use autocomplete for table names (try typing `FROM ` in SQL queries)

## Prerequisites
1. Open SQLite Notebook tab
2. Load one or more files (e.g., MOCK_DATA.csv, data.xlsx)
3. Run this macro

## Code

```javascript
(function() {
    var oWorksheet = Api.GetActiveSheet();

    // Get all loaded sources
    var sourcesResult = Api.SqlNotebook.getSources();

    if (!sourcesResult.ok) {
        oWorksheet.GetRange("A1").SetValue("ERROR: " + sourcesResult.error.message);
        return;
    }

    var sources = sourcesResult.sources;
    var row = 1;

    // Title
    oWorksheet.GetRange("A" + row).SetValue("Loaded Data Sources in SQLite Notebook");
    oWorksheet.GetRange("A" + row).SetBold(true);
    oWorksheet.GetRange("A" + row).SetFontSize(16);
    var titleColor = Api.CreateColorFromRGB(50, 100, 150);
    oWorksheet.GetRange("A" + row).SetFontColor(titleColor);
    row += 2;

    // Check if any sources are loaded
    if (sources.length === 0) {
        oWorksheet.GetRange("A" + row).SetValue("⚠ No data sources loaded.");
        oWorksheet.GetRange("A" + row + 1).SetValue("Please load files in SQLite Notebook first.");
        return;
    }

    // Summary
    oWorksheet.GetRange("A" + row).SetValue("Total sources: " + sources.length);
    oWorksheet.GetRange("A" + row).SetBold(true);
    row += 2;

    // Table headers
    var headers = ["Name", "Format", "Status", "Label"];
    for (var col = 0; col < headers.length; col++) {
        var cell = oWorksheet.GetRange(String.fromCharCode(65 + col) + row);
        cell.SetValue(headers[col]);
        cell.SetBold(true);
        var headerBg = Api.CreateColorFromRGB(200, 220, 240);
        cell.SetFillColor(headerBg);
    }
    row++;

    // List all sources
    sources.forEach(function(source) {
        // Name
        oWorksheet.GetRange("A" + row).SetValue(source.name);

        // Format
        oWorksheet.GetRange("B" + row).SetValue(source.format.toUpperCase());

        // Status
        var status = source.active ? "✅ Active" : "⭕ Inactive";
        oWorksheet.GetRange("C" + row).SetValue(status);

        if (source.active) {
            var activeColor = Api.CreateColorFromRGB(0, 150, 0);
            oWorksheet.GetRange("C" + row).SetFontColor(activeColor);
        }

        // Label
        oWorksheet.GetRange("D" + row).SetValue(source.label);

        row++;
    });

    row += 2;

    // Autocomplete tip
    oWorksheet.GetRange("A" + row).SetValue("💡 Autocomplete Tip:");
    oWorksheet.GetRange("A" + row).SetBold(true);
    var tipColor = Api.CreateColorFromRGB(255, 140, 0);
    oWorksheet.GetRange("A" + row).SetFontColor(tipColor);
    row++;

    oWorksheet.GetRange("A" + row).SetValue("When writing SQL queries in the macro editor:");
    row++;
    oWorksheet.GetRange("A" + row).SetValue("1. Type: var result = Api.SqlNotebook.query('SELECT * FROM ");
    row++;
    oWorksheet.GetRange("A" + row).SetValue("2. After FROM, press Ctrl+Space or just start typing");
    row++;
    oWorksheet.GetRange("A" + row).SetValue("3. You'll see autocomplete suggestions for all loaded sources!");
    row += 2;

    // Example queries for each source
    oWorksheet.GetRange("A" + row).SetValue("Example Queries:");
    oWorksheet.GetRange("A" + row).SetBold(true);
    row++;

    sources.forEach(function(source, index) {
        var exampleQuery = 'var result' + (index + 1) + ' = Api.SqlNotebook.query(\'SELECT * FROM "' + source.name + '" LIMIT 10\', "' + source.name + '");';
        oWorksheet.GetRange("A" + row).SetValue(exampleQuery);
        oWorksheet.GetRange("A" + row).SetFontSize(9);
        var codeColor = Api.CreateColorFromRGB(80, 80, 80);
        oWorksheet.GetRange("A" + row).SetFontColor(codeColor);

        // Highlight the source name
        var codeBg = Api.CreateColorFromRGB(240, 240, 240);
        oWorksheet.GetRange("A" + row).SetFillColor(codeBg);

        row++;
    });

    row += 2;
    oWorksheet.GetRange("A" + row).SetValue("✅ Sources listed successfully!");
    var doneColor = Api.CreateColorFromRGB(0, 150, 0);
    oWorksheet.GetRange("A" + row).SetFontColor(doneColor);

})();
```

## Testing Autocomplete

After running this macro, try creating a new macro with:

```javascript
(function() {
    // Method 1: Type after FROM with space
    var result = Api.SqlNotebook.query('SELECT * FROM ', 'MOCK_DATA.csv');
    //                                                ↑
    // After typing "FROM " (with space), press Ctrl+Space to see suggestions!

    // Method 2: Start typing table name
    var result2 = Api.SqlNotebook.query('SELECT * FROM M', 'MOCK_DATA.csv');
    //                                                 ↑
    // Type first letters, Monaco will show matching sources

    // Method 3: Use in batch queries
    var batchResult = Api.SqlNotebook.batch([
        { sql: 'SELECT * FROM ', filename: 'data.csv' }
        //                    ↑ Press Ctrl+Space here too!
    ]);
})();
```

**Important:** Make sure cursor is INSIDE the SQL string (before the closing quote)!

## Expected Output

The macro will display:
- Total number of loaded sources
- Table with source details (Name, Format, Status, Label)
- Active sources highlighted in green
- Example SQL queries for each source
- Tips on using autocomplete

## Autocomplete Features

When typing SQL queries in the macro editor:

1. **Trigger**: Type `FROM ` (with space) or press `Ctrl+Space`
2. **Suggestions**: See all loaded data sources
3. **Details**: View format (CSV, XLSX, etc.) and active status
4. **Priority**: Active sources appear first
5. **Quick insert**: Press Enter to insert the quoted table name

## Use Cases

- **Dynamic queries**: Build queries based on available sources
- **Validation**: Check if required files are loaded before running queries
- **Multi-source operations**: Process data from multiple files
- **Error prevention**: Avoid typos in table names with autocomplete
