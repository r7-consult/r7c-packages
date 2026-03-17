# Api.SqlNotebook - SQLite Notebook Integration API

## Overview / Обзор

The `Api.SqlNotebook` namespace provides seamless integration between OnlyOffice macros and SQLite Notebook, allowing you to execute SQL queries against loaded data sources directly from your spreadsheet macros.

API `Api.SqlNotebook` обеспечивает бесшовную интеграцию между макросами OnlyOffice и SQLite Notebook, позволяя выполнять SQL запросы к загруженным источникам данных непосредственно из макросов.

## Architecture / Архитектура

### How It Works / Как это работает

The Api.SqlNotebook uses a **pre-execution** pattern:

1. **Macro Parsing**: When you run a macro containing `Api.SqlNotebook` calls, the Macro Manager first parses the code
2. **Pre-execution**: SQL queries are executed in the **plugin context** (where SQLite Notebook backend is available)
3. **Result Inlining**: Query results are converted to JSON and **inlined** into the macro code
4. **Document Execution**: The modified code (with results) is then executed in the document context

**Пример трансформации:**

```javascript
// BEFORE pre-execution (your code):
var result = Api.SqlNotebook.query('SELECT * FROM "data.csv" LIMIT 5', 'data.csv');

// AFTER pre-execution (what actually runs in document):
var result = {"ok":true,"result":{"columns":["id","name"],"rows":[[1,"Alice"],[2,"Bob"]],"rowCount":2,"columnCount":2}};
```

### Why Pre-execution? / Почему pre-execution?

OnlyOffice macros execute in the **document context**, which does not have access to plugin-specific objects like `SQLiteWASMBackend`. By pre-executing queries in the plugin context and inlining results, we achieve:

- ✅ **Synchronous API**: No callbacks or promises needed
- ✅ **No bridge complexity**: No localStorage/postMessage hacks
- ✅ **High performance**: Direct WASM SQLite access
- ✅ **Error handling**: Errors are captured and included in results

**Limitations:**
- ❌ Dynamic SQL queries (where SQL is constructed at runtime) won't work
- ❌ Dynamic URLs/filenames in Api.SqlNotebook calls won't be pre-executed (use literal strings)
- ✅ Static SQL queries (hardcoded strings) work perfectly

## Methods / Методы

### 1. Api.SqlNotebook.query()

Execute a single SQL query against a loaded file.

**Syntax:**
```javascript
Api.SqlNotebook.query(sql, filename, options)
```

**Parameters:**
- `sql` (string): SQL query to execute
- `filename` (string): Name of the loaded file (e.g., "MOCK_DATA.csv")
- `options` (object, optional): Additional options
  - `autoAttach` (boolean, default: true): Automatically activate the source if not active
  - `paging` (object, optional): Auto paging settings `{ auto, page, pageSize }` (applies to SELECT/WITH without LIMIT)

**Returns:**
```javascript
{
  ok: true,
  result: {
    columns: ["col1", "col2", ...],
    rows: [[val1, val2, ...], ...],
    rowCount: number,
    columnCount: number
  }
}
// OR on error:
{
  ok: false,
  error: {
    code: "ERROR_CODE",
    message: "Error description"
  }
}
```

**Example:** See [query.md](./query.md)

---

### 2. Api.SqlNotebook.batch()

Execute multiple SQL queries at once.

**Syntax:**
```javascript
Api.SqlNotebook.batch(queries)
```

**Parameters:**
- `queries` (Array): Array of query objects
  ```javascript
  [
    { sql: "SELECT ...", filename: "file1.csv", options: {...} },
    { sql: "SELECT ...", filename: "file2.csv" }
  ]
  ```

**Returns:**
```javascript
{
  ok: true,
  results: [
    { ok: true, result: {...} },  // Query 1 result
    { ok: false, error: {...} },  // Query 2 failed
    ...
  ]
}
```

**Example:** See [batch.md](./batch.md)

---

### 3. Api.SqlNotebook.getSources()

Get list of all loaded data sources in SQLite Notebook.

**Syntax:**
```javascript
Api.SqlNotebook.getSources()
```

**Returns:**
```javascript
{
  ok: true,
  sources: [
    {
      name: "MOCK_DATA.csv",
      format: "csv",
      active: true,
      label: "MOCK_DATA.csv"
    },
    ...
  ]
}
```

**Example:** See [getSources.md](./getSources.md)

---

### 4. Api.SqlNotebook.loadFromUrl()

Load a file from URL into SQLite Notebook (available immediately for queries).
If the backend is not loaded yet, `loadFromUrl` will attempt to initialize SQLite Notebook automatically.
Supports local Windows paths (`C:\\...` or `file:///C:/...`) using Node `fs` when available, or an `XMLHttpRequest` `file://` fallback in desktop mode.

**Syntax:**
```javascript
Api.SqlNotebook.loadFromUrl(url, options)
```

**Parameters:**
- `url` (string): Source URL or local Windows path (`C:\\...` / `file:///C:/...`)
- `options` (object, optional): Additional options
  - `fileName` (string): Override source name
  - `format` (number|string): Format id or string key (csv, xlsx, json, etc.)
  - `attach` (boolean, default: true): Attach to active workbook if available
  - `forceNewWorkbook` (boolean, default: false): Force creating a new workbook source
  - `delimiter` (string, default: ','): CSV delimiter
  - `hasHeaderRow` (boolean, default: true): CSV header row flag
  - `fetch` (object): Fetch options (`headers`, `credentials`, `mode`)

**Returns:**
```javascript
{
  ok: true,
  result: {
    mode: "attached" | "loaded",
    sourceId: "source-id",
    label: "file.csv",
    format: 1,
    bytes: 12345
  }
}
// OR on error:
{
  ok: false,
  error: { code: "ERROR_CODE", message: "Error description" }
}
```

**Example:** See [loadFromUrl.md](./loadFromUrl.md)

---

### 5. Api.SqlNotebook.getSchema()

Get schema metadata for loaded sources.

**Syntax:**
```javascript
Api.SqlNotebook.getSchema()
```

**Returns:**
```javascript
{
  ok: true,
  databases: [
    {
      name: "source",
      attached: true,
      schemas: [
        { name: "main", tables: [{ name: "table", rowCount: 123 }] }
      ]
    }
  ]
}
```

---

### 6. Api.SqlNotebook.describeDataset()

Get column metadata for a dataset.

**Syntax:**
```javascript
Api.SqlNotebook.describeDataset(name)
```

**Returns:**
```javascript
{
  ok: true,
  columns: [{ name: "col", type: "TEXT", nullable: true }],
  rowCount: 123
}
```

---

## Autocomplete Feature / Функция автодополнения

The Monaco editor provides intelligent autocomplete for table names after the `FROM` keyword.

### How to Use:
1. Type `Api.SqlNotebook.query('SELECT * FROM `
2. Press **Ctrl+Space**
3. See a list of all loaded sources with their formats and active status
4. Select a source to auto-insert the quoted table name

### Features:
- ✅ Shows all loaded files from SQLite Notebook
- ✅ Displays file format (CSV, XLSX, etc.)
- ✅ Indicates active/inactive status
- ✅ Active sources appear first
- ✅ Works in both `.query()` and `.batch()` methods

---

## Error Handling / Обработка ошибок

### Common Error Codes:

| Code | Description | Solution |
|------|-------------|----------|
| `EXTENSION_NOT_LOADED` | SQLite Notebook extension not available | Load SQLite Notebook plugin first |
| `BACKEND_NOT_AVAILABLE` | SQLite backend not loaded | Open SQLite Notebook tab |
| `FILE_NOT_FOUND` | File not loaded in SQLite Notebook | Load the file first |
| `INVALID_SQL` | SQL query is empty or invalid | Check SQL syntax |
| `SQL_EXECUTION_ERROR` | SQL execution failed | Check query syntax and table names |
| `NO_FILES_LOADED` | No files loaded in SQLite Notebook | Load at least one file |
| `INVALID_URL` | URL is empty or invalid | Provide a valid URL for loadFromUrl |
| `FETCH_ERROR` | Failed to fetch URL | Check URL availability and CORS |
| `LOAD_FAILED` | Failed to load file into SQLite Notebook | Verify file format/support |
| `ATTACH_FAILED` | Failed to attach to active workbook | Ensure an active workbook is available |
| `LOCAL_FILE_NOT_SUPPORTED` | Local file access not available | Use desktop build or `file:///` access |
| `LOCAL_FILE_READ_ERROR` | Failed to read local file | Check path and permissions |

### Best Practices:

```javascript
(function() {
    // Always check result.ok before using data
    var result = Api.SqlNotebook.query('SELECT * FROM "data.csv"', 'data.csv');

    if (!result.ok) {
        // Handle error
        var sheet = Api.GetActiveSheet();
        sheet.GetRange("A1").SetValue("ERROR: " + result.error.message);
        return;
    }

    // Use result.result safely
    var data = result.result;
    // ... process data
})();
```

---

## Complete Examples / Полные примеры

### Example 1: Simple Data Export
```javascript
(function() {
    var oWorksheet = Api.GetActiveSheet();
    var result = Api.SqlNotebook.query(
        'SELECT * FROM "MOCK_DATA.csv" LIMIT 10',
        'MOCK_DATA.csv'
    );

    if (result.ok) {
        var data = result.result;

        // Headers
        for (var i = 0; i < data.columns.length; i++) {
            oWorksheet.GetRange(String.fromCharCode(65 + i) + "1").SetValue(data.columns[i]);
        }

        // Data rows
        for (var row = 0; row < data.rows.length; row++) {
            for (var col = 0; col < data.rows[row].length; col++) {
                oWorksheet.GetRange(String.fromCharCode(65 + col) + (row + 2)).SetValue(data.rows[row][col]);
            }
        }
    }
})();
```

### Example 2: Data Analysis with Multiple Queries
```javascript
(function() {
    var batchResult = Api.SqlNotebook.batch([
        { sql: 'SELECT COUNT(*) as total FROM "data.csv"', filename: 'data.csv' },
        { sql: 'SELECT AVG(price) as avg_price FROM "data.csv"', filename: 'data.csv' },
        { sql: 'SELECT category, COUNT(*) FROM "data.csv" GROUP BY category', filename: 'data.csv' }
    ]);

    if (batchResult.ok) {
        var sheet = Api.GetActiveSheet();
        var row = 1;

        batchResult.results.forEach(function(result, index) {
            if (result.ok) {
                sheet.GetRange("A" + row).SetValue("Query " + (index + 1) + ":");
                row++;
                // ... display results
            }
        });
    }
})();
```

### Example 3: Dynamic Source List
```javascript
(function() {
    var sourcesResult = Api.SqlNotebook.getSources();

    if (sourcesResult.ok && sourcesResult.sources.length > 0) {
        var sheet = Api.GetActiveSheet();

        // Process each available source
        sourcesResult.sources.forEach(function(source, index) {
            if (source.active) {
                var query = 'SELECT * FROM "' + source.name + '" LIMIT 5';
                var result = Api.SqlNotebook.query(query, source.name);

                if (result.ok) {
                    // Export first 5 rows from each source
                    // ... display data
                }
            }
        });
    }
})();
```

---

## Technical Details / Технические детали

### Implementation Files:

1. **API Definition**: `modules/macros_ide/libs/cell/api.js` (lines 10710-10811)
   - Defines `ApiInterface.prototype.SqlNotebook` namespace
   - Delegates to `SqlNotebookApiExtension`

2. **Extension Backend**: `modules/macros_ide/scripts/api/sqlite-notebook-api-extension.js`
   - Implements `query()`, `batch()`, `getSources()` methods
   - Interfaces with `SQLiteWASMBackend`

3. **Pre-execution Logic**: `modules/macros_ide/scripts/managers/macro-manager.js` (#injectSqlNotebookSupport)
   - Parses macro code for `Api.SqlNotebook` calls
   - Executes queries in plugin context
   - Inlines JSON results

4. **Autocomplete Provider**: `modules/macros_ide/scripts/services/sql-completion-provider.js`
   - Monaco completion provider
   - Suggests table names after FROM keyword

### Prerequisites:
- SQLite Notebook backend must be available (loadFromUrl tries to initialize it automatically)
- At least one data source must be loaded in SQLite Notebook (or use `loadFromUrl`)
- Files referenced in queries must match loaded source names exactly
- Local path loading requires Node `fs` (desktop app)

---

## FAQ

**Q: Can I use variables in SQL queries?**
A: No, SQL queries must be static strings. Use string concatenation before the API call:
```javascript
// ❌ Won't work (dynamic):
var tableName = "MOCK_DATA.csv";
var result = Api.SqlNotebook.query('SELECT * FROM "' + tableName + '"', tableName);

// ✅ Works (static):
var result = Api.SqlNotebook.query('SELECT * FROM "MOCK_DATA.csv"', 'MOCK_DATA.csv');
```

**Q: How do I know which files are loaded?**
A: Use `Api.SqlNotebook.getSources()` to get a list of all loaded sources.

**Q: describeDataset returns empty columns/rowCount. Why?**
A: It expects the actual table name. Use `getSchema()` to read the real table name and pass it to `describeDataset()`.

**Q: What SQL syntax is supported?**
A: Standard SQLite syntax. See [SQLite documentation](https://www.sqlite.org/lang.html).

**Q: Can I JOIN multiple files?**
A: Yes! Each loaded file is a separate table:
```javascript
var result = Api.SqlNotebook.query(
    'SELECT a.*, b.price FROM "customers.csv" a JOIN "orders.csv" b ON a.id = b.customer_id',
    'customers.csv'
);
```

**Q: Does it support INSERT/UPDATE/DELETE?**
A: Yes, but changes are not persisted to the original files. They exist only in the SQLite session.

---

## See Also / См. также

- [Query Method Example](./query.md)
- [Batch Method Example](./batch.md)
- [GetSources Method Example](./getSources.md)
- [LoadFromUrl Method Example](./loadFromUrl.md)
- [SQLite Syntax Reference](https://www.sqlite.org/lang.html)
- [SQLite Notebook Plugin Documentation](../../../../../../README.md)

---

## Support / Поддержка

For issues or questions:
- GitHub: https://github.com/r7-consult/r7-code
- Email: support@r7-consult.ru
- Documentation: https://r7-consult.ru/docs

---

**Version**: 1.0.0
**Last Updated**: 2026-02-04
**Author**: R7-Consult
**License**: Apache 2.0
