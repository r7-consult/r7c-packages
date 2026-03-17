# SQLite Notebook - Autocomplete Test

Quick test for SQL table name autocomplete feature.

## How to Test

1. **Load files in SQLite Notebook** (e.g., MOCK_DATA.csv)
2. **Create a new macro** in Macros tab
3. **Copy this code:**

```javascript
(function() {
    // TEST 1: Basic autocomplete
    // Position cursor after "FROM " (with space) and press Ctrl+Space
    var result1 = Api.SqlNotebook.query('SELECT * FROM ', 'MOCK_DATA.csv');

    // TEST 2: Start typing table name
    // After typing first letter, Monaco should show matching sources
    var result2 = Api.SqlNotebook.query('SELECT * FROM M', 'MOCK_DATA.csv');

    // TEST 3: Batch query autocomplete
    var batchResult = Api.SqlNotebook.batch([
        { sql: 'SELECT * FROM ', filename: 'MOCK_DATA.csv' }
    ]);

    // After testing autocomplete, you can run this to verify getSources works:
    var sources = Api.SqlNotebook.getSources();
    if (sources.ok) {
        var sheet = Api.GetActiveSheet();
        sheet.GetRange("A1").SetValue("Available sources:");
        sources.sources.forEach(function(source, index) {
            sheet.GetRange("A" + (index + 2)).SetValue(source.name + " (" + source.format + ")");
        });
    }
})();
```

## Testing Steps

### Step 1: Check Console
Open browser DevTools (F12) and check Console tab for these messages:
- `[SqlCompletionProvider] Initialized`
- `[MonacoEditor] SQL Completion Provider registered successfully`

### Step 2: Trigger Autocomplete
1. Position cursor after `FROM ` (with trailing space)
2. Press **Ctrl+Space** or **Ctrl+Shift+Space**
3. You should see a dropdown with loaded table names

### Step 3: Debug (if not working)
Check console for debug messages:
- `[SqlCompletionProvider] Triggered at position...`
- `[SqlCompletionProvider] In SQL context: true`
- `[SqlCompletionProvider] After FROM keyword: true`
- `[SqlCompletionProvider] Loaded sources: X`

### Common Issues

**Issue 1: Autocomplete not triggering**
- Make sure cursor is INSIDE the SQL string (before closing quote)
- Try pressing Ctrl+Space explicitly
- Check if SQLite Notebook has loaded files

**Issue 2: No sources shown**
- Load at least one file in SQLite Notebook first
- Check console: `Api.SqlNotebook.getSources()` should return data

**Issue 3: Provider not registered**
- Check console for: `[SqlCompletionProvider] Initialized`
- Reload page if needed

## Manual Test Commands

Run these in browser console (F12):

```javascript
// Check if provider is loaded
console.log('Provider:', window.SqlCompletionProvider);

// Check if sources are available
console.log('Sources:', window.SqlNotebookApiExtension.getSources());

// Test provider directly
if (window.SqlCompletionProvider) {
    console.log('Provider registered:', window.SqlCompletionProvider._isRegistered);
}
```

## Expected Behavior

When working correctly:
1. Type `FROM ` in SQL query
2. Press Ctrl+Space
3. See dropdown with:
   - Table names from loaded files
   - Format indicators (CSV, XLSX, etc.)
   - Active status (✅ for active sources)
   - Active sources listed first

## Tips

- Autocomplete works best with **space after FROM**: `FROM `
- Works in both `.query()` and `.batch()` methods
- Reloads source list dynamically when files are added/removed
- Press Escape to dismiss suggestions
