# OnlyOffice Macros Plugin - Examples Directory

This directory contains example macros organized by API category. Examples are automatically discovered and loaded by the plugin using a pre-generated manifest.

## 📁 Directory Structure

```
examples/
├── engine-test.md              # Root-level example files
├── error-examples.md
├── parameterized-macro.md
├── spreadsheet-example.md
├── syntax-error-test.md
├── syntax-test.md
├── wasm-document-snapshot-and-export.md
├── Api/                        # API categories
│   └── Methods/
│       ├── GetActiveSheet.md
│       ├── Save.md
│       └── ... (60 more files)
├── ApiRange/
│   └── Methods/
│       ├── GetValue.md
│       ├── SetValue.md
│       └── ... (65 more files)
├── ApiWorksheetFunction/
│   └── Methods/
│       ├── SUM.md
│       ├── AVERAGE.md
│       └── ... (414 more files)
└── ... (36 more categories)
```

**Total:** 39 categories, 1080 files, 2.36 MB

## 🔄 Adding New Examples

### 1. Create Your Example File

Add a new `.md` file in the appropriate category's `Methods/` folder:

```bash
# Example: Add a new VLOOKUP example
examples/ApiWorksheetFunction/Methods/VLOOKUP_Advanced.md
```

### 2. Regenerate the Manifest

**Option A: Using the shell script (Recommended)**
```bash
cd modules/macros_ide/resources/examples
./regenerate-manifest.sh
```

**Option B: Using npm**
```bash
cd /path/to/plugin/root
npm run generate-manifest
```

**Option C: Manual Node.js**
```bash
cd /path/to/plugin/root
node modules/macros_ide/scripts/tools/generate-examples-manifest.js
```

### 3. Restart OnlyOffice

Close and reopen OnlyOffice to reload the plugin with the updated manifest.

### 4. Verify

- Open the Macros plugin
- Navigate to Examples section
- Check that your new example appears in the tree

## 📋 Manifest File

The manifest is auto-generated at:
```
modules/macros_ide/resources/examples-manifest.json
```

**Contents:**
- `version`: Manifest format version (1.0.0)
- `generatedAt`: ISO timestamp of generation
- `structure`: Full directory tree (hierarchical)
- `categories`: Flat map `{ categoryName: [files] }` for quick lookup
- `statistics`: Total categories, files, size metrics

**Example structure:**
```json
{
  "version": "1.0.0",
  "generatedAt": "2025-01-20T12:34:56.789Z",
  "structure": {
    "type": "directory",
    "name": "examples",
    "children": [...]
  },
  "categories": {
    "Api": [
      { "name": "GetActiveSheet.md", "path": "Api/Methods/GetActiveSheet.md", "size": 1234 }
    ],
    "ApiRange": [...]
  },
  "statistics": {
    "totalCategories": 39,
    "totalFiles": 1080,
    "totalSize": 2421862
  }
}
```

## 🔧 Maintenance

### When to Regenerate

Regenerate the manifest whenever you:
- ✅ Add new example files
- ✅ Remove example files
- ✅ Rename example files
- ✅ Create new category directories
- ✅ Move files between categories

### Automated Regeneration

You can add the regeneration script to your development workflow:

**Pre-commit hook:**
```bash
# .git/hooks/pre-commit
#!/bin/bash
cd modules/macros_ide/resources/examples
./regenerate-manifest.sh
git add ../examples-manifest.json
```

**Cron job (daily):**
```bash
# Run daily at 3 AM
0 3 * * * cd /path/to/plugin/examples && ./regenerate-manifest.sh
```

## 🚀 Performance

**Manifest-Based Loading (METHOD 1):**
- ✅ All 39 categories discovered automatically (vs 8 hardcoded)
- ✅ 1080 files indexed (vs ~200 hardcoded)
- ✅ Instant category lookup (no sequential HEAD requests)
- ✅ No hardcoded file lists to maintain
- ✅ Zero user interaction (seamless UX)

**Previous Whitelist Approach:**
- ❌ Only 8 categories (hardcoded)
- ❌ ~200 files (hardcoded lists)
- ❌ Sequential validation (slow)
- ❌ 110+ lines of hardcoded arrays
- ❌ Manual maintenance required

## 📝 File Format

Example files are Markdown (`.md`) containing JavaScript code blocks:

```markdown
/**
 * Example: Get Active Sheet
 *
 * This example demonstrates how to get the currently active worksheet.
 */

(function() {
  var oWorksheet = Api.GetActiveSheet();
  var sName = oWorksheet.GetName();

  oWorksheet.GetRange("A1").SetValue("Active sheet: " + sName);
})();
```

## 🐛 Troubleshooting

### Manifest not updating
1. Check Node.js is installed (`node --version`)
2. Verify script has execute permissions (`ls -la regenerate-manifest.sh`)
3. Check generator script exists at `../../scripts/tools/generate-examples-manifest.js`
4. Run with `bash regenerate-manifest.sh` for verbose output

### Examples not appearing in plugin
1. Verify manifest was regenerated (check `generatedAt` timestamp)
2. Check browser console for loading errors (F12 → Console)
3. Verify `macro-manager.js` loads manifest correctly
4. Restart OnlyOffice to reload plugin

### Permission errors
```bash
chmod +x regenerate-manifest.sh
```

## 📚 Related Documentation

- **Plugin Architecture:** `/CODE_STANDARD.MD` - Section 2.2 (Manager Layer)
- **Manifest Generator:** `/modules/macros_ide/scripts/tools/generate-examples-manifest.js`
- **Macro Manager:** `/modules/macros_ide/scripts/managers/macro-manager.js`
- **Package Scripts:** `/package.json` - `generate-manifest` command
 - **Universal Data SDK & Document Export:** `.memory_bank/adr/archive/ADR-085-export-large-tables-to-document-from-wasm.md` (once promoted)

## 🎯 Quick Reference

| Task | Command |
|------|---------|
| Add new example | Create `.md` in `CategoryName/Methods/` |
| Regenerate manifest | `./regenerate-manifest.sh` |
| Check statistics | Look for "📊 Statistics:" in output |
| Verify manifest | `cat ../examples-manifest.json \| jq .statistics` |
| Test in plugin | Restart OnlyOffice, open Macros plugin |

---

**Last Updated:** 2025-01-20
**Manifest Version:** 1.0.0
**Total Examples:** 1080 files across 39 categories
