/**
 * WASM + CurrentDocumentAPI + DocumentExportAPI Example
 *
 * Scenario:
 * 1. Read the current spreadsheet document into a snapshot object
 *    using CurrentDocumentAPI.getSnapshot().
 * 2. Materialise the first sheet as a virtual table in the WASM
 *    SQLite backend (one row per spreadsheet row, columns A..N).
 * 3. Run a simple aggregation query in WASM (COUNT(*) per column)
 *    to demonstrate compute on the snapshot.
 * 4. Export the aggregated result back into the active document
 *    as a named table via DocumentExportAPI.exportResultToDocument().
 *
 * Requirements:
 * - ENABLE_DOCUMENT_EXPORT_API = true
 * - WASM data service available (SQLiteBackend / SQLiteWASMBackend)
 * - CurrentDocumentAPI and DocumentExportAPI scripts loaded
 */

(function () {
  if (!window.CurrentDocumentAPI || typeof window.CurrentDocumentAPI.getSnapshot !== 'function') {
    Api.ShowMessage('WASM Snapshot Macro', 'CurrentDocumentAPI is not available.')
    return
  }
  if (!window.DocumentExportAPI || typeof window.DocumentExportAPI.exportResultToDocument !== 'function') {
    Api.ShowMessage('WASM Snapshot Macro', 'DocumentExportAPI is not available.')
    return
  }
  if (!window.R7DataService) {
    Api.ShowMessage('WASM Snapshot Macro', 'R7DataService (WASM data service) is not available.')
    return
  }

  var svc = window.R7DataService

  function log (msg) {
    try {
      console.log('[WASM Snapshot Macro]', msg)
    } catch (_) {}
  }

  function ensureBackend () {
    return svc.ensureBackend()
  }

  function asSqlIdentifier (name) {
    if (!name) return 'sheet1'
    return String(name).replace(/[^a-zA-Z0-9_]/g, '_') || 'sheet1'
  }

  function escapeIdent (name) {
    return '"' + String(name).replace(/"/g, '""') + '"'
  }

  function buildInsertStatements (tableName, headerRow, dataRows) {
    var stmts = []
    var cols = headerRow || []
    var colNames = []
    for (var i = 0; i < cols.length; i++) {
      var c = cols[i]
      var name = c || ('col_' + (i + 1))
      colNames.push(escapeIdent(name))
    }
    var columnsList = colNames.join(', ')

    for (var r = 0; r < dataRows.length; r++) {
      var row = dataRows[r] || []
      var values = []
      for (var j = 0; j < colNames.length; j++) {
        var v = row[j]
        if (v == null) {
          values.push('NULL')
        } else {
          var s = String(v)
          var n = Number(s)
          if (!isNaN(n) && s.trim() !== '') {
            values.push(s)
          } else {
            values.push("'" + s.replace(/'/g, "''") + "'")
          }
        }
      }
      var sql = 'INSERT INTO ' + escapeIdent(tableName) + ' (' + columnsList + ') VALUES (' + values.join(', ') + ');'
      stmts.push(sql)
    }
    return stmts
  }

  ensureBackend().then(function (ok) {
    if (!ok) {
      Api.ShowMessage('WASM Snapshot Macro', 'WASM backend is not available.')
      return
    }

    log('Fetching current document snapshot…')
    return window.CurrentDocumentAPI.getSnapshot()
  }).then(function (res) {
    if (!res || res.ok === false) {
      Api.ShowMessage('WASM Snapshot Macro', 'Snapshot failed: ' + (res && res.error && res.error.message ? res.error.message : 'Unknown error'))
      return
    }

    var snapshot = res.snapshot
    if (!snapshot || !snapshot.data || !snapshot.data.length) {
      Api.ShowMessage('WASM Snapshot Macro', 'No sheet data available in snapshot.')
      return
    }

    var firstSheet = snapshot.data[0]
    var sheetName = firstSheet[0]
    var rows = firstSheet[1] || []
    if (!rows.length) {
      Api.ShowMessage('WASM Snapshot Macro', 'First sheet "' + sheetName + '" has no data.')
      return
    }

    var headerRow = rows[0]
    var dataRows = rows.slice(1)
    var tableName = asSqlIdentifier(sheetName)

    log('Creating in‑memory table from sheet: ' + sheetName)
    var createSql = 'CREATE TABLE ' + escapeIdent(tableName) + ' (' +
      headerRow.map(function (name, idx) {
        var colName = name || ('col_' + (idx + 1))
        return escapeIdent(colName) + ' TEXT'
      }).join(', ') +
      ');'

    var statements = [createSql].concat(buildInsertStatements(tableName, headerRow, dataRows))

    function execNext (index) {
      if (index >= statements.length) {
        return Promise.resolve()
      }
      var sql = statements[index]
      return svc.runQuery(sql, { paging: { auto: false } }).then(function () {
        return execNext(index + 1)
      })
    }

    return execNext(0).then(function () {
      log('Running aggregation query in WASM…')
      var aggSql = 'SELECT ' +
        headerRow.map(function (name, idx) {
          var colName = name || ('col_' + (idx + 1))
          var ident = escapeIdent(colName)
          return "'" + colName.replace(/'/g, "''") + "' AS column_name, COUNT('x') AS row_count FROM ' + " +
            escapeIdent(tableName)
        })[0]

      // Simplified: count total rows in the table
      aggSql = 'SELECT COUNT(*) AS row_count FROM ' + escapeIdent(tableName) + ';'
      return svc.runQuery(aggSql, { paging: { auto: false } })
    }).then(function (aggRes) {
      if (!aggRes || aggRes.ok === false) {
        var msg = aggRes && aggRes.error && aggRes.error.message ? aggRes.error.message : 'Unknown aggregation error'
        Api.ShowMessage('WASM Snapshot Macro', 'Aggregation failed: ' + msg)
        return
      }

      var exportPayload = {
        ok: true,
        result: {
          columns: ['metric', 'value'],
          rows: [
            ['sheet_name', sheetName],
            ['rows_in_sheet', (dataRows && dataRows.length) || 0],
            ['rows_in_wasm_table', (aggRes.result && aggRes.result.rows && aggRes.result.rows[0] && aggRes.result.rows[0][0]) || 0]
          ]
        }
      }

      log('Exporting aggregation result back to document…')
      return window.DocumentExportAPI.exportResultToDocument(exportPayload, {
        tableName: 'wasm_snapshot_metrics'
      })
    }).then(function (exportRes) {
      if (!exportRes) return
      if (exportRes.ok === false) {
        var msg = exportRes.error && exportRes.error.message ? exportRes.error.message : 'Unknown export error'
        Api.ShowMessage('WASM Snapshot Macro', 'Export failed: ' + msg)
      } else {
        Api.ShowMessage('WASM Snapshot Macro', 'Snapshot metrics exported to table "wasm_snapshot_metrics".')
      }
    }).catch(function (error) {
      Api.ShowMessage('WASM Snapshot Macro', 'Unexpected error: ' + (error && error.message ? error.message : String(error)))
    })
  }).catch(function (error) {
    Api.ShowMessage('WASM Snapshot Macro', 'Unexpected error: ' + (error && error.message ? error.message : String(error)))
  })
})();

