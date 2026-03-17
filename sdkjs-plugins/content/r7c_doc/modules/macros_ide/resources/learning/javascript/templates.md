# Macro Templates (Read-Only)

These templates are for learning and reference only. SmartDocumentation does not execute macros.

## Template: Basic Structure

```javascript
(function () {
  // Read-only example template
  const workbook = Api.GetWorkbook()
  const sheet = workbook.GetActiveSheet()

  // Example logic
  const range = sheet.GetRange("A1:C3")
  range.SetValue("Example")
})()
```

## Template: Loop Over Rows

```javascript
(function () {
  const sheet = Api.GetActiveSheet()
  const range = sheet.GetRange("A1:A10")
  const values = range.GetValue()

  for (let i = 0; i < values.length; i += 1) {
    // Example transformation
    values[i][0] = String(values[i][0] || "").trim()
  }

  range.SetValue(values)
})()
```

## Template: Create a Summary Sheet

```javascript
(function () {
  const workbook = Api.GetWorkbook()
  const summary = workbook.CreateSheet("Summary")

  summary.GetRange("A1").SetValue("Report")
  summary.GetRange("A2").SetValue(new Date().toISOString())
})()
```
