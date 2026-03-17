# Data Dashboard Analyzer - OnlyOffice Macro

A comprehensive OnlyOffice macro that demonstrates advanced use of the OnlyOffice UI SDK for spreadsheet data analysis and interactive dashboard creation.

## Overview

This macro transforms OnlyOffice Spreadsheets into a powerful data analysis platform with:
- Interactive data exploration and filtering
- Real-time statistical analysis
- Professional dashboard interface
- Data export and sharing capabilities
- Advanced SDK component integration

## Features

### 📊 **Data Analysis Capabilities**
- **Range Detection**: Auto-detect data ranges or manual selection
- **Dynamic Filtering**: Multi-column filtering with visual feedback
- **Statistical Analysis**: Summary statistics, trends, distribution, and correlation
- **Real-time Updates**: Instant dashboard updates as filters change

### 🛠 **Advanced SDK Components**
- **Table Component**: Interactive data table with sorting and selection
- **Modal Dialogs**: Export options, data sharing, and detailed views
- **State Management**: Persistent application state across operations
- **Event System**: Component communication and workflow coordination
- **Security**: Input validation and sanitization

### 🔧 **OnlyOffice Integration**
- **Spreadsheet API**: Direct reading from cells and ranges
- **Selection Tracking**: Automatic range detection from user selection
- **Document Events**: Response to selection changes and content updates

## Installation

### 1. Copy Files
Copy the entire `02-data-dashboard` directory to your OnlyOffice plugins directory.

### 2. SDK Dependencies
Ensure OnlyOffice UI SDK files are available:
```
├── sdk/
│   ├── onlyoffice-ui-sdk-plugin.js
│   └── onlyoffice-ui-sdk-bundle.js
```

### 3. Spreadsheet Setup
- Open OnlyOffice Spreadsheet Editor
- Prepare data with headers in the first row
- Launch the macro from Plugins → Data Dashboard Analyzer

## Usage

### 1. Load Data
- **Auto-Detection**: Click "Auto-Detect Range" to find data automatically
- **Manual Selection**: Select cells in spreadsheet and click "Load Data"
- **Range Input**: Type range (e.g., "A1:D10") and click "Load Data"

### 2. Filter Data
- Select column from dropdown
- Enter filter value
- Click "Add Filter" to apply
- View active filters as chips with removal option

### 3. Analyze Data
- Choose analysis type (Summary, Trends, Distribution, Correlation)
- Click "Run Analysis" to generate insights
- View results in dashboard widgets

### 4. Export Results
- **CSV Export**: Download filtered data as CSV
- **Report Export**: Generate comprehensive analysis report
- **Share Insights**: Email key findings to stakeholders

## Architecture

### Plugin Structure
```
02-data-dashboard/
├── config.json          # OnlyOffice plugin configuration
├── index.html           # Dashboard interface and styling
├── main.js             # Application logic and analysis engine
└── README.md           # This documentation
```

### Data Flow
```
Spreadsheet → Range Detection → Data Loading → Filtering → Analysis → Dashboard
```

### SDK Integration Examples

#### Advanced State Management
```javascript
// Initialize comprehensive state tracking
sdk.state.set('rawData', []);
sdk.state.set('filteredData', []);
sdk.state.set('currentFilters', []);
sdk.state.set('analysisResults', {});

// Subscribe to state changes with complex logic
sdk.state.subscribe('currentFilters', (filters) => {
    currentFilters = filters || [];
    updateFilterChips();
    applyFilters();
});
```

#### Interactive Table Component
```javascript
// Create table with advanced features
dataTable = sdk.createTable({
    container: '#data-table-widget',
    columns: generateColumnsFromData(data),
    data: data,
    sortable: true,
    selectable: true,
    className: 'dashboard-table'
});

// Handle table interactions
sdk.on('table:row-selected', (event) => {
    showRowDetails(event.row);
});

sdk.on('table:sorted', (event) => {
    sdk.emit('dashboard:table-sorted', event);
});
```

#### Complex Modal Workflows
```javascript
// Export options modal
const modal = sdk.createModal({
    title: 'Export Analysis Report',
    content: generateExportForm(),
    closable: true,
    className: 'export-modal'
});

// Share insights modal with email functionality
const shareModal = sdk.createModal({
    title: 'Share Insights',
    content: generateShareForm(insights),
    closable: true
});
```

#### Event-Driven Architecture
```javascript
// Data loading events
sdk.emit('data:loaded', {
    range: range,
    rowCount: data.length,
    columnCount: data[0] ? Object.keys(data[0]).length : 0,
    timestamp: new Date().toISOString()
});

// Analysis completion events
sdk.emit('analysis:completed', {
    type: analysisType,
    results: results,
    dataCount: filteredData.length
});
```

## Analysis Types

### 1. Summary Statistics
- **Count**: Number of data points
- **Sum**: Total of numeric values
- **Mean**: Average values
- **Min/Max**: Range boundaries
- **Standard Deviation**: Data spread

```javascript
function computeSummaryStats(data) {
    const stats = {};
    const numericColumns = getNumericColumns(data);
    
    numericColumns.forEach(column => {
        const values = extractNumericValues(data, column);
        stats[column] = {
            count: values.length,
            mean: calculateMean(values),
            std: calculateStandardDeviation(values)
        };
    });
    
    return stats;
}
```

### 2. Trend Analysis
- **Direction**: Increasing, decreasing, stable
- **Slope**: Rate of change
- **Change**: Absolute and percentage change

### 3. Distribution Analysis
- **Unique Values**: Count of distinct values
- **Top Values**: Most frequent entries
- **Distribution Patterns**: Value frequency analysis

### 4. Correlation Analysis
- **Correlation Coefficients**: Relationships between numeric columns
- **Strength**: Weak, moderate, strong correlations
- **Direction**: Positive or negative relationships

## Advanced Features

### Dynamic Filtering System
```javascript
// Multi-condition filtering
const filterEngine = {
    operators: ['contains', 'equals', 'starts', 'ends'],
    
    applyFilter(data, filter) {
        return data.filter(row => {
            const cellValue = String(row[filter.column] || '').toLowerCase();
            const filterValue = filter.value.toLowerCase();
            
            switch (filter.operator) {
                case 'contains': return cellValue.includes(filterValue);
                case 'equals': return cellValue === filterValue;
                // ... other operators
            }
        });
    }
};
```

### Real-time Dashboard Updates
```javascript
// Automatic updates when data changes
sdk.state.subscribe('filteredData', (newData) => {
    updateDataTable(newData);
    updateSummaryStats();
    recalculateAnalysis();
});
```

### Data Export System
```javascript
function exportToCSV() {
    const csvContent = convertToCSV(filteredData);
    downloadFile(csvContent, 'dashboard-data.csv', 'text/csv');
    
    sdk.emit('data:exported', {
        format: 'csv',
        rowCount: filteredData.length
    });
}
```

## Configuration

### Analysis Settings
Customize analysis types by extending the `ANALYSIS_TYPES` object:

```javascript
ANALYSIS_TYPES['custom-analysis'] = {
    name: 'Custom Analysis',
    description: 'Your custom analysis description',
    compute: function(data) {
        // Your analysis logic
        return results;
    }
};
```

### UI Customization
- Modify CSS in `index.html` for styling changes
- Adjust dashboard layout with CSS Grid
- Customize chart placeholders with actual visualization libraries
- Add new filter operators and UI controls

### Integration Options
- Connect to external data sources
- Add chart libraries (Chart.js, D3.js)
- Implement database connectivity
- Add machine learning analysis

## Error Handling

### Data Validation
```javascript
function validateDataRange(range) {
    if (!range || !range.trim()) {
        throw new Error('Data range is required');
    }
    
    if (!isValidRange(range)) {
        throw new Error('Invalid range format');
    }
}
```

### Analysis Error Recovery
```javascript
try {
    const results = analysisConfig.compute(filteredData);
    sdk.state.set('analysisResults', results);
    showAlert('Analysis completed successfully', 'success');
} catch (error) {
    console.error('Analysis error:', error);
    showAlert('Analysis failed: ' + error.message, 'error');
}
```

### User Feedback
- Loading indicators during data processing
- Progress messages for long operations
- Error alerts with recovery suggestions
- Success confirmations for completed actions

## Performance Considerations

### Large Dataset Handling
- Implement data pagination for large datasets
- Use virtual scrolling for tables
- Lazy loading of analysis results
- Memory cleanup for destroyed components

### Optimization Strategies
```javascript
// Debounced filtering for better performance
const debouncedFilter = debounce(applyFilters, 300);

// Efficient data processing
function processLargeDataset(data) {
    const batchSize = 1000;
    const batches = chunkArray(data, batchSize);
    
    return batches.reduce((results, batch) => {
        return mergeResults(results, processBatch(batch));
    }, {});
}
```

## Extending the Macro

### Adding New Analysis Types
1. Define computation function
2. Add to `ANALYSIS_TYPES` configuration
3. Update UI to display results
4. Test with various data types

### Custom Visualizations
```javascript
// Integration with chart libraries
function createChart(data, type) {
    const chartContainer = document.getElementById('trend-chart');
    
    // Example with Chart.js (if included)
    new Chart(chartContainer, {
        type: type,
        data: formatDataForChart(data),
        options: getChartOptions(type)
    });
}
```

### Advanced Filtering
- Add date range filters
- Implement numeric range filtering
- Create regex pattern matching
- Add multi-select dropdown filters

### Data Sources
- Connect to REST APIs
- Import from external databases
- Real-time data streaming
- File upload capabilities

## Troubleshooting

### Common Issues

1. **Data Not Loading**
   - Check range format (e.g., "A1:D10")
   - Verify data contains headers
   - Ensure spreadsheet has data in specified range

2. **Analysis Errors**
   - Confirm numeric data for statistical analysis
   - Check for empty cells or invalid values
   - Verify sufficient data for chosen analysis type

3. **Performance Issues**
   - Limit data range for large datasets
   - Clear unused filters
   - Refresh browser if memory usage is high

### Debug Information
Enable detailed logging:
```javascript
// Set debug mode
sdk.state.set('debugMode', true);

// Monitor all state changes
sdk.state.subscribe('*', (key, newValue) => {
    console.log(`[State] ${key}:`, newValue);
});
```

## Best Practices

### Data Preparation
- Use clear, descriptive column headers
- Ensure consistent data types per column
- Remove or handle missing values
- Keep datasets focused and relevant

### Analysis Workflow
1. Start with summary statistics
2. Apply filters to focus analysis
3. Run specific analysis types
4. Export results for documentation
5. Share insights with stakeholders

### Performance
- Process data in manageable chunks
- Use filters to reduce dataset size
- Clear analysis results when switching datasets
- Monitor memory usage in browser

## Future Enhancements

### Planned Features
- Advanced charting with multiple visualization types
- Machine learning integration for predictive analysis
- Real-time collaboration on analysis results
- Integration with external business intelligence tools
- Advanced statistical tests and hypothesis testing

### Integration Possibilities
- Connect to Google Sheets or Excel Online
- Integration with R or Python for advanced analysis
- Export to PowerBI or Tableau
- API endpoints for external tool integration

## License

This example is part of the OnlyOffice UI SDK and is provided for educational and development purposes.

## Support

For questions about this macro or the OnlyOffice UI SDK:
- Review the main SDK documentation
- Check other example macros for patterns
- Consult OnlyOffice Spreadsheet API documentation
- Visit OnlyOffice developer community forums