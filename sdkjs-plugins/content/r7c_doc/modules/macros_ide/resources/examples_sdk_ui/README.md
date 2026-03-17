# OnlyOffice UI SDK - Complete Macro Examples

This directory contains **complete, deployable OnlyOffice macro applications** that demonstrate practical use of the OnlyOffice UI SDK in real business scenarios.

## 🎯 New Approach: Complete Business Solutions

Unlike traditional cookbook-style examples that show isolated functions, these are **complete macro plugins** that solve real business problems while demonstrating SDK components in context.

### Why This Approach?

1. **Real Business Value**: Each macro solves actual problems users face
2. **Complete Architecture**: Shows proper OnlyOffice plugin structure and lifecycle
3. **SDK in Context**: Demonstrates how components work together
4. **Production Ready**: Can be deployed and used immediately
5. **Best Practices**: Shows proper error handling, state management, and user experience

## 📁 Available Macro Examples

### 1. Document Template Generator
**Path**: `01-document-template-generator/`  
**Purpose**: Generate personalized documents from predefined templates  
**Use Case**: Business letters, proposals, invoices, reports  

**SDK Components Demonstrated**:
- ✅ Modal components for user input forms
- ✅ State management for form data persistence
- ✅ Security utilities for input validation
- ✅ Event system for workflow coordination

**OnlyOffice Integration**:
- Document API for content insertion
- Text formatting and structure
- Plugin lifecycle management

**Business Scenario**: HR department needs to generate personalized employment letters, contracts, and reports with consistent formatting and branding.

### 2. Data Dashboard Analyzer
**Path**: `02-data-dashboard/`  
**Purpose**: Analyze spreadsheet data with interactive dashboards  
**Use Case**: Data analysis, reporting, business intelligence  

**SDK Components Demonstrated**:
- ✅ Table component with sorting and filtering
- ✅ Complex state management with subscriptions
- ✅ Modal dialogs for export and sharing
- ✅ Event-driven architecture
- ✅ Advanced security validation

**OnlyOffice Integration**:
- Spreadsheet API for data reading
- Range detection and validation
- Selection tracking

**Business Scenario**: Sales team needs to analyze quarterly performance data, identify trends, and create executive reports from spreadsheet data.

### 3. Document Workflow Manager
**Path**: `03-document-workflow/`  
**Purpose**: Multi-step document review and approval processes  
**Use Case**: Document approval workflows, collaboration, tracking  

**SDK Components Demonstrated**:
- ✅ Multi-step modal workflows
- ✅ Complex state transitions
- ✅ Event coordination between components
- ✅ Progress tracking and validation

**OnlyOffice Integration**:
- Comments and tracking changes
- Document locking and permissions
- Collaborative features

**Business Scenario**: Legal team needs structured review process for contracts with multiple approval stages, comment tracking, and status management.

### 4. External System Integration
**Path**: `04-external-integration/`  
**Purpose**: Connect OnlyOffice with external APIs and systems  
**Use Case**: CRM integration, database connectivity, third-party services  

**SDK Components Demonstrated**:
- ✅ Async operation handling
- ✅ Error recovery and retry logic
- ✅ Progress indicators and user feedback
- ✅ Secure API communication

**OnlyOffice Integration**:
- Document synchronization
- External data importing
- Real-time updates

**Business Scenario**: Marketing team needs to pull customer data from CRM system and generate personalized proposals directly in OnlyOffice.

### 5. Batch Document Processor
**Path**: `05-batch-processor/`  
**Purpose**: Bulk operations on multiple documents  
**Use Case**: Document standardization, format conversion, content updates  

**SDK Components Demonstrated**:
- ✅ Progress tracking for long operations
- ✅ Error handling with detailed reporting
- ✅ Queue management and prioritization
- ✅ Batch operation patterns

**OnlyOffice Integration**:
- Multiple document handling
- Format conversion
- Bulk content operations

**Business Scenario**: Operations team needs to standardize formatting across hundreds of legacy documents and apply new branding guidelines.

## 🚀 Quick Start Guide

### 1. Choose Your Use Case
Select the macro that best matches your business need:
- **Document Generation** → Template Generator
- **Data Analysis** → Data Dashboard  
- **Approval Processes** → Workflow Manager
- **System Integration** → External Integration
- **Bulk Operations** → Batch Processor

### 2. Installation
1. Copy the macro directory to your OnlyOffice plugins folder
2. Ensure SDK files are available (see [Installation Guide](#installation))
3. Register the plugin in OnlyOffice
4. Launch from the Plugins menu

### 3. Customization
Each macro includes:
- **Configuration files** for easy customization
- **Detailed README** with architecture explanation
- **Extensible code structure** for adding features
- **Best practices** for OnlyOffice macro development

## 🛠 Installation Requirements

### SDK Dependencies
All macros require the OnlyOffice UI SDK:
```
├── sdk/
│   ├── onlyoffice-ui-sdk-plugin.js    # Main plugin registration
│   └── onlyoffice-ui-sdk-bundle.js    # Complete SDK bundle
```

### Plugin Structure
Each macro follows the standard OnlyOffice plugin structure:
```
macro-name/
├── config.json          # OnlyOffice plugin configuration
├── index.html           # User interface and styling
├── main.js              # Application logic and SDK integration
└── README.md            # Detailed documentation
```

### Browser Requirements
- Modern browser with ES6 support
- OnlyOffice 6.0.0 or later
- JavaScript enabled

## 🎓 Learning Path

### For Beginners
1. **Start with**: Document Template Generator (simplest workflow)
2. **Learn**: Basic SDK patterns and plugin lifecycle
3. **Practice**: Customize templates and add new form fields

### For Intermediate Developers
1. **Explore**: Data Dashboard Analyzer (complex state management)
2. **Learn**: Table components and event systems
3. **Practice**: Add new analysis types and visualizations

### For Advanced Developers
1. **Study**: External Integration (async patterns and error handling)
2. **Learn**: API integration and real-time updates
3. **Practice**: Connect to your own systems and databases

## 📚 SDK Components Reference

### Modal Components
```javascript
// Basic modal
const modal = OnlyOfficeUISDK.createModal({
    title: 'Modal Title',
    content: '<p>Content</p>',
    closable: true
});

// Complex workflow modal
const workflowModal = sdk.createModal({
    title: 'Multi-Step Process',
    content: generateStepContent(currentStep),
    className: 'workflow-modal'
});
```

### Table Components
```javascript
// Interactive data table
const table = sdk.createTable({
    container: '#table-container',
    columns: [
        { key: 'name', title: 'Name' },
        { key: 'value', title: 'Value' }
    ],
    data: dataArray,
    sortable: true,
    selectable: true
});
```

### State Management
```javascript
// Set and subscribe to state changes
sdk.state.set('key', value);
sdk.state.subscribe('key', (newValue, oldValue) => {
    // Handle state changes
});
```

### Event System
```javascript
// Emit and listen for events
sdk.emit('custom:event', { data: 'value' });
sdk.on('custom:event', (eventData) => {
    // Handle event
});
```

### Security Utilities
```javascript
// Input sanitization
const safeInput = OnlyOfficeUISDK.sanitizeInput(userInput);
const safeHtml = OnlyOfficeUISDK.escapeHtml(htmlContent);
```

## 🔧 Customization Guide

### Adding New Features
1. **Extend State Management**: Add new state keys for your data
2. **Create UI Components**: Use SDK components for new interfaces
3. **Implement Business Logic**: Add your specific workflow requirements
4. **Handle Events**: Coordinate between components with events

### Styling Customization
- Modify CSS in `index.html` files
- Follow OnlyOffice design patterns
- Ensure responsive design for different screen sizes
- Maintain accessibility standards

### Integration Patterns
- **Database Connection**: Use External Integration patterns
- **File Processing**: Follow Batch Processor approach
- **Real-time Updates**: Implement event-driven updates
- **Error Handling**: Use comprehensive error recovery

## 🚨 Best Practices

### Plugin Development
1. **Follow OnlyOffice Lifecycle**: Use proper init and content ready hooks
2. **Handle Errors Gracefully**: Provide user feedback and recovery options
3. **Manage State Properly**: Use SDK state management for persistence
4. **Clean Up Resources**: Destroy components when plugin closes

### User Experience
1. **Loading Indicators**: Show progress for long operations
2. **Clear Feedback**: Provide success/error messages
3. **Intuitive Navigation**: Guide users through workflows
4. **Responsive Design**: Work on different screen sizes

### Performance
1. **Lazy Loading**: Load components when needed
2. **Efficient State Updates**: Batch state changes when possible
3. **Memory Management**: Clean up event listeners and components
4. **Optimize API Calls**: Cache results and batch requests

## 🔍 Troubleshooting

### Common Issues

1. **SDK Not Available**
   ```javascript
   if (typeof OnlyOfficeUISDK === 'undefined') {
       console.error('SDK not loaded');
       // Handle gracefully
   }
   ```

2. **Plugin Lifecycle Issues**
   ```javascript
   window.Asc.plugin.init = function() {
       $(function() {
           // Ensure DOM is ready
           initializeApp();
       });
   };
   ```

3. **State Management Problems**
   ```javascript
   // Always check if SDK is available
   if (sdk && sdk.state) {
       sdk.state.set('key', value);
   }
   ```

### Debug Mode
Enable debugging in any macro:
```javascript
// Set debug flag
sdk.state.set('debugMode', true);

// Monitor events
sdk.on('*', (eventName, data) => {
    console.log('[Event]', eventName, data);
});
```

## 📈 Performance Monitoring

### Key Metrics
- **Initialization Time**: How long the macro takes to load
- **User Interaction Response**: Time from user action to UI update
- **Memory Usage**: Monitor for memory leaks in long-running sessions
- **API Response Times**: Track OnlyOffice API call performance

### Optimization Strategies
- Use efficient data structures
- Implement virtual scrolling for large datasets
- Cache computed results
- Debounce user input events

## 🤝 Contributing

### Adding New Macros
1. Create new directory following naming convention
2. Include all required files (config.json, index.html, main.js, README.md)
3. Document SDK component usage patterns
4. Provide real business scenario and use case
5. Include error handling and user feedback
6. Test thoroughly with different data sizes

### Improving Existing Macros
- Add new features while maintaining simplicity
- Improve error handling and user experience
- Optimize performance for larger datasets
- Enhance documentation with more examples
- Add integration with additional OnlyOffice features

## 📄 License

These examples are part of the OnlyOffice UI SDK and are provided for educational and development purposes.

## 🆘 Support

### Documentation
- Review individual macro README files for detailed information
- Check the main SDK integration guide at `/home/er77/_wrk24/r7-plugins-content/macros_r7/resources/MACRO_INTEGRATION_GUIDE.md`
- Consult OnlyOffice API documentation

### Community
- OnlyOffice Developer Forums
- GitHub Issues for bug reports
- Community contributions welcome

### Professional Support
- OnlyOffice Enterprise Support
- Custom macro development services
- Training and consultation available

---

## 🎯 Next Steps

1. **Choose a macro** that matches your business need
2. **Deploy and test** with your data
3. **Customize** for your specific requirements  
4. **Extend** with additional features
5. **Share** your improvements with the community

These complete macro examples demonstrate the power of combining OnlyOffice's document capabilities with the UI SDK's component system to create professional business applications.