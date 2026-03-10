# Document Template Generator - OnlyOffice Macro

A complete OnlyOffice macro application that demonstrates practical use of the OnlyOffice UI SDK to create personalized documents from predefined templates.

## Overview

This macro provides a professional document generation interface that allows users to:
- Select from multiple document templates (Business Letter, Proposal, Invoice, Report)
- Fill in template-specific information through dynamic forms
- Generate formatted documents with their data
- Validate input and provide user feedback

## Features

### 🎯 **Business-Focused Templates**
- **Business Letter**: Professional correspondence with sender/recipient details
- **Project Proposal**: Structured business proposals with objectives and scope
- **Invoice**: Professional billing documents with itemization
- **Business Report**: Formatted reports with executive summary and findings

### 🛠 **SDK Components Demonstrated**
- **Modal Components**: Input validation dialogs and user feedback
- **State Management**: Form data persistence and application state
- **Security Utilities**: Input sanitization and validation
- **Event System**: Component communication and workflow coordination

### 🔧 **OnlyOffice Integration**
- Document API integration for content insertion
- Proper plugin lifecycle management
- Error handling and user feedback
- Professional UI following OnlyOffice design patterns

## Installation

### 1. Copy Files
Copy the entire `01-document-template-generator` directory to your OnlyOffice plugins directory.

### 2. Install SDK Dependencies
Ensure the OnlyOffice UI SDK files are available:
```
├── sdk/
│   ├── onlyoffice-ui-sdk-plugin.js
│   └── onlyoffice-ui-sdk-bundle.js
```

### 3. Register Plugin
Add the plugin to your OnlyOffice installation by copying the directory to the plugins folder or registering it through the admin interface.

## Usage

### 1. Launch the Macro
- Open OnlyOffice Document Editor
- Go to Plugins → Document Template Generator
- The macro interface will open in a modal window

### 2. Select Template
- Choose from four available template types
- Each template has specific fields optimized for its use case

### 3. Fill in Information
- Complete the required fields (marked with *)
- Optional fields can be left blank
- Input is automatically validated and sanitized

### 4. Generate Document
- Click "Generate Document" to create your personalized document
- The macro will insert formatted content into the current document
- Success/error messages provide feedback on the operation

## Architecture

### Plugin Structure
```
01-document-template-generator/
├── config.json          # OnlyOffice plugin configuration
├── index.html           # User interface and styling
├── main.js             # Application logic and SDK integration
└── README.md           # This documentation
```

### Key Components

#### Template Engine
```javascript
const TEMPLATES = {
    'business-letter': {
        name: 'Business Letter',
        fields: [...],           // Form field definitions
        template: generateBusinessLetter  // Template function
    }
    // ... other templates
};
```

#### Form Generation
Dynamic forms are generated based on template requirements:
```javascript
function generateTemplateForm(templateType) {
    const template = TEMPLATES[templateType];
    // Creates form fields dynamically
    // Sets up validation and state management
}
```

#### Document Generation
Each template has a dedicated generation function:
```javascript
async function generateBusinessLetter(data) {
    const letterContent = `
        ${data.senderName}
        ${data.senderCompany}
        // ... formatted content
    `;
    insertContentToDocument(letterContent, data.documentTitle);
}
```

### SDK Integration Examples

#### State Management
```javascript
// Initialize SDK with state management
sdk = OnlyOfficeUISDK.createSDK({
    container: document.body
});

// Track form data changes
sdk.state.subscribe('formData', (newData) => {
    formData = newData;
});

// Update state when user inputs data
sdk.state.set('formData', currentData);
```

#### Modal Components
```javascript
// Validation error modal
const modal = sdk.createModal({
    title: 'Please Fix Errors',
    content: `<ul>${errors.map(e => `<li>${e}</li>`).join('')}</ul>`,
    closable: true
});
modal.open();
```

#### Input Validation
```javascript
// Sanitize user input using SDK security utilities
let value = input.value;
if (OnlyOfficeUISDK.sanitizeInput) {
    value = OnlyOfficeUISDK.sanitizeInput(value);
}
```

#### Event System
```javascript
// Emit events for document generation
sdk.emit('document:generated', {
    template: selectedTemplate,
    data: formData,
    timestamp: new Date().toISOString()
});
```

## Configuration

### Plugin Settings (config.json)
- **GUID**: Unique identifier for the plugin
- **Editors Support**: Limited to "word" for document generation
- **Modal Mode**: Opens in a modal window for focused interaction
- **Buttons**: "Generate Document" and "Cancel" actions

### Template Customization
Add new templates by extending the `TEMPLATES` object:

```javascript
TEMPLATES['custom-template'] = {
    name: 'Custom Template',
    fields: [
        { name: 'field1', label: 'Field 1', type: 'text', required: true },
        // ... more fields
    ],
    template: generateCustomTemplate
};

async function generateCustomTemplate(data) {
    // Your template generation logic
}
```

## Error Handling

### Validation
- Required field validation
- Input sanitization
- Type checking for specific fields
- User-friendly error messages

### Error Recovery
```javascript
try {
    await template.template(data);
    showSuccess('Document generated successfully!');
} catch (error) {
    console.error('Generation error:', error);
    showError('Failed to generate: ' + error.message);
}
```

### User Feedback
- Loading indicators during generation
- Success/error message display
- Validation modal dialogs
- Auto-hiding notifications

## Best Practices Demonstrated

### 1. Proper SDK Initialization
```javascript
window.Asc.plugin.init = function() {
    $(function() {
        if (typeof OnlyOfficeUISDK !== 'undefined') {
            sdk = OnlyOfficeUISDK.createSDK({
                container: document.body
            });
            initializeApp();
        }
    });
};
```

### 2. Lifecycle Management
- Plugin initialization in the correct lifecycle hooks
- Proper cleanup when closing the plugin
- State management throughout the application lifecycle

### 3. Security Best Practices
- Input sanitization using SDK utilities
- Validation of all user inputs
- Error handling without exposing sensitive information

### 4. User Experience
- Professional UI design
- Clear workflow and navigation
- Immediate feedback for user actions
- Accessible form design

## Troubleshooting

### Common Issues

1. **SDK Not Available**
   - Ensure SDK files are included in correct order
   - Check browser console for loading errors

2. **Document Generation Fails**
   - Verify OnlyOffice Document API is available
   - Check that document is in edit mode
   - Review console logs for specific errors

3. **Form Validation Issues**
   - Check required field definitions
   - Verify input sanitization is working
   - Test with different input types

### Debug Mode
Enable debugging by opening browser console and checking for:
- SDK initialization messages
- State management updates
- Event system notifications
- Error messages and stack traces

## Extending the Macro

### Adding New Templates
1. Define template structure in `TEMPLATES` object
2. Create generation function
3. Test with various input scenarios
4. Update UI if needed for template-specific features

### Enhancing UI
- Modify CSS in `index.html` for styling changes
- Add new form field types in `createFormField()`
- Implement additional validation rules
- Add progress indicators or multi-step workflows

### Integration Features
- Connect to external APIs for data
- Add file export capabilities
- Implement collaborative features
- Add template sharing functionality

## License

This example is part of the OnlyOffice UI SDK and is provided for educational and development purposes.

## Support

For questions about this macro or the OnlyOffice UI SDK:
- Review the main SDK documentation
- Check the other example macros
- Consult the OnlyOffice Developer API documentation
- Visit the OnlyOffice community forums