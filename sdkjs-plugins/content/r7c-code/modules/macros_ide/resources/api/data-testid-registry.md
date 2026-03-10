# Data-testid Registry (kebab-case)

## Naming rule
`<scope>-<block>-<element>[-<action|state>][-<entity-key>]`

## Static app shell
- `app-workbench`
- `activity-left-bar`
- `activity-right-bar`
- `activity-right-pane`
- `activity-right-pane-content`
- `activity-right-pane-close-btn`
- `activity-right-pane-resizer`
- `editor-tab-bar`
- `editor-breadcrumb`
- `console-panel`
- `settings-modal-overlay`
- `context-menu-root`

## Sidebar / Explorer / Tree
- `sidebar-panel-<view-id>`
- `sidebar-panel-body-<view-id>`
- `sidebar-explorer-tree-host`
- `sidebar-macros-root`
- `sidebar-macros-header`
- `sidebar-macros-toggle`
- `sidebar-macros-title`
- `sidebar-macros-count`
- `sidebar-macros-container`
- `sidebar-tree-group-<group>`
- `sidebar-tree-work-subgroup-<group>`
- `sidebar-tree-macro-item-<guid|index>`
- `sidebar-tree-example-item-<index>`
- `sidebar-tree-macro-action-btn`
- `sidebar-tree-action-<class-name>-btn`

## Activity bars / Right pane tools
- `activity-right-toggle-pane-btn`
- `activity-right-pane-ai-connection-btn`
- `activity-right-pane-ai-management-btn`
- `activity-right-pane-chat-btn`
- `activity-right-pane-ai-job-monitor-btn`
- `activity-right-pane-debug-insights-btn`
- `activity-right-picker-<select-owner>`
- `activity-right-picker-list-<select-owner>`
- `activity-right-picker-option-<select-owner>-<value>`

## Tabs / Editor / Tooltip / Split
- `editor-tab-item-<tab-key>`
- `editor-tab-title-<tab-key>`
- `editor-tab-close-btn-<tab-key>`
- `editor-breadcrumb-item-<index>-<label>`
- `editor-breadcrumb-separator-<index>`
- `editor-custom-dsl-tooltip`
- `editor-split-right-pane`

## Console / Error popup
- `console-header`
- `console-output`
- `console-save-btn`
- `console-clear-btn`
- `console-log-row-<level>-<timestamp>`
- `macro-error-popup-root`
- `macro-error-popup-item-<id>`
- `macro-error-popup-details-btn`
- `macro-error-popup-close-btn`

## Settings
- `settings-modal`
- `settings-save-btn`
- `settings-cancel-btn`
- `settings-reset-btn`
- `settings-tab-<tab>-btn`
- `settings-tab-<tab>-panel`
- `settings-theme-select`
- `settings-openai-api-key-input`
- `settings-openai-model-select`
- `settings-use-responses-checkbox`
- `settings-enable-streaming-checkbox`
- `settings-responses-desktop-checkbox`
- `settings-enable-tools-checkbox`
- `settings-translation-prompt-textarea`
- `settings-translation-prompt-counter`
- `settings-vba-decoder-engine-select`
- `settings-docs-viewer-modal`
- `settings-docs-nav-list`
- `settings-docs-loading`
- `settings-docs-content`
- `settings-dev-tools-tab-btn`
- `settings-dev-tools-categories-select`
- `settings-dev-tools-status`
- `settings-dev-tools-results-section`
- `settings-dev-tools-results-content`
- `settings-debug-flags-root`
- `settings-use-current-workbook-source-checkbox`
- `settings-sqlite-notebook-enforce-file-size-checkbox`
- `settings-sqlite-notebook-max-file-size-mb-input`
- `settings-sqlite-notebook-warn-file-size-mb-input`

## Context menu / Dialogs / Welcome popup
- `context-menu-root-target-<guid|index|none>`
- `vba-conversion-dialog-overlay`
- `vba-conversion-dialog-content`
- `vba-conversion-dialog-preview`
- `vba-conversion-dialog-vba-pane`
- `vba-conversion-dialog-vba-preview`
- `vba-conversion-dialog-js-pane`
- `vba-conversion-dialog-js-preview`
- `vba-conversion-dialog-status`
- `vba-conversion-dialog-risk-summary`
- `vba-conversion-dialog-validation-summary`
- `vba-conversion-dialog-quality-summary`
- `vba-conversion-dialog-warnings`
- `vba-conversion-dialog-help-link`
- `vba-conversion-dialog-actions`
- `vba-conversion-dialog-accept-btn`
- `vba-conversion-dialog-edit-btn`
- `vba-conversion-dialog-cancel-btn`
- `welcome-popup-overlay`
- `welcome-popup-modal`
- `welcome-popup-content`
- `welcome-popup-trial-info`
- `welcome-popup-trial-days`
- `welcome-popup-trial-deadline`
- `welcome-popup-close-btn`
- `welcome-popup-ok-btn`

## AI Job Monitor / Agent Tree
- `ai-job-monitor-root`
- `ai-job-monitor-body`
- `ai-job-monitor-filters`
- `ai-job-monitor-filter-type`
- `ai-job-monitor-filter-status`
- `ai-job-monitor-list`
- `ai-job-monitor-details`
- `ai-job-monitor-row-<job-id>`
- `ai-job-monitor-row-status`
- `ai-job-monitor-details-header`
- `ai-job-monitor-actions`
- `ai-job-monitor-cancel-btn`
- `ai-job-monitor-copy-json-btn`
- `ai-job-monitor-logs`
- `ai-job-monitor-log-row-<index>`
- `ai-job-monitor-picker-<owner>`
- `ai-job-monitor-picker-list-<owner>`
- `ai-job-monitor-picker-option-<owner>-<value>`
- `agent-tree-toolbar`
- `agent-tree-version-select`
- `agent-tree-folder-node-<folder-id>`
- `agent-tree-folder-header-<folder-id>`
- `agent-tree-agent-item-<agent-id>`

## Selenium wait policy (ready-to-use)
- Open modal/pane: `visibilityOfElementLocated([data-testid="..."])`
- Close modal/pane: `invisibilityOfElementLocated([data-testid="..."])`
- Dynamic row by key: `presenceOfElementLocated([data-testid="...-<key>"])`
- Re-rendered lists: always re-find element by `data-testid` after action
