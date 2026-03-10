/**
 * @fileoverview OnlyOffice UI SDK DataGrid Component
 * @description Flexible data grid with sorting, filtering, virtual scrolling, and column management
 * @see {@link https://api.onlyoffice.com/plugin/basic} OnlyOffice Plugin API
 * @see {@link /CODE_STANDARD.MD} Plugin Architecture Guide
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

/**
 * DataGrid Component for tabular data display
 * Provides sorting, filtering, virtual scrolling, and column management
 */
class DataGrid {
    #element = null;
    #config = {};
    #state = {
        data: [],
        columns: [],
        sortColumn: null,
        sortDirection: 'asc',
        filters: new Map(),
        selection: new Set(),
        scrollTop: 0,
        visibleRange: { start: 0, end: 0 }
    };
    #eventSystem = null;
    #stateManager = null;
    #virtualScrolling = true;
    #rowHeight = 32;
    #headerHeight = 40;
    #visibleRows = 20;

    /**
     * Creates a new DataGrid instance
     * @param {Object} options - Component options
     * @param {HTMLElement|string} options.container - Container element or selector
     * @param {Array} [options.columns=[]] - Column definitions
     * @param {Array} [options.data=[]] - Initial data
     * @param {boolean} [options.virtualScrolling=true] - Enable virtual scrolling
     * @param {boolean} [options.sortable=true] - Enable sorting
     * @param {boolean} [options.filterable=true] - Enable filtering
     * @param {boolean} [options.selectable=true] - Enable row selection
     * @param {number} [options.rowHeight=32] - Row height in pixels
     * @param {Object} options.eventSystem - Event system instance
     * @param {Object} options.stateManager - State manager instance
     */
    constructor(options = {}) {
        this.#config = {
            virtualScrolling: true,
            sortable: true,
            filterable: true,
            selectable: true,
            multiSelect: false,
            rowHeight: 32,
            headerHeight: 40,
            bufferSize: 5,
            ...options
        };

        this.#eventSystem = options.eventSystem;
        this.#stateManager = options.stateManager;
        this.#rowHeight = this.#config.rowHeight;
        this.#headerHeight = this.#config.headerHeight;
        this.#virtualScrolling = this.#config.virtualScrolling;

        this.#setupContainer(options.container);
        this.#bindEvents();
    }

    /**
     * Initializes the DataGrid component
     * @returns {Promise<void>}
     */
    async initialize() {
        try {
            this.#createDataGridStructure();
            this.#setupEventListeners();
            
            if (this.#config.columns?.length) {
                this.setColumns(this.#config.columns);
            }
            
            if (this.#config.data?.length) {
                this.setData(this.#config.data);
            }

            this.#eventSystem?.emit('datagrid:initialized', {
                component: this,
                timestamp: new Date().toISOString()
            });

        } catch (error) {
            console.error('DataGrid initialization failed:', error);
            throw error;
        }
    }

    /**
     * Sets column definitions
     * @param {Array} columns - Column definitions
     * @param {string} columns[].key - Column key
     * @param {string} columns[].title - Column title
     * @param {number} [columns[].width] - Column width
     * @param {string} [columns[].type='text'] - Column type (text, number, date, boolean)
     * @param {boolean} [columns[].sortable=true] - Is column sortable
     * @param {boolean} [columns[].filterable=true] - Is column filterable
     * @param {Function} [columns[].formatter] - Custom formatter function
     */
    setColumns(columns) {
        this.#state.columns = columns.map(col => ({
            key: col.key,
            title: col.title || col.key,
            width: col.width || 'auto',
            type: col.type || 'text',
            sortable: col.sortable !== false && this.#config.sortable,
            filterable: col.filterable !== false && this.#config.filterable,
            formatter: col.formatter || this.#getDefaultFormatter(col.type),
            ...col
        }));

        this.#renderHeader();
        this.#renderData();
        
        this.#eventSystem?.emit('datagrid:columns:updated', {
            columns: this.#state.columns,
            component: this
        });
    }

    /**
     * Sets data for the grid
     * @param {Array} data - Array of data objects
     */
    setData(data) {
        this.#state.data = Array.isArray(data) ? data : [];
        this.#state.selection.clear();
        this.#updateVirtualScrolling();
        this.#renderData();
        
        this.#eventSystem?.emit('datagrid:data:updated', {
            data: this.#state.data,
            component: this
        });
    }

    /**
     * Gets current data
     * @returns {Array} Current data array
     */
    getData() {
        return [...this.#state.data];
    }

    /**
     * Gets filtered and sorted data
     * @returns {Array} Processed data array
     */
    getProcessedData() {
        let processedData = [...this.#state.data];

        // Apply filters
        if (this.#state.filters.size > 0) {
            processedData = processedData.filter(row => {
                return Array.from(this.#state.filters.entries()).every(([columnKey, filterValue]) => {
                    const cellValue = row[columnKey];
                    return this.#applyFilter(cellValue, filterValue);
                });
            });
        }

        // Apply sorting
        if (this.#state.sortColumn) {
            const column = this.#state.columns.find(col => col.key === this.#state.sortColumn);
            if (column) {
                processedData.sort((a, b) => {
                    return this.#compareValues(
                        a[this.#state.sortColumn],
                        b[this.#state.sortColumn],
                        column.type,
                        this.#state.sortDirection
                    );
                });
            }
        }

        return processedData;
    }

    /**
     * Sorts data by column
     * @param {string} columnKey - Column key to sort by
     * @param {string} [direction] - Sort direction ('asc' or 'desc')
     */
    sortByColumn(columnKey, direction) {
        const column = this.#state.columns.find(col => col.key === columnKey);
        if (!column || !column.sortable) return;

        if (this.#state.sortColumn === columnKey) {
            this.#state.sortDirection = direction || (this.#state.sortDirection === 'asc' ? 'desc' : 'asc');
        } else {
            this.#state.sortColumn = columnKey;
            this.#state.sortDirection = direction || 'asc';
        }

        this.#renderData();
        this.#updateSortIndicators();
        
        this.#eventSystem?.emit('datagrid:sorted', {
            column: columnKey,
            direction: this.#state.sortDirection,
            component: this
        });
    }

    /**
     * Filters data by column
     * @param {string} columnKey - Column key to filter
     * @param {any} filterValue - Filter value
     */
    filterByColumn(columnKey, filterValue) {
        if (filterValue === null || filterValue === undefined || filterValue === '') {
            this.#state.filters.delete(columnKey);
        } else {
            this.#state.filters.set(columnKey, filterValue);
        }

        this.#renderData();
        
        this.#eventSystem?.emit('datagrid:filtered', {
            filters: Array.from(this.#state.filters.entries()),
            component: this
        });
    }

    /**
     * Clears all filters
     */
    clearFilters() {
        this.#state.filters.clear();
        this.#renderData();
        this.#clearFilterInputs();
        
        this.#eventSystem?.emit('datagrid:filters:cleared', {
            component: this
        });
    }

    /**
     * Selects rows
     * @param {Array|number} indices - Row indices to select
     */
    selectRows(indices) {
        if (!this.#config.selectable) return;

        if (!this.#config.multiSelect) {
            this.#state.selection.clear();
        }

        const indexArray = Array.isArray(indices) ? indices : [indices];
        indexArray.forEach(index => {
            if (index >= 0 && index < this.#state.data.length) {
                this.#state.selection.add(index);
            }
        });

        this.#updateRowSelection();
        
        this.#eventSystem?.emit('datagrid:selection:changed', {
            selection: Array.from(this.#state.selection),
            component: this
        });
    }

    /**
     * Clears selection
     */
    clearSelection() {
        this.#state.selection.clear();
        this.#updateRowSelection();
        
        this.#eventSystem?.emit('datagrid:selection:cleared', {
            component: this
        });
    }

    /**
     * Gets selected row indices
     * @returns {Array} Selected row indices
     */
    getSelection() {
        return Array.from(this.#state.selection);
    }

    /**
     * Gets selected row data
     * @returns {Array} Selected row data objects
     */
    getSelectedData() {
        const processedData = this.getProcessedData();
        return Array.from(this.#state.selection).map(index => processedData[index]).filter(Boolean);
    }

    /**
     * Destroys the component and cleans up resources
     * @returns {Promise<void>}
     */
    async destroy() {
        this.#removeEventListeners();
        
        if (this.#element) {
            this.#element.innerHTML = '';
            this.#element.classList.remove('onlyoffice-datagrid');
        }

        this.#element = null;
        this.#eventSystem = null;
        this.#stateManager = null;
    }

    /**
     * Sets up container element
     * @param {HTMLElement|string} container - Container element or selector
     * @private
     */
    #setupContainer(container) {
        if (typeof container === 'string') {
            this.#element = document.querySelector(container);
        } else if (container instanceof HTMLElement) {
            this.#element = container;
        } else {
            throw new Error('Invalid container provided to DataGrid');
        }

        if (!this.#element) {
            throw new Error('DataGrid container not found');
        }

        this.#element.classList.add('onlyoffice-datagrid');
    }

    /**
     * Creates the DataGrid DOM structure
     * @private
     */
    #createDataGridStructure() {
        this.#element.innerHTML = `
            <div class="datagrid-container">
                <div class="datagrid-header"></div>
                <div class="datagrid-body">
                    <div class="datagrid-viewport">
                        <div class="datagrid-content"></div>
                    </div>
                    <div class="datagrid-scrollbar">
                        <div class="datagrid-scrollbar-thumb"></div>
                    </div>
                </div>
            </div>
        `;

        // Get references to key elements
        this.#headerElement = this.#element.querySelector('.datagrid-header');
        this.#bodyElement = this.#element.querySelector('.datagrid-body');
        this.#viewportElement = this.#element.querySelector('.datagrid-viewport');
        this.#contentElement = this.#element.querySelector('.datagrid-content');
        this.#scrollbarElement = this.#element.querySelector('.datagrid-scrollbar');
        this.#scrollbarThumb = this.#element.querySelector('.datagrid-scrollbar-thumb');
    }

    /**
     * Sets up event listeners
     * @private
     */
    #setupEventListeners() {
        // Header sorting events
        this.#headerElement.addEventListener('click', this.#handleHeaderClick.bind(this));
        
        // Viewport scrolling events
        this.#viewportElement.addEventListener('scroll', this.#handleScroll.bind(this));
        
        // Row selection events
        this.#contentElement.addEventListener('click', this.#handleRowClick.bind(this));
        
        // Window resize events
        window.addEventListener('resize', this.#handleResize.bind(this));
        
        // Filter input events
        this.#headerElement.addEventListener('input', this.#handleFilterInput.bind(this));
    }

    /**
     * Removes event listeners
     * @private
     */
    #removeEventListeners() {
        if (this.#headerElement) {
            this.#headerElement.removeEventListener('click', this.#handleHeaderClick.bind(this));
            this.#headerElement.removeEventListener('input', this.#handleFilterInput.bind(this));
        }
        
        if (this.#viewportElement) {
            this.#viewportElement.removeEventListener('scroll', this.#handleScroll.bind(this));
        }
        
        if (this.#contentElement) {
            this.#contentElement.removeEventListener('click', this.#handleRowClick.bind(this));
        }
        
        window.removeEventListener('resize', this.#handleResize.bind(this));
    }

    /**
     * Renders the header
     * @private
     */
    #renderHeader() {
        if (!this.#headerElement || !this.#state.columns.length) return;

        const headerHtml = this.#state.columns.map(column => {
            const sortIcon = this.#getSortIcon(column.key);
            const filterInput = column.filterable ? 
                `<input type="text" class="datagrid-filter" data-column="${column.key}" placeholder="Filter...">` : '';
            
            return `
                <div class="datagrid-header-cell ${column.sortable ? 'sortable' : ''}" 
                     data-column="${column.key}" 
                     style="width: ${column.width === 'auto' ? 'auto' : column.width + 'px'}">
                    <div class="datagrid-header-content">
                        <span class="datagrid-header-title">${column.title}</span>
                        ${sortIcon}
                    </div>
                    ${filterInput}
                </div>
            `;
        }).join('');

        this.#headerElement.innerHTML = headerHtml;
    }

    /**
     * Renders the data
     * @private
     */
    #renderData() {
        if (!this.#contentElement || !this.#state.columns.length) return;

        const processedData = this.getProcessedData();
        
        if (this.#virtualScrolling) {
            this.#renderVirtualRows(processedData);
        } else {
            this.#renderAllRows(processedData);
        }
    }

    /**
     * Renders virtual rows for performance
     * @param {Array} data - Data to render
     * @private
     */
    #renderVirtualRows(data) {
        const containerHeight = this.#viewportElement.clientHeight;
        this.#visibleRows = Math.ceil(containerHeight / this.#rowHeight) + this.#config.bufferSize * 2;
        
        const startIndex = Math.max(0, Math.floor(this.#state.scrollTop / this.#rowHeight) - this.#config.bufferSize);
        const endIndex = Math.min(data.length, startIndex + this.#visibleRows);
        
        this.#state.visibleRange = { start: startIndex, end: endIndex };
        
        // Set content height for scrollbar
        const totalHeight = data.length * this.#rowHeight;
        this.#contentElement.style.height = `${totalHeight}px`;
        
        // Create visible rows
        const rowsHtml = [];
        for (let i = startIndex; i < endIndex; i++) {
            const row = data[i];
            const isSelected = this.#state.selection.has(i);
            rowsHtml.push(this.#createRowHtml(row, i, isSelected));
        }
        
        // Position content
        const offsetY = startIndex * this.#rowHeight;
        this.#contentElement.innerHTML = `
            <div class="datagrid-rows" style="transform: translateY(${offsetY}px)">
                ${rowsHtml.join('')}
            </div>
        `;
        
        this.#updateScrollbar(data.length);
    }

    /**
     * Renders all rows (non-virtual mode)
     * @param {Array} data - Data to render
     * @private
     */
    #renderAllRows(data) {
        const rowsHtml = data.map((row, index) => {
            const isSelected = this.#state.selection.has(index);
            return this.#createRowHtml(row, index, isSelected);
        }).join('');
        
        this.#contentElement.innerHTML = `<div class="datagrid-rows">${rowsHtml}</div>`;
    }

    /**
     * Creates HTML for a single row
     * @param {Object} row - Row data
     * @param {number} index - Row index
     * @param {boolean} isSelected - Is row selected
     * @returns {string} Row HTML
     * @private
     */
    #createRowHtml(row, index, isSelected) {
        const cellsHtml = this.#state.columns.map(column => {
            const value = row[column.key];
            const formattedValue = column.formatter(value, row, column);
            
            return `
                <div class="datagrid-cell" data-column="${column.key}">
                    ${formattedValue}
                </div>
            `;
        }).join('');

        return `
            <div class="datagrid-row ${isSelected ? 'selected' : ''}" 
                 data-index="${index}"
                 style="height: ${this.#rowHeight}px">
                ${cellsHtml}
            </div>
        `;
    }

    /**
     * Updates virtual scrolling calculations
     * @private
     */
    #updateVirtualScrolling() {
        if (!this.#virtualScrolling) return;
        
        const containerHeight = this.#viewportElement?.clientHeight || 400;
        this.#visibleRows = Math.ceil(containerHeight / this.#rowHeight) + this.#config.bufferSize * 2;
    }

    /**
     * Updates scrollbar
     * @param {number} totalRows - Total number of rows
     * @private
     */
    #updateScrollbar(totalRows) {
        if (!this.#scrollbarElement || !this.#virtualScrolling) return;
        
        const containerHeight = this.#viewportElement.clientHeight;
        const contentHeight = totalRows * this.#rowHeight;
        
        if (contentHeight <= containerHeight) {
            this.#scrollbarElement.style.display = 'none';
            return;
        }
        
        this.#scrollbarElement.style.display = 'block';
        
        const thumbHeight = Math.max(20, (containerHeight / contentHeight) * containerHeight);
        const thumbTop = (this.#state.scrollTop / contentHeight) * containerHeight;
        
        this.#scrollbarThumb.style.height = `${thumbHeight}px`;
        this.#scrollbarThumb.style.top = `${thumbTop}px`;
    }

    /**
     * Gets sort icon for column
     * @param {string} columnKey - Column key
     * @returns {string} Sort icon HTML
     * @private
     */
    #getSortIcon(columnKey) {
        if (this.#state.sortColumn !== columnKey) {
            return '<span class="datagrid-sort-icon"></span>';
        }
        
        const direction = this.#state.sortDirection;
        const icon = direction === 'asc' ? '▲' : '▼';
        return `<span class="datagrid-sort-icon active">${icon}</span>`;
    }

    /**
     * Updates sort indicators in header
     * @private
     */
    #updateSortIndicators() {
        const headerCells = this.#headerElement.querySelectorAll('.datagrid-header-cell');
        headerCells.forEach(cell => {
            const columnKey = cell.dataset.column;
            const sortIcon = cell.querySelector('.datagrid-sort-icon');
            
            if (sortIcon) {
                if (this.#state.sortColumn === columnKey) {
                    const icon = this.#state.sortDirection === 'asc' ? '▲' : '▼';
                    sortIcon.textContent = icon;
                    sortIcon.classList.add('active');
                } else {
                    sortIcon.textContent = '';
                    sortIcon.classList.remove('active');
                }
            }
        });
    }

    /**
     * Updates row selection visual state
     * @private
     */
    #updateRowSelection() {
        const rows = this.#contentElement.querySelectorAll('.datagrid-row');
        rows.forEach(row => {
            const index = parseInt(row.dataset.index);
            row.classList.toggle('selected', this.#state.selection.has(index));
        });
    }

    /**
     * Clears filter inputs
     * @private
     */
    #clearFilterInputs() {
        const filterInputs = this.#headerElement.querySelectorAll('.datagrid-filter');
        filterInputs.forEach(input => {
            input.value = '';
        });
    }

    /**
     * Gets default formatter for column type
     * @param {string} type - Column type
     * @returns {Function} Formatter function
     * @private
     */
    #getDefaultFormatter(type) {
        switch (type) {
            case 'number':
                return (value) => typeof value === 'number' ? value.toLocaleString() : value;
            case 'date':
                return (value) => value instanceof Date ? value.toLocaleDateString() : value;
            case 'boolean':
                return (value) => value ? '✓' : '✗';
            default:
                return (value) => value != null ? String(value) : '';
        }
    }

    /**
     * Applies filter to value
     * @param {any} cellValue - Cell value
     * @param {any} filterValue - Filter value
     * @returns {boolean} Whether value passes filter
     * @private
     */
    #applyFilter(cellValue, filterValue) {
        if (!filterValue) return true;
        
        const cellStr = String(cellValue || '').toLowerCase();
        const filterStr = String(filterValue).toLowerCase();
        
        return cellStr.includes(filterStr);
    }

    /**
     * Compares values for sorting
     * @param {any} a - First value
     * @param {any} b - Second value
     * @param {string} type - Column type
     * @param {string} direction - Sort direction
     * @returns {number} Comparison result
     * @private
     */
    #compareValues(a, b, type, direction) {
        let result = 0;
        
        switch (type) {
            case 'number':
                result = (Number(a) || 0) - (Number(b) || 0);
                break;
            case 'date':
                const dateA = a instanceof Date ? a : new Date(a);
                const dateB = b instanceof Date ? b : new Date(b);
                result = dateA.getTime() - dateB.getTime();
                break;
            default:
                result = String(a || '').localeCompare(String(b || ''));
        }
        
        return direction === 'desc' ? -result : result;
    }

    /**
     * Binds component events
     * @private
     */
    #bindEvents() {
        // No additional event binding needed here
    }

    /**
     * Handles header click events
     * @param {Event} event - Click event
     * @private
     */
    #handleHeaderClick(event) {
        const headerCell = event.target.closest('.datagrid-header-cell');
        if (!headerCell || !headerCell.classList.contains('sortable')) return;
        
        const columnKey = headerCell.dataset.column;
        this.sortByColumn(columnKey);
    }

    /**
     * Handles viewport scroll events
     * @param {Event} event - Scroll event
     * @private
     */
    #handleScroll(event) {
        this.#state.scrollTop = event.target.scrollTop;
        
        if (this.#virtualScrolling) {
            this.#renderData();
        }
    }

    /**
     * Handles row click events
     * @param {Event} event - Click event
     * @private
     */
    #handleRowClick(event) {
        if (!this.#config.selectable) return;
        
        const row = event.target.closest('.datagrid-row');
        if (!row) return;
        
        const index = parseInt(row.dataset.index);
        
        if (this.#config.multiSelect && event.ctrlKey) {
            if (this.#state.selection.has(index)) {
                this.#state.selection.delete(index);
            } else {
                this.#state.selection.add(index);
            }
        } else {
            this.selectRows([index]);
        }
    }

    /**
     * Handles window resize events
     * @private
     */
    #handleResize() {
        this.#updateVirtualScrolling();
        this.#renderData();
    }

    /**
     * Handles filter input events
     * @param {Event} event - Input event
     * @private
     */
    #handleFilterInput(event) {
        if (!event.target.classList.contains('datagrid-filter')) return;
        
        const columnKey = event.target.dataset.column;
        const filterValue = event.target.value.trim();
        
        this.filterByColumn(columnKey, filterValue);
    }
}

// Export DataGrid component
if (typeof window !== 'undefined') {
    window.DataGrid = DataGrid;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = DataGrid;
}

export default DataGrid;