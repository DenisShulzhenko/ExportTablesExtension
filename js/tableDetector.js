/**
 * Table Detector Module
 * Detects tables in the page that match the expected pattern (like the class journal table)
 */

var TableDetector = (function () {
    'use strict';

    function isExportableTable(table) {
        // Skip tables that are too small (likely not data tables)
        if (table.rows.length < 2) {
            return false;
        }

        // Check if table has a body (tbody)
        var tbody = table.querySelector('tbody');
        if (!tbody || tbody.rows.length === 0) {
            return false;
        }

        // Skip tables with specific IDs that are likely UI elements (not data)
        var tableId = table.id.toLowerCase();
        var skipIds = ['datatable-buttons'];
        var shouldSkip = false;
        for (var i = 0; i < skipIds.length; i++) {
            if (tableId.indexOf(skipIds[i]) !== -1) {
                shouldSkip = true;
                break;
            }
        }
        if (shouldSkip) {
            return false;
        }

        // Filter out tables without proper headers (thead or header cells)
        var thead = table.querySelector('thead');
        if (!thead) {
            // Check if first row contains mostly header cells (th)
            var firstRow = table.rows[0];
            if (firstRow) {
                var thCount = firstRow.querySelectorAll('th').length;
                var tdCount = firstRow.querySelectorAll('td').length;
                var hasHeaderCells = thCount > 0;
                var isLikelyHeaderRow = thCount > (tdCount / 2); // More than half are th elements
                
                if (!hasHeaderCells || !isLikelyHeaderRow) {
                    return false;
                }
            }
        }

        // Filter out button tables (tables with mostly buttons, links, or interactive elements)
        var totalCells = 0;
        var interactiveCells = 0;
        
        for (var r = 0; r < table.rows.length && r < 5; r++) { // Check first 5 rows
            var row = table.rows[r];
            for (var c = 0; c < row.cells.length; c++) {
                var cell = row.cells[c];
                totalCells++;
                
                // Check if cell contains primarily interactive elements
                var buttons = cell.querySelectorAll('button');
                var links = cell.querySelectorAll('a');
                var inputs = cell.querySelectorAll('input');
                
                var cellText = cell.textContent.trim();
                var hasInteractiveElements = buttons.length > 0 || links.length > 0 || inputs.length > 0;
                var hasMinimalText = cellText.length < 3;
                
                if (hasInteractiveElements && hasMinimalText) {
                    interactiveCells++;
                }
            }
        }
        
        // If more than 50% of checked cells are interactive with minimal text, skip this table
        if (totalCells > 0 && (interactiveCells / totalCells) > 0.5) {
            return false;
        }

        // Filter out very small tables (less than 2 columns and 2 data rows)
        var columnCount = table.rows[0] ? table.rows[0].cells.length : 0;
        var dataRowCount = tbody ? tbody.rows.length : 0;
        
        if (columnCount < 2 || dataRowCount < 2) {
            return false;
        }

        return true;
    }

    /**
     * Finds all exportable tables in the document
     * @returns {HTMLTableElement[]} - Array of exportable table elements
     */
    function findAllTables() {
        var tables = document.querySelectorAll('table');
        var result = [];
        for (var i = 0; i < tables.length; i++) {
            if (isExportableTable(tables[i])) {
                result.push(tables[i]);
            }
        }
        return result;
    }

    /**
     * Extracts tooltip text from an anchor element
     * @param {HTMLAnchorElement} anchor - The anchor element
     * @returns {string|null} - The tooltip text or null if not found
     */
    function getAnchorTooltip(anchor) {
        // Try data-original-title first (Bootstrap tooltip)
        var tooltip = anchor.getAttribute('data-original-title');
        if (tooltip) {
            return tooltip.trim();
        }
        // Try title attribute as fallback
        tooltip = anchor.getAttribute('title');
        if (tooltip) {
            return tooltip.trim();
        }
        return null;
    }

    /**
     * Collects tooltip text from a header cell
     * @param {HTMLTableCellElement} cell - The header cell
     * @returns {string|null} - The tooltip text or null if not found
     */
    function getHeaderCellTooltip(cell) {
        // Try data-original-title first (Bootstrap tooltip)
        var tooltip = cell.getAttribute('data-original-title');
        if (tooltip) {
            return tooltip.trim();
        }
        // Try title attribute as fallback
        tooltip = cell.getAttribute('title');
        if (tooltip) {
            return tooltip.trim();
        }
        // Try getting tooltip from anchor inside the cell
        var anchor = cell.querySelector('a');
        if (anchor) {
            return getAnchorTooltip(anchor);
        }
        return null;
    }

    function getTableData(table) {
        var data = {
            headers: [],
            headerRows: [],
            rows: [],
            merges: [],
            headerLevels: 1,
            cellMeta: []
        };

        if (!table || !table.rows || table.rows.length === 0) {
            return data;
        }

        // Use the universal TableParser to build the grid with metadata
        var parsedWithMeta = TableParser.parseTableWithMeta(table);
        var grid = parsedWithMeta.grid;
        var cellMeta = parsedWithMeta.cellMeta;

        // Also get the parsed result for headers
        var parsed = TableParser.parseTable(table, {raw: true, withHeader: true});

        // Get thead for header processing
        var thead = table.querySelector('thead');
        var headerRowCount = thead ? thead.rows.length : 0;

        // Set header rows from parsed result
        if (parsed.headers && parsed.headers.length > 0) {
            data.headerRows = parsed.headers;
            data.headers = parsed.headers[0];
            data.headerLevels = parsed.headers.length;
        }

        // Set data rows from grid (skip header rows)
        var dataStartIndex = headerRowCount > 0 ? headerRowCount : 1;
        for (var r = dataStartIndex; r < grid.length; r++) {
            data.rows.push(grid[r].slice());
        }

        // Build complete merges from cellMeta for both colspan and rowspan
        // Only process cells that are marked as "original" (first cell of a span)
        for (var r = 0; r < cellMeta.length; r++) {
            var metaRow = cellMeta[r];
            for (var c = 0; c < metaRow.length; c++) {
                var meta = metaRow[c];
                if (meta && meta.isOriginal) {
                    // If colspan > 1 or rowspan > 1, we need to create a merge
                    if (meta.colspan > 1 || meta.rowspan > 1) {
                        data.merges.push({
                            row: r,
                            col: c,
                            cells: meta.colspan,
                            rows: meta.rowspan
                        });
                    }
                }
            }
        }

        data.cellMeta = cellMeta;

        return data;
    }

    /**
     * Gets a unique identifier for a table based on its position and attributes
     * @param {HTMLTableElement} table - The table element
     * @returns {string} - Unique identifier
     */
    function getTableId(table) {
        if (table.id) {
            return table.id;
        }

        // Generate ID based on table's position in the document
        var allTables = findAllTables();
        var index = -1;
        for (var i = 0; i < allTables.length; i++) {
            if (allTables[i] === table) {
                index = i;
                break;
            }
        }

        return 'table-' + index;
    }

    return {
        findAllTables: findAllTables,
        getTableData: getTableData,
        getTableId: getTableId,
        isExportableTable: isExportableTable
    };
})();