/**
 * Universal Table Parser
 * Simple and robust HTML table parsing with rowspan/colspan support
 */

var TableParser = (function() {
    'use strict';

    /**
     * Parses an HTML table element into a structured object
     * @param {HTMLTableElement} table - The table element to parse
     * @param {Object} options - Optional configuration
     * @param {boolean} options.withHeader - Include first row as headers (default: true)
     * @param {boolean} options.raw - Include raw 2D array output (default: true)
     * @returns {Object} Parsed table data
     */
    function parseTable(table, options) {
        options = options || {};
        var withHeader = options.withHeader !== false;
        var raw = options.raw !== false;

        var result = {
            rows: []
        };

        if (raw) {
            result.raw = [];
        }

        if (!table || !table.rows || table.rows.length === 0) {
            return result;
        }

        // Build a 2D grid representation of the table
        var grid = buildGrid(table);

        if (grid.length === 0) {
            return result;
        }

        // Determine header rows and data rows
        var thead = table.querySelector('thead');
        var headerRowCount = thead ? thead.rows.length : 0;

        // Extract headers
        var headers = [];
        if (headerRowCount > 0) {
            // Multi-row headers from thead
            for (var hr = 0; hr < headerRowCount; hr++) {
                headers.push(grid[hr].slice());
            }
        } else if (withHeader) {
            // Use first data row as header
            headers = [grid[0].slice()];
            grid = grid.slice(1);
        }

        // Extract data rows
        for (var r = 0; r < grid.length; r++) {
            var rowData = grid[r];

            if (raw) {
                result.raw.push(rowData);
            }

            // Convert to object if we have headers
            if (headers.length > 0) {
                var rowObj = {};
                var headerRow = headers[0];
                for (var c = 0; c < headerRow.length; c++) {
                    var key = headerRow[c] || 'column_' + c;
                    rowObj[key] = rowData[c] || '';
                }
                result.rows.push(rowObj);
            } else {
                result.rows.push(rowData);
            }
        }

        // Add headers to result
        if (headers.length > 0) {
            result.headers = headers;
        }

        return result;
    }

    /**
     * Builds a 2D grid representation of the table, handling rowspan and colspan
     * @param {HTMLTableElement} table - The table element
     * @returns {Array[]} 2D array of cell values
     */
    function buildGrid(table) {
        var grid = [];
        var rowspans = []; // Track remaining rowspan for each column

        for (var r = 0; r < table.rows.length; r++) {
            var row = table.rows[r];
            var gridRow = [];
            var colIndex = 0;

            // Fill in cells from previous rowspans
            for (var c = 0; c < rowspans.length; c++) {
                if (rowspans[c] >= 1) {
                    // This column is occupied by a rowspan from a previous row
                    gridRow[c] = grid[grid.length - 1][c];
                    rowspans[c]--;
                }
            }

            // Adjust colIndex to first empty column
            while (colIndex < rowspans.length && rowspans[colIndex] >= 1) {
                colIndex++;
            }

            // Process cells in current row
            for (var c = 0; c < row.cells.length; c++) {
                var cell = row.cells[c];
                var colspan = parseInt(cell.getAttribute('colspan')) || 1;
                var rowspan = parseInt(cell.getAttribute('rowspan')) || 1;

                // Skip to next available column
                while (colIndex < rowspans.length && rowspans[colIndex] >= 1) {
                    colIndex++;
                }

                var cellText = getCellText(cell);

                // Add cell value for colspan
                for (var cs = 0; cs < colspan; cs++) {
                    gridRow[colIndex + cs] = cellText;
                }

                // Set rowspan for future rows
                if (rowspan > 1) {
                    for (var cs = 0; cs < colspan; cs++) {
                        rowspans[colIndex + cs] = rowspan;
                    }
                }

                colIndex += colspan;
            }

            grid.push(gridRow);
        }

        return grid;
    }

    /**
     * Builds a 2D grid with cell metadata (colspan/rowspan) for proper Excel export
     * @param {HTMLTableElement} table - The table element
     * @returns {Object} Object with grid (2D array of cell values) and cellMeta (2D array of metadata)
     */
    function buildGridWithMeta(table) {
        var grid = [];
        var cellMeta = []; // 2D array storing {colspan, rowspan, isOriginal} for each cell
        var rowspans = []; // Track remaining rowspan for each column

        for (var r = 0; r < table.rows.length; r++) {
            var row = table.rows[r];
            var gridRow = [];
            var metaRow = [];
            var colIndex = 0;

            // Fill in cells from previous rowspans
            for (var c = 0; c < rowspans.length; c++) {
                if (rowspans[c] >= 1) {
                    // This column is occupied by a rowspan from a previous row
                    gridRow[c] = grid[grid.length - 1][c];
                    // The cell is covered by a previous rowspan, so mark it as non-original in this row
                    metaRow[c] = {
                        colspan: 1,
                        rowspan: 1,
                        isOriginal: false
                    };
                    rowspans[c]--;
                }
            }

            // Adjust colIndex to first empty column
            while (colIndex < rowspans.length && rowspans[colIndex] >= 1) {
                colIndex++;
            }

            // Process cells in current row
            for (var c = 0; c < row.cells.length; c++) {
                var cell = row.cells[c];
                var colspan = parseInt(cell.getAttribute('colspan')) || 1;
                var rowspan = parseInt(cell.getAttribute('rowspan')) || 1;

                // Skip to next available column
                while (colIndex < rowspans.length && rowspans[colIndex] >= 1) {
                    colIndex++;
                }

                var cellText = getCellText(cell);

                // Add cell value for colspan
                for (var cs = 0; cs < colspan; cs++) {
                    gridRow[colIndex + cs] = cellText;
                    // Only the first cell gets the original metadata
                    if (cs === 0) {
                        metaRow[colIndex + cs] = {
                            colspan: colspan,
                            rowspan: rowspan,
                            isOriginal: true
                        };
                    } else {
                        metaRow[colIndex + cs] = {
                            colspan: 1,
                            rowspan: 1,
                            isOriginal: false
                        };
                    }
                }

                // Set rowspan for future rows
                if (rowspan > 1) {
                    for (var cs = 0; cs < colspan; cs++) {
                        rowspans[colIndex + cs] = rowspan;
                    }
                }

                colIndex += colspan;
            }

            grid.push(gridRow);
            cellMeta.push(metaRow);
        }

        return {
            grid: grid,
            cellMeta: cellMeta
        };
    }

    /**
     * Extracts clean text from a table cell
     * @param {HTMLTableCellElement} cell - The table cell
     * @returns {string} Clean cell text
     */
    function getCellText(cell) {
        if (!cell) return '';

        // Check for select elements inside the cell
        var selectElement = cell.querySelector('select');
        if (selectElement) {
            var selectedOption = selectElement.options[selectElement.selectedIndex];
            var text = selectedOption ? selectedOption.text : '';
            text = text.replace(/\s+/g, ' ').trim();
            return text;
        }

        var text = '';
        var clone = cell.cloneNode(true);
        var brElements = clone.querySelectorAll('br');
        for (var i = 0; i < brElements.length; i++) {
            var br = brElements[i];
            var spaceNode = document.createTextNode(' ');
            br.parentNode.replaceChild(spaceNode, br);
        }
        text = clone.textContent || '';
        text = text.replace(/\s+/g, ' ').trim();

        if (!text) {
            text = cell.getAttribute('value') || cell.value || cell.getAttribute('title') || cell.getAttribute('data-original-title') || '';
            text = String(text).replace(/\s+/g, ' ').trim();
        }

        return text;
    }

    /**
     * Parses a table and returns only the raw 2D array
     * @param {HTMLTableElement} table - The table element
     * @returns {Array[]} 2D array of cell values
     */
    function parseTableRaw(table) {
        var result = parseTable(table, { raw: true, withHeader: true });
        return result.raw || [];
    }

    /**
     * Parses a table and returns array of row objects
     * @param {HTMLTableElement} table - The table element
     * @returns {Object[]} Array of row objects
     */
    function parseTableAsObjects(table) {
        var result = parseTable(table, { raw: false, withHeader: true });
        return result.rows || [];
    }

    /**
     * Parses a table and returns grid with cell metadata for proper Excel export
     * @param {HTMLTableElement} table - The table element
     * @returns {Object} Object with grid (2D array) and cellMeta (2D array of metadata)
     */
    function parseTableWithMeta(table) {
        return buildGridWithMeta(table);
    }

    return {
        parseTable: parseTable,
        parseTableRaw: parseTableRaw,
        parseTableAsObjects: parseTableAsObjects,
        parseTableWithMeta: parseTableWithMeta
    };
})();

// Export for different environments
if (typeof module !== 'undefined' && module.exports) {
    module.exports = TableParser;
}