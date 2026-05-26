/**
 * Istok Extension - Table Export Content Script
 * Main entry point that coordinates table detection, button injection, and Excel export
 */
(function () {
    'use strict';

    console.log('[Istok Table Export] Content script loaded');

    /**
     * Main initialization function
     */
    function init() {
        // Wait for DOM to be ready
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', onDOMReady);
        } else {
            onDOMReady();
        }
    }

    /**
     * Called when DOM is ready
     */
    function onDOMReady() {
        console.log('[Istok Table Export] DOM ready, scanning for tables...');

        // Initial scan
        scanAndAddButtons();

        // Set up mutation observer to handle dynamically loaded tables
        setupMutationObserver();
    }

    /**
     * Scans the page for tables and adds export buttons
     */
    function scanAndAddButtons() {
        // Find all exportable tables
        var tables = TableDetector.findAllTables();

        console.log('[Istok Table Export] Found', tables.length, 'exportable table(s)');

        if (tables.length === 0) {
            console.log('[Istok Table Export] No tables found to export');
            return;
        }

        // Create exporters object with all export functions
        var exporters = {
            excel: function (table) {
                return handleExport(table, 'excel');
            },
            csv: function (table) {
                return handleExport(table, 'csv');
            },
            doc: function (table) {
                return handleExport(table, 'doc');
            },
            ods: function (table) {
                return handleExport(table, 'ods');
            }
        };

        // Add export buttons to tables that don't have them
        var buttonsAdded = 0;
        for (var i = 0; i < tables.length; i++) {
            var table = tables[i];
            var added = ExportButton.injectButtonAboveTable(table, exporters);
            if (added) {
                buttonsAdded++;
                console.log('[Istok Table Export] Added export button to table:', TableDetector.getTableId(table));
            }
        }

        console.log('[Istok Table Export] Added', buttonsAdded, 'export button(s)');
    }

    /**
     * Handles the export button click
     * @param {HTMLTableElement} table - The table to export
     * @param {string} format - Export format ('excel', 'csv', 'doc', 'ods')
     */
    function handleExport(table, format) {
        console.log('[Istok Table Export] Export requested for table:', TableDetector.getTableId(table), 'format:', format);

        // Get table data
        var tableData = TableDetector.getTableData(table);

        // Validate data
        if (!tableData.headers || tableData.headers.length === 0) {
            throw new Error('Таблица не содержит заголовков');
        }

        if (!tableData.rows || tableData.rows.length === 0) {
            throw new Error('Таблица не содержит данных');
        }

        console.log('[Istok Table Export] Table data:', {
            headers: tableData.headers.length,
            rows: tableData.rows.length
        });

        // Generate filename with page title prefix
        var pageTitle = document.title || 'export';
        // Sanitize page title for filename (remove invalid characters)
        var sanitizedTitle = pageTitle.replace(/[<>:"/\\|?*]/g, '').trim().substring(0, 50);
        var tableId = TableDetector.getTableId(table);
        var filename = sanitizedTitle + '_' + tableId + '_' + new Date().toISOString().slice(0, 10);

        // Export based on format
        switch (format) {
            case 'csv':
                CsvExporter.exportToCsv(tableData, filename);
                break;
            case 'doc':
                DocExporter.exportToDoc(tableData, filename);
                break;
            case 'ods':
                OdsExporter.exportToOds(tableData, filename);
                break;
            case 'excel':
            default:
                ExcelExporter.exportToExcel(tableData, filename);
                break;
        }
    }

    /**
     * Sets up a mutation observer to handle dynamically loaded tables
     */
    function setupMutationObserver() {
        // Check if MutationObserver is supported
        if (!window.MutationObserver) {
            console.warn('[Istok Table Export] MutationObserver not supported');
            return;
        }

        var observer = new MutationObserver(function (mutations) {
            var shouldRescan = false;

            for (var m = 0; m < mutations.length; m++) {
                var mutation = mutations[m];
                // Check if new nodes were added
                if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
                    for (var n = 0; n < mutation.addedNodes.length; n++) {
                        var node = mutation.addedNodes[n];
                        // Check if a table was added
                        if (node.nodeName === 'TABLE' ||
                            (node.nodeType === 1 && node.querySelector && node.querySelector('table'))) {
                            shouldRescan = true;
                            break;
                        }
                    }
                }

                if (shouldRescan) break;
            }

            if (shouldRescan) {
                // Debounce the rescan
                clearTimeout(window._istokTableRescanTimeout);
                window._istokTableRescanTimeout = setTimeout(function () {
                    scanAndAddButtons();
                }, 500);
            }
        });

        // Start observing the document
        observer.observe(document.body || document.documentElement, {
            childList: true,
            subtree: true
        });

        console.log('[Istok Table Export] Mutation observer set up');
    }

    // Initialize when script is loaded
    init();
})();