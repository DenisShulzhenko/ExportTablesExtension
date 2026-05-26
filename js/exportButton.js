/**
 * Export Button Module
 * Handles injecting export buttons above detected tables
 */

var ExportButton = (function() {
    'use strict';

    // Style for the export button
    var BUTTON_STYLE_ID = 'istok-export-button-style';
    var BUTTON_CONTAINER_CLASS = 'istok-export-button-container';
    var DROPDOWN_CLASS = 'istok-export-dropdown';

    /**
     * Injects the CSS styles for the export button
     */
    function injectStyles() {
        if (document.getElementById(BUTTON_STYLE_ID)) {
            return;
        }

        var style = document.createElement('style');
        style.id = BUTTON_STYLE_ID;
        style.textContent = '.' + BUTTON_CONTAINER_CLASS + ' {\n' +
            '    display: flex;\n' +
            '    justify-content: flex-end;\n' +
            '    padding: 10px 15px;\n' +
            '    margin-bottom: 5px;\n' +
            '    background-color: #f8f9fa;\n' +
            '    border-bottom: 1px solid #dee2e6;\n' +
            '}\n' +
            '.istok-export-btn {\n' +
            '    display: inline-flex;\n' +
            '    align-items: center;\n' +
            '    gap: 6px;\n' +
            '    padding: 8px 16px;\n' +
            '    font-size: 14px;\n' +
            '    font-weight: 500;\n' +
            '    color: #fff;\n' +
            '    background-color: #28a745;\n' +
            '    border: none;\n' +
            '    border-radius: 4px;\n' +
            '    cursor: pointer;\n' +
            '    transition: background-color 0.2s ease, box-shadow 0.2s ease;\n' +
            '    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;\n' +
            '}\n' +
            '.istok-export-btn:hover {\n' +
            '    background-color: #218838;\n' +
            '    box-shadow: 0 2px 4px rgba(0,0,0,0.15);\n' +
            '}\n' +
            '.istok-export-btn:active {\n' +
            '    background-color: #1e7e34;\n' +
            '    transform: translateY(1px);\n' +
            '}\n' +
            '.istok-export-btn svg {\n' +
            '    width: 16px;\n' +
            '    height: 16px;\n' +
            '    fill: currentColor;\n' +
            '}\n' +
            '.istok-export-btn.exporting {\n' +
            '    background-color: #6c757d;\n' +
            '    pointer-events: none;\n' +
            '}\n' +
            /* Dropdown styles */
            '.istok-export-dropdown {\n' +
            '    position: relative;\n' +
            '    display: inline-block;\n' +
            '}\n' +
            '.istok-export-dropdown-content {\n' +
            '    display: none;\n' +
            '    position: absolute;\n' +
            '    right: 0;\n' +
            '    min-width: 200px;\n' +
            '    background-color: #fff;\n' +
            '    box-shadow: 0 4px 8px rgba(0,0,0,0.15);\n' +
            '    border-radius: 4px;\n' +
            '    z-index: 1000;\n' +
            '    overflow: hidden;\n' +
            '    margin-top: 4px;\n' +
            '}\n' +
            '.istok-export-dropdown-content.show {\n' +
            '    display: block;\n' +
            '}\n' +
            '.istok-export-dropdown-item {\n' +
            '    display: flex;\n' +
            '    align-items: center;\n' +
            '    gap: 10px;\n' +
            '    padding: 10px 16px;\n' +
            '    color: #333;\n' +
            '    text-decoration: none;\n' +
            '    cursor: pointer;\n' +
            '    transition: background-color 0.15s ease;\n' +
            '    font-size: 14px;\n' +
            '    border: none;\n' +
            '    background: none;\n' +
            '    width: 100%;\n' +
            '    text-align: left;\n' +
            '    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;\n' +
            '}\n' +
            '.istok-export-dropdown-item:hover {\n' +
            '    background-color: #f8f9fa;\n' +
            '}\n' +
            '.istok-export-dropdown-item:active {\n' +
            '    background-color: #e9ecef;\n' +
            '}\n' +
            '.istok-export-dropdown-item .format-icon {\n' +
            '    width: 20px;\n' +
            '    height: 20px;\n' +
            '    display: flex;\n' +
            '    align-items: center;\n' +
            '    justify-content: center;\n' +
            '    font-weight: bold;\n' +
            '    font-size: 10px;\n' +
            '    border-radius: 3px;\n' +
            '    color: #fff;\n' +
            '}\n' +
            '.istok-export-dropdown-item .format-icon.excel { background-color: #217346; }\n' +
            '.istok-export-dropdown-item .format-icon.doc { background-color: #2b579a; }\n' +
            '.istok-export-dropdown-divider {\n' +
            '    height: 1px;\n' +
            '    background-color: #dee2e6;\n' +
            '    margin: 4px 0;\n' +
            '}';
        document.head.appendChild(style);
    }

    /**
     * Creates the SVG icon for the export button
     * @returns {string} - SVG icon string
     */
    function createExportIcon() {
        return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">' +
            '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>' +
            '<path d="M14 2v6h6"/>' +
            '<path d="M12 18v-6"/>' +
            '<path d="M9 15l3-3 3 3"/>' +
            '</svg>';
    }

    /**
     * Creates a dropdown export button for a table
     * @param {HTMLTableElement} table - The table to export
     * @param {Object} exporters - Object containing export functions
     * @returns {HTMLElement} - The created button container
     */
    function createExportButton(table, exporters) {
        // Create container
        var container = document.createElement('div');
        container.className = BUTTON_CONTAINER_CLASS;

        // Create dropdown wrapper
        var dropdown = document.createElement('div');
        dropdown.className = DROPDOWN_CLASS;

        // Create main button
        var button = document.createElement('button');
        button.className = 'istok-export-btn';
        button.innerHTML = createExportIcon() + 'Экспорт ▼';
        button.title = 'Экспортировать таблицу';

        // Create dropdown content
        var dropdownContent = document.createElement('div');
        dropdownContent.className = 'istok-export-dropdown-content';

        // Add export options
        var formats = [
            { name: 'Excel (.xls)', icon: 'XLS', format: 'excel', exporter: exporters.excel },
            { name: 'Word (.doc)', icon: 'DOC', format: 'doc', exporter: exporters.doc }
        ];

        for (var i = 0; i < formats.length; i++) {
            var fmt = formats[i];
            
            // Add divider before ODS
            if (i === 2) {
                var divider = document.createElement('div');
                divider.className = 'istok-export-dropdown-divider';
                dropdownContent.appendChild(divider);
            }

            var item = document.createElement('button');
            item.className = 'istok-export-dropdown-item';
            item.innerHTML = '<span class="format-icon ' + fmt.format + '">' + fmt.icon + '</span>' + fmt.name;
            
            item.addEventListener('click', function(format, exp) {
                return function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    // Hide dropdown
                    dropdownContent.classList.remove('show');
                    
                    // Add visual feedback
                    button.classList.add('exporting');
                    button.innerHTML = 'Экспорт...';
                    
                    try {
                        exp(table);
                    } catch (error) {
                        console.error('[ExportButton] Export error:', error);
                        alert('Ошибка при экспорте: ' + error.message);
                    } finally {
                        // Reset button state after a short delay
                        setTimeout(function() {
                            button.classList.remove('exporting');
                            button.innerHTML = createExportIcon() + 'Экспорт ▼';
                        }, 1000);
                    }
                };
            }(fmt, fmt.exporter));

            dropdownContent.appendChild(item);
        }

        // Toggle dropdown on button click
        button.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            dropdownContent.classList.toggle('show');
        });

        // Close dropdown when clicking outside
        document.addEventListener('click', function(e) {
            if (!dropdown.contains(e.target)) {
                dropdownContent.classList.remove('show');
            }
        });

        dropdown.appendChild(button);
        dropdown.appendChild(dropdownContent);
        container.appendChild(dropdown);
        return container;
    }

    /**
     * Injects an export button above a table
     * @param {HTMLTableElement} table - The table to add button to
     * @param {Object} exporters - Object containing export functions
     * @returns {boolean} - True if button was added successfully
     */
    function injectButtonAboveTable(table, exporters) {
        // Check if button already exists
        var existingContainer = table.previousElementSibling;
        if (existingContainer && existingContainer.className === BUTTON_CONTAINER_CLASS) {
            return false;
        }

        // Inject styles
        injectStyles();

        // Create and insert button
        var buttonContainer = createExportButton(table, exporters);
        
        // Insert before the table
        if (table.parentNode) {
            table.parentNode.insertBefore(buttonContainer, table);
            return true;
        }

        return false;
    }

    /**
     * Removes all export buttons from the page
     */
    function removeAllButtons() {
        var containers = document.querySelectorAll('.' + BUTTON_CONTAINER_CLASS);
        for (var i = 0; i < containers.length; i++) {
            containers[i].remove();
        }
    }

    /**
     * Gets all tables that don't have export buttons yet
     * @param {HTMLTableElement[]} tables - Array of tables to check
     * @returns {HTMLTableElement[]} - Tables without buttons
     */
    function getTablesWithoutButtons(tables) {
        var result = [];
        for (var i = 0; i < tables.length; i++) {
            var table = tables[i];
            var previous = table.previousElementSibling;
            if (!previous || previous.className !== BUTTON_CONTAINER_CLASS) {
                result.push(table);
            }
        }
        return result;
    }

    return {
        injectButtonAboveTable: injectButtonAboveTable,
        removeAllButtons: removeAllButtons,
        getTablesWithoutButtons: getTablesWithoutButtons,
        injectStyles: injectStyles
    };
})();