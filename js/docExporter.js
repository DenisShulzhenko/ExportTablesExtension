/**
 * DOC Exporter Module
 * Handles exporting table data to DOC format (.doc)
 * Uses HTML-based DOC format that Word can open
 */

var DocExporter = (function() {
    'use strict';

    /**
     * Escapes special HTML characters in a string
     * @param {string} str - String to escape
     * @returns {string} - Escaped string
     */
    function escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&')
            .replace(/</g, '<')
            .replace(/>/g, '>')
            .replace(/"/g, '"')
            .replace(/'/g, '\'');
    }

    /**
     * Converts a cell value to HTML table cell
     * @param {string} cell - Cell value
     * @param {boolean} isHeader - Whether this is a header cell
     * @returns {string} - HTML table cell
     */
    function cellToHtml(cell, isHeader) {
        var numValue = parseFloat(cell.replace(/,/g, '.'));
        var isNumber = !isNaN(numValue) && cell.trim() !== '';
        
        var cellTag = isHeader ? 'th' : 'td';
        var style = isHeader ? ' style="background-color: #e9ecef; font-weight: bold; text-align: center; padding: 8px; border: 1px solid #dee2e6;"' : ' style="padding: 8px; border: 1px solid #dee2e6;"';
        
        if (isNumber && cell.match(/^-?\d+([.,]\d+)?$/)) {
            return '<' + cellTag + style + '>' + numValue + '</' + cellTag + '>';
        } else {
            return '<' + cellTag + style + '>' + escapeHtml(cell) + '</' + cellTag + '>';
        }
    }

    /**
     * Creates an HTML document string from table data
     * @param {Object} tableData - Object with headers and rows
     * @param {string} title - Title for the document
     * @returns {string} - HTML document string
     */
    function createDocHtml(tableData, title) {
        var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
            'xmlns:w="urn:schemas-microsoft-com:office:word" ' +
            'xmlns="http://www.w3.org/TR/REC-html40">' +
            '<head><meta charset="utf-8"><title>' + escapeHtml(title) + '</title>' +
            '<style>' +
            'body { font-family: "Times New Roman", serif; font-size: 12pt; } ' +
            'table { border-collapse: collapse; width: 100%; margin: 20px 0; } ' +
            'td, th { border: 1px solid #000; padding: 5px; } ' +
            '</style></head><body>';
        
        html += '<h2>' + escapeHtml(title) + '</h2>';
        html += '<table>';
        
        // Check if we have multi-row headers (headerRows array)
        if (tableData.headerRows && tableData.headerRows.length > 0) {
            // Add each header row
            for (var hr = 0; hr < tableData.headerRows.length; hr++) {
                var headerRow = tableData.headerRows[hr];
                html += '<tr>';
                for (var h = 0; h < headerRow.length; h++) {
                    html += cellToHtml(headerRow[h], true);
                }
                html += '</tr>';
            }
        } else if (tableData.headers && tableData.headers.length > 0) {
            // Single row header (backward compatibility)
            html += '<tr>';
            for (var i = 0; i < tableData.headers.length; i++) {
                html += cellToHtml(tableData.headers[i], true);
            }
            html += '</tr>';
        }

        // Add data rows
        for (var j = 0; j < tableData.rows.length; j++) {
            var row = tableData.rows[j];
            html += '<tr>';
            for (var k = 0; k < row.length; k++) {
                html += cellToHtml(row[k], false);
            }
            html += '</tr>';
        }

        html += '</table></body></html>';
        return html;
    }

    /**
     * Downloads the DOC file
     * @param {string} content - DOC HTML content
     * @param {string} filename - Name of the file to download
     */
    function downloadDocFile(content, filename) {
        var blob = new Blob([content], { type: 'application/msword;charset=utf-8' });
        var link = document.createElement('a');
        
        if (navigator.msSaveBlob) {
            // IE 10+
            navigator.msSaveBlob(blob, filename);
        } else {
            // Modern browsers
            var url = URL.createObjectURL(blob);
            link.href = url;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }
    }

    /**
     * Exports table data to DOC file
     * @param {Object} tableData - Object with headers and rows
     * @param {string} filename - Optional filename (without extension)
     */
    function exportToDoc(tableData, filename) {
        var defaultName = 'table_export_' + new Date().toISOString().slice(0, 10);
        var finalFilename = (filename || defaultName) + '.doc';
        
        var docContent = createDocHtml(tableData, finalFilename.replace('.doc', ''));
        downloadDocFile(docContent, finalFilename);
        
        console.log('[DocExporter] Exported table to:', finalFilename);
    }

    /**
     * Exports multiple tables to a single DOC file
     * @param {Array} tables - Array of objects with tableData and optional name
     * @param {string} filename - Optional filename (without extension)
     */
    function exportMultipleToDoc(tables, filename) {
        var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
            'xmlns:w="urn:schemas-microsoft-com:office:word" ' +
            'xmlns="http://www.w3.org/TR/REC-html40">' +
            '<head><meta charset="utf-8"><title>Экспорт таблиц</title>' +
            '<style>' +
            'body { font-family: "Times New Roman", serif; font-size: 12pt; } ' +
            'h2 { margin-top: 30px; color: #333; } ' +
            'table { border-collapse: collapse; width: 100%; margin: 20px 0; } ' +
            'td, th { border: 1px solid #000; padding: 5px; } ' +
            '</style></head><body>';
        
        for (var t = 0; t < tables.length; t++) {
            var table = tables[t];
            var tableData = table.tableData;
            var tableName = table.name || 'Таблица ' + (t + 1);

            html += '<h2>' + escapeHtml(tableName) + '</h2>';
            html += '<table>';

            // Check if we have multi-row headers (headerRows array)
            if (tableData.headerRows && tableData.headerRows.length > 0) {
                // Add each header row
                for (var hr = 0; hr < tableData.headerRows.length; hr++) {
                    var headerRow = tableData.headerRows[hr];
                    html += '<tr>';
                    for (var h = 0; h < headerRow.length; h++) {
                        html += cellToHtml(headerRow[h], true);
                    }
                    html += '</tr>';
                }
            } else if (tableData.headers && tableData.headers.length > 0) {
                // Single row header (backward compatibility)
                html += '<tr>';
                for (var i = 0; i < tableData.headers.length; i++) {
                    html += cellToHtml(tableData.headers[i], true);
                }
                html += '</tr>';
            }

            // Add tooltip row after header (if present)
            if (tableData.tooltipRow && tableData.tooltipRow.length > 0) {
                var hasTooltip = false;
                for (var ti = 0; ti < tableData.tooltipRow.length; ti++) {
                    if (tableData.tooltipRow[ti] && tableData.tooltipRow[ti].trim() !== '') {
                        hasTooltip = true;
                        break;
                    }
                }
                if (hasTooltip) {
                    html += '<tr>';
                    for (var tj = 0; tj < tableData.tooltipRow.length; tj++) {
                        html += cellToHtml(tableData.tooltipRow[tj], false);
                    }
                    html += '</tr>';
                }
            }

            // Add data rows
            for (var r = 0; r < tableData.rows.length; r++) {
                var row = tableData.rows[r];
                html += '<tr>';
                for (var c = 0; c < row.length; c++) {
                    html += cellToHtml(row[c], false);
                }
                html += '</tr>';
            }

            html += '</table>';
        }

        html += '</body></html>';

        var defaultName = 'tables_export_' + new Date().toISOString().slice(0, 10);
        var finalFilename = (filename || defaultName) + '.doc';
        
        downloadDocFile(html, finalFilename);
        
        console.log('[DocExporter] Exported', tables.length, 'tables to:', finalFilename);
    }

    return {
        exportToDoc: exportToDoc,
        exportMultipleToDoc: exportMultipleToDoc,
        createDocHtml: createDocHtml
    };
})();