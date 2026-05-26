/**
 * Excel Exporter Module
 * Handles exporting table data to Excel format (.xlsx)
 */

var ExcelExporter = (function() {
    'use strict';

    /**
     * Escapes special XML characters in a string
     * @param {string} str - String to escape
     * @returns {string} - Escaped string
     */
    function escapeXml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&')
            .replace(/</g, '<')
            .replace(/>/g, '>')
            .replace(/"/g, '"')
            .replace(/'/g, '\'');
    }

    /**
     * Converts a cell value to Excel XML cell element
     * @param {string} cell - Cell value
     * @returns {string} - Excel XML cell element
     */
    function cellToXml(cell, attributes) {
        var numValue = parseFloat(cell.replace(/,/g, '.'));
        var isNumber = !isNaN(numValue) && cell.trim() !== '';
        var attrText = attributes || '';

        if (isNumber && cell.match(/^-?\d+([.,]\d+)?$/)) {
            return '<Cell' + attrText + '><Data ss:Type="Number">' + numValue + '</Data></Cell>';
        } else {
            return '<Cell' + attrText + '><Data ss:Type="String">' + escapeXml(cell) + '</Data></Cell>';
        }
    }

    /**
     * Creates Excel XML for data rows using metadata to preserve rowspan/colspan offsets
     * @param {Array[]} dataRows - Array of table data rows
     * @param {Array[]} dataCellMeta - Array of corresponding metadata rows
     * @returns {string} - Excel XML for rows
     */
    function createExcelDataRowsXml(dataRows, dataCellMeta) {
        if (!dataRows || dataRows.length === 0) {
            return '';
        }

        var xml = '';
        for (var r = 0; r < dataRows.length; r++) {
            var row = dataRows[r] || [];
            var metaRow = dataCellMeta && dataCellMeta[r] ? dataCellMeta[r] : [];
            var currentCol = 1;
            var outputCount = 0;
            var columns = Math.max(row.length, metaRow.length || 0);

            xml += '<Row>';
            for (var c = 0; c < columns; c++) {
                var meta = metaRow[c] || { colspan: 1, rowspan: 1, isOriginal: true };

                if (meta.isOriginal === false) {
                    currentCol += 1;
                    continue;
                }

                var attributes = '';
                if (outputCount === 0 && currentCol > 1) {
                    attributes += ' ss:Index="' + currentCol + '"';
                }
                if (meta.colspan > 1) {
                    attributes += ' ss:MergeAcross="' + (meta.colspan - 1) + '"';
                }
                if (meta.rowspan > 1) {
                    attributes += ' ss:MergeDown="' + (meta.rowspan - 1) + '"';
                }

                xml += cellToXml(row[c] || '', attributes);
                currentCol += (meta.colspan > 0 ? meta.colspan : 1);
                outputCount++;
            }
            xml += '</Row>';
        }

        return xml;
    }

    /**
     * Adds merged cells to Excel XML
     * @param {Array} merges - Array of merge objects
     * @returns {string} - Excel XML merge cells
     */
    function createExcelHeaderRowsXml(headerRows, headerCellMeta) {
        if (!headerRows || headerRows.length === 0) {
            return '';
        }

        var xml = '';
        for (var hr = 0; hr < headerRows.length; hr++) {
            var headerRow = headerRows[hr];
            var metaRow = headerCellMeta && headerCellMeta[hr] ? headerCellMeta[hr] : [];
            var currentCol = 1;
            var outputCount = 0;

            xml += '<Row>';
            for (var c = 0; c < headerRow.length; c++) {
                var meta = metaRow[c];

                // Skip cells that are covered by a span from an earlier row
                if (meta && meta.isOriginal === false) {
                    currentCol++;
                    continue;
                }

                var attributes = '';
                if (outputCount === 0 && currentCol > 1) {
                    attributes += ' ss:Index="' + currentCol + '"';
                }

                if (meta) {
                    if (meta.colspan > 1) {
                        attributes += ' ss:MergeAcross="' + (meta.colspan - 1) + '"';
                    }
                    if (meta.rowspan > 1) {
                        attributes += ' ss:MergeDown="' + (meta.rowspan - 1) + '"';
                    }
                }

                xml += '<Cell' + attributes + '><Data ss:Type="String">' + escapeXml(headerRow[c] || '') + '</Data></Cell>';
                currentCol += (meta && meta.colspan > 0 ? meta.colspan : 1);
                outputCount++;
            }
            xml += '</Row>';
        }

        return xml;
    }

    function addMergesToXml(merges) {
        if (!merges || merges.length === 0) return '';
        
        var xml = '<MergeCells>';
        for (var m = 0; m < merges.length; m++) {
            var merge = merges[m];
            var startCol = getExcelColumn(merge.col);
            var endCol = getExcelColumn(merge.col + (merge.cells || 1) - 1);
            // Adjust row for multi-row headers: row 0 is Excel row 1, row 1 is Excel row 2, etc.
            var startRow = merge.row + 1;
            // If rowspan > 1, extend the merge vertically
            var endRow = startRow + (merge.rows || 1) - 1;
            xml += '<MergeCell ss:Range="' + startCol + startRow + ':' + endCol + endRow + '"/>';
        }
        xml += '</MergeCells>';
        return xml;
    }

    /**
     * Creates an Excel XML string from table data
     * @param {Object} tableData - Object with headers and rows
     * @param {string} sheetName - Name of the sheet
     * @returns {string} - Excel XML string
     */
    function createExcelXML(tableData, sheetName) {
        var xml = '<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>';
        xml += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ';
        xml += 'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
        xml += '<Worksheet ss:Name="' + escapeXml(sheetName) + '"><Table>';

        // Check if we have multi-row headers (headerRows array)
        if (tableData.headerRows && tableData.headerRows.length > 0) {
            var headerMeta = tableData.cellMeta ? tableData.cellMeta.slice(0, tableData.headerRows.length) : [];
            xml += createExcelHeaderRowsXml(tableData.headerRows, headerMeta);
        } else if (tableData.headers && tableData.headers.length > 0) {
            // Single row header (backward compatibility)
            xml += '<Row>';
            for (var i = 0; i < tableData.headers.length; i++) {
                xml += '<Cell><Data ss:Type="String">' + escapeXml(tableData.headers[i]) + '</Data></Cell>';
            }
            xml += '</Row>';
        }

        // Add tooltip row after header (if present)
        if (tableData.tooltipRow && tableData.tooltipRow.length > 0) {
            var hasTooltip = false;
            for (var t = 0; t < tableData.tooltipRow.length; t++) {
                if (tableData.tooltipRow[t] && tableData.tooltipRow[t].trim() !== '') {
                    hasTooltip = true;
                    break;
                }
            }
            if (hasTooltip) {
                xml += '<Row>';
                for (var ti = 0; ti < tableData.tooltipRow.length; ti++) {
                    xml += '<Cell><Data ss:Type="String">' + escapeXml(tableData.tooltipRow[ti]) + '</Data></Cell>';
                }
                xml += '</Row>';
            }
        }

        // Add data rows using metadata so rowspan/colspan offsets are preserved
        var dataRowOffset = tableData.headerRows && tableData.headerRows.length > 0 ? tableData.headerRows.length : 0;
        var dataMeta = tableData.cellMeta ? tableData.cellMeta.slice(dataRowOffset) : [];
        xml += createExcelDataRowsXml(tableData.rows, dataMeta);

        // Add merged cells if present
        xml += addMergesToXml(tableData.merges, tableData.headerLevels);

        xml += '</Table></Worksheet></Workbook>';
        return xml;
    }

    // Helper function to convert column index to Excel column letter
    function getExcelColumn(columnIndex) {
        var result = '';
        var col = columnIndex;
        while (col >= 0) {
            result = String.fromCharCode(65 + (col % 26)) + result;
            col = Math.floor(col / 26) - 1;
        }
        return result;
    }

    /**
     * Downloads the Excel file
     * @param {string} content - Excel XML content
     * @param {string} filename - Name of the file to download
     */
    function downloadExcelFile(content, filename) {
        var blob = new Blob([content], { type: 'application/vnd.ms-excel;charset=utf-8' });
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
     * Exports table data to Excel file
     * @param {Object} tableData - Object with headers and rows
     * @param {string} filename - Optional filename (without extension)
     */
    function exportToExcel(tableData, filename) {
        var defaultName = 'table_export_' + new Date().toISOString().slice(0, 10);
        var finalFilename = (filename || defaultName) + '.xls';
        
        var excelContent = createExcelXML(tableData, 'Sheet1');
        downloadExcelFile(excelContent, finalFilename);
        
        console.log('[ExcelExporter] Exported table to:', finalFilename);
    }

    /**
     * Exports multiple tables to a single Excel file with multiple sheets
     * @param {Array} tables - Array of objects with tableData and optional sheetName
     * @param {string} filename - Optional filename (without extension)
     */
    function exportMultipleToExcel(tables, filename) {
        var xml = '<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>';
        xml += '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ';
        xml += 'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';

        for (var t = 0; t < tables.length; t++) {
            var table = tables[t];
            var sheetName = table.sheetName || 'Sheet' + (t + 1);
            var tableData = table.tableData;

            xml += '<Worksheet ss:Name="' + escapeXml(sheetName) + '"><Table>';

            // Check if we have multi-row headers (headerRows array)
            if (tableData.headerRows && tableData.headerRows.length > 0) {
                var headerMeta = tableData.cellMeta ? tableData.cellMeta.slice(0, tableData.headerRows.length) : [];
                xml += createExcelHeaderRowsXml(tableData.headerRows, headerMeta);
            } else if (tableData.headers && tableData.headers.length > 0) {
                // Single row header (backward compatibility)
                xml += '<Row>';
                for (var h = 0; h < tableData.headers.length; h++) {
                    xml += '<Cell><Data ss:Type="String">' + escapeXml(tableData.headers[h]) + '</Data></Cell>';
                }
                xml += '</Row>';
            }

            // Add data rows using metadata so rowspan/colspan offsets are preserved
            var dataRowOffset = tableData.headerRows && tableData.headerRows.length > 0 ? tableData.headerRows.length : 0;
            var dataMeta = tableData.cellMeta ? tableData.cellMeta.slice(dataRowOffset) : [];
            xml += createExcelDataRowsXml(tableData.rows, dataMeta);

            // Add merged cells if present
            xml += addMergesToXml(tableData.merges, tableData.headerLevels);

            xml += '</Table></Worksheet>';
        }

        xml += '</Workbook>';

        var defaultName = 'tables_export_' + new Date().toISOString().slice(0, 10);
        var finalFilename = (filename || defaultName) + '.xls';
        
        downloadExcelFile(xml, finalFilename);
        
        console.log('[ExcelExporter] Exported', tables.length, 'tables to:', finalFilename);
    }

    return {
        exportToExcel: exportToExcel,
        exportMultipleToExcel: exportMultipleToExcel,
        createExcelXML: createExcelXML
    };
})();