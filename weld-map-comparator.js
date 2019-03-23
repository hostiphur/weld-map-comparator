'use strict';

const config = require('./config/config-loader.js');
const XLSX = require('xlsx-style');
const fs = require('fs');
const path = require('path');

class WeldMapComparator {
    constructor(inspectorFilePath, surveyFilePath) {
        this.inspectorFilePath = inspectorFilePath;
        this.surveyFilePath = surveyFilePath;
        this.comparisonResults = [];
        this.inspectorData = [];
        this.inspectorWorkbook = undefined;
        this.surveyData = [];
        this.comparisonResultsStartingColumn = 19;
    }

    async streamFile(filePath){
        let buffers = [];
        return new Promise((resolve) => {
            const stream = fs.createReadStream(filePath);
            stream.on('data', function(data) { buffers.push(data); });
            stream.on('end', function() {
                let buffer = Buffer.concat(buffers);
                let workbook = XLSX.read(buffer, {
                    type:"buffer",
                    cellStyles: true,
                    cellNF: true
                });
                resolve(workbook);
            });
        });
    }

    convertSheetToArray(sheet, startingRow, columnMapping){
        let result = [];
        let rowNum;
        let range = XLSX.utils.decode_range(sheet['!ref']);

        for(rowNum = startingRow; rowNum <= range.e.r; rowNum++){

            const entry = {
                rowNumber: rowNum
            };

            Object.keys(columnMapping).forEach(columnKey => {
                const cellAddress = `${columnMapping[columnKey]}${rowNum+1}`;
                const nextCell = sheet[cellAddress];

                if( typeof nextCell === 'undefined' ){
                    entry[columnKey] = undefined;
                } else {
                    entry[columnKey] = nextCell.w;
                }
            });

            result.push(entry);
        }
        return result;
    };

    async loadExcelFile(filePath) {
        console.log(`>>>>>>> ${filePath}`);

        return await this.streamFile(filePath);
    }

    getWorksheet(workbook, sheetNumber) {
        const sheetName = workbook.SheetNames[sheetNumber];
        return workbook.Sheets[sheetName];
    }

    async loadFiles() {
        this.inspectorWorkbook = await this.loadExcelFile(this.inspectorFilePath);
        this.surveyWorkbook = await this.loadExcelFile(this.surveyFilePath);

        let worksheet = this.getWorksheet(this.inspectorWorkbook, config.inspectorWeldMappingSheet.worksheetIndex);

        this.inspectorData = this.convertSheetToArray(
            worksheet,
            config.inspectorWeldMappingSheet.startingRow,
            config.inspectorWeldMappingSheet.columnMapping);

        worksheet = this.getWorksheet(this.surveyWorkbook, config.surveySheet.worksheetIndex);

        this.surveyData = this.convertSheetToArray(
            worksheet,
            config.surveySheet.startingRow,
            config.surveySheet.columnMapping);
    }

    async process() {
        await this.loadFiles();
        this.compareFiles();
        this.markupWorkbook();
        this.writeMarkedUpWorkbook();
    }

    ensureAdequateWorksheetRange(worksheet) {
        const outputColumns = config.outputSheet.columns;
        const worksheetRange = XLSX.utils.decode_range(worksheet['!ref']);

        let maxColumn = 0;

        Object.keys(outputColumns).forEach(columnKey => {
            const columnIndex = XLSX.utils.decode_col(outputColumns[columnKey].columnMapping);
            if (columnIndex > maxColumn) {
                maxColumn = columnIndex;
            }
        });

        if (maxColumn > worksheetRange.e.c) {
            worksheetRange.e.c = maxColumn;
        }

        worksheet['!ref'] = XLSX.utils.encode_range(worksheetRange);
    }

    markupWorkbook() {
        const workbook = this.inspectorWorkbook;
        const sheetName = workbook.SheetNames[config.inspectorWeldMappingSheet.worksheetIndex];
        const worksheet = workbook.Sheets[sheetName];
        const colorCodes = config.outputSheet.colorCodes;
        const outputColumns = config.outputSheet.columns;

        this.ensureAdequateWorksheetRange(worksheet);

        // Output column headers
        Object.keys(outputColumns).forEach(columnKey => {
            const columnIndex = XLSX.utils.decode_col(outputColumns[columnKey].columnMapping);
            const sheetCoordinates = XLSX.utils.encode_cell({r: outputColumns[columnKey].headerRow, c: columnIndex});
            worksheet[sheetCoordinates] = {
                t: 's',
                v: outputColumns[columnKey].headerText
            };
        });

        // Output comparison results
        this.comparisonResults.forEach(result => {
            let style;
            if (result.notFound === true) {
                style = {patternType: 'solid', fill: {fgColor: {rgb: colorCodes.recordNotFound}}};
            } else if (result.noReferenceNumber === true) {
                style = {patternType: 'solid', fill: {fgColor: {rgb: colorCodes.noReferenceNumber}}};
            } else if (result.discrepancy === true) {
                style = {patternType: 'solid', fill: {fgColor: {rgb: colorCodes.discrepancy}}};
            }

            // For each comparison result, look for data from each column of interest
            Object.keys(outputColumns).forEach(columnKey => {
                if (result[columnKey]) {
                    const columnIndex = XLSX.utils.decode_col(outputColumns[columnKey].columnMapping);
                    const sheetCoordinates = XLSX.utils.encode_cell({r: result.inspectorRow, c: columnIndex});
                    worksheet[sheetCoordinates] = {
                        t: 's',
                        v: result[columnKey]
                    }
                    if (style) {
                        worksheet[sheetCoordinates].s = style;
                    }
                }
            });
        });
    }

    writeMarkedUpWorkbook() {
        let fileName = config.outputSheet.name;
        fileName = fileName.replace('{{original-file-name}}', path.basename(this.inspectorFilePath));
        XLSX.writeFile(this.inspectorWorkbook, `./out/${fileName}`);
    }

    compareFiles() {

        this.inspectorData.forEach((iRow) => {

            let curOrNumber = iRow.orNumber || '';

            // Ignore rows with no OR number
            if (curOrNumber.trim() === '') {
                this.pushIgnoredRow(iRow, false, true);
                return;
            }

            // Replace <zero>R with OR. For fat fingered typists
            curOrNumber = curOrNumber.replace('0R', 'OR');
            const orBreakdown = curOrNumber.match(/(OR)([0-9]+)([a-zA-Z]*)/);

            if (!orBreakdown) {
                console.log(`[${this.inspectorFilePath}:${iRow.rowNumber}] Invalid OR number format detected: ${curOrNumber}`);
                this.pushIgnoredRow(iRow, false, true);
                return;
            }

            /*
            After the OR number is split apart by the above regular expression it will be in an array.
                Example: OR235B -> ['OR235B', 'OR', '235', 'B']
                Example: OR235 -> ['OR235B', 'OR', '235', null]
            */
            iRow.orBreakdown = orBreakdown;
            const orNum = parseInt(orBreakdown[2]);

            //Try to find the entry by index
            if (this.surveyData[orNum] &&
                this.rowsMatchByORNumber(iRow, this.surveyData[orNum])) {
                this.compareRows(iRow, this.surveyData[orNum]);
                return;
            }
            
            // Look for a match in surrounding rows
            let index = orNum - 5;
            if (index < 0) {
                index = 0;
            }

            for (let i = 0; i < 10; i++) {
                const surveyRow = this.surveyData[index + i];
                if (!surveyRow) {
                    break;
                }
                if (this.rowsMatchByORNumber(iRow, surveyRow)) {
                    this.compareRows(iRow, surveyRow)
                    return;
                }
            }

            // Look through the whole damn thing from top to bottom
            this.surveyData.forEach((surveyRow) => {
                if (this.rowsMatchByORNumber(iRow, surveyRow)) {
                    this.compareRows(iRow, surveyRow)
                    return;
                }
            });

            this.pushIgnoredRow(iRow, true, false);

            console.log(`OOF!!! no matching OR number found for ${iRow.orNumber}`);
            return;
        });
    }

    pushIgnoredRow(inspectorRow, notFound, noReferenceNumber) {
        this.comparisonResults.push({
            discrepancy: true,
            notFound,
            noReferenceNumber,
            inspectorRow: inspectorRow.rowNumber,
            surveyRow: -1,
            pipeNumber: inspectorRow.pipeNumber,
            heatNumber: inspectorRow.heatNumber
        });
    }

    getStringDifference(string1 = "", string2 = "") {
        if (string1.length > string2.length) {
            return string1.replace(string2, '');
        }
        return string2.replace(string1, '');
    }

    rowsMatchByORNumber(inspectorRow, surveyRow) {
        let ior = inspectorRow.orBreakdown[2] || '';
        let sor = surveyRow.orNumber || '';

        ior = ior.replace(/^0+/, '');
        sor = sor.replace(/^0+/, '');

        if (ior.trim() === sor.trim()) {
            return true;
        }

        return false;
    }

    comparePipeNumbers(inspectorRow, surveyRow) {
        const sor = surveyRow.orNumber || '';
        const ipn = inspectorRow.pipeNumber || '';
        const spn = surveyRow.pipeNumber || '';

        const stringDiff = this.getStringDifference(ipn.toLowerCase().trim(), spn.toLowerCase().trim());
        const letter = inspectorRow.orBreakdown[3] || "no-letter-appended";
        const differByLetterOnly = stringDiff === letter.toLowerCase();

        if (stringDiff.length > 0 && differByLetterOnly === false){
            console.log(`[${inspectorRow.rowNumber}:${surveyRow.rowNumber}] pipe number discrepancy found: i-${inspectorRow.orNumber} / s-${sor} ... ${ipn} / ${spn}`);

            return {
                discrepancy: true,
                inspectorRow: inspectorRow.rowNumber,
                surveyRow: surveyRow.rowNumber,
                pipeNumber: spn
            };
        }

        return {
            discrepancy: false,
            inspectorRow: inspectorRow.rowNumber,
            surveyRow: surveyRow.rowNumber,
            pipeNumber: ipn
        };
    }

    compareHeatNumbers(inspectorRow, surveyRow) {
        const sor = surveyRow.orNumber || '';
        const ihn = inspectorRow.heatNumber || '';
        const shn = surveyRow.heatNumber || '';

        if (ihn.toLowerCase().trim() !== shn.toLowerCase().trim()){

            console.log(`[${inspectorRow.rowNumber}:${surveyRow.rowNumber}] heat number discrepancy found: i-${inspectorRow.orNumber} / s-${sor} ... ${ihn} / ${shn}`);

            return {
                discrepancy: true,
                inspectorRow: inspectorRow.rowNumber,
                surveyRow: surveyRow.rowNumber,
                heatNumber: shn
            };
        }

        return {
            discrepancy: false,
            inspectorRow: inspectorRow.rowNumber,
            surveyRow: surveyRow.rowNumber,
            heatNumber: ihn
        };
    }

    compareRows(inspectorRow, surveyRow) {
        this.comparisonResults.push(this.comparePipeNumbers(inspectorRow, surveyRow));
        this.comparisonResults.push(this.compareHeatNumbers(inspectorRow, surveyRow));
    }
}

module.exports = WeldMapComparator;
