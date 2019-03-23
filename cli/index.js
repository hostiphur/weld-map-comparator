'use strict';

const WeldMapComparator = require('../weld-map-comparator.js');
const args = process.argv.slice(2);
let inspectorWorkbookPath;
let surveyWorkbookPath;

[inspectorWorkbookPath, surveyWorkbookPath] = args;

function printUsage() {
    console.log(`
    Please provide the full path to the inspector workbook as the first parameter, in double quotes.
    Please provide the full path to the survey workbook as the second parameter, in double quotes.
    `);
};

if (!inspectorWorkbookPath || !surveyWorkbookPath) {
    printUsage();
} else {
    const wmc = new WeldMapComparator(inspectorWorkbookPath, surveyWorkbookPath);
    wmc.process();
}
