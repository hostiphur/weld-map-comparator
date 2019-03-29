const { ipcRenderer } = require('electron');
const { dialog } = require('electron').remote;
const events = require('../events');

const internals = {};

module.exports = {
    init() {
        ipcRenderer.on(events.COMPARISON_STATUS_LOG, (event, logMessage) => {
            const message = `${logMessage}<br/>`;
            document.querySelector('#outputArea').innerHTML += message;
        });
    },
    setOutputText(text) {
        document.querySelector('#outputArea').innerHTML = text;
    },
    setupClickHandlerForRunComparison(buttonID) {
        document.querySelector(buttonID).addEventListener('click', function (event) {
            const inspectorWorkbookPath = (document.querySelector('#weld-map-file-path').value || '').trim();
            const surveyWorkbookPath = (document.querySelector('#survey-file-path').value || '').trim();

            if (inspectorWorkbookPath.length == 0 || surveyWorkbookPath.length == 0) {
                document.querySelector('#outputArea').innerHTML = 'Please choose both an inspector and survey workbook';
                return;
            }

            module.exports.setOutputText('Please wait. Processing...');
            ipcRenderer.send(events.RUN_COMPARISON, {inspectorWorkbookPath, surveyWorkbookPath});
        });
    },
    setupClickHandlerForFileLoad(buttonID, textPathID) {
        document.querySelector(buttonID).addEventListener('click', function (event) {
            dialog.showOpenDialog({
                properties: ['openFile']
            }, function (files) {
                if (files !== undefined) {
                    document.querySelector(textPathID).value = files[0];
                }
            });
        });
    }
}
