const { ipcRenderer } = require('electron');
const { shell } = require("electron");
const XLSXChart = require("xlsx-chart");
var fs = require("fs");

let extractedData = [];
let bitrateData = [];
let indexArray = []; // This will store the array of indices
let inputText = "";
function showAlert(message) {
  const alertOverlay = document.getElementById("alertOverlay");
  const alertMessage = document.getElementById("alertMessage");
  const alertCloseBtn = document.getElementById("alertCloseBtn");
  alertMessage.innerHTML = message.replace("\n", "<br>");
  alertOverlay.style.display = "flex";
  alertCloseBtn.addEventListener("click", () => {
    alertOverlay.style.display = 'none';
    closeAlert();
  });
}
function closeAlert() {
  const alertOverlay = document.getElementById("alertOverlay");
  alertOverlay.style.display = "none";
}
function openFile() {
  ipcRenderer.send("open-file-dialog");
    
}
function openSaveDialog4() {
  ipcRenderer.send("open-save-dialog4");
}
function saveFile() {
  extractTrafficData(inputText);
  const loadingOverlay = document.getElementById("loadingOverlay");
  loadingOverlay.style.display = "flex";
  if (!window.currentFilePath4) {
    showAlert("ファイルパスを入力してください！");
    loadingOverlay.style.display = "none";
  } else {
    createExcelFile(window.currentFilePath4);
  }
}
ipcRenderer.on("file-content", (event, content) => {
  inputText = content;
  console.log("Input Data: ",inputText)
});

function displayFilePath(filePath, elementId) {
  const filePathElement = document.getElementById(elementId);
  const filePathLink = document.createElement("a");
  filePathLink.href = "#";
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");

  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
}

ipcRenderer.on("file-path", (event, filePath) => {
  
  const filePathElement = document.getElementById("uploaded-file-path");
  const filePathLink = document.createElement("a");
  filePathLink.href = "#";
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");
  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
});

ipcRenderer.on("file-created-4", (event, filePath) => {
  window.currentFilePath4 = filePath;
  const filePathElement = document.getElementById("created-file4");
  const filePathLink = document.createElement("a");
  filePathLink.href = "#";
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");
  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
});

function extractTrafficData(logText) {
  extractedData = [];
  bitrateData = [];
  indexArray = [];
  const sections = logText.split(/----- トラフィックログ取得日時 /).slice(1);
  sections.forEach((section) => {
    const lines = section.split("\n");
    
    // Extract Timestamp
    const timestampMatch = lines[0].match(/(\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}:\d{2})/);
    const timestamp = timestampMatch ? timestampMatch[1] : null;

    // Extract "瞬間受信速度" (preserve .0) and rename to "bitrate"
    const bitrateMatch = section.match(/瞬間受信速度\s+([\d.]+) Kbps/);
    const bitrate = bitrateMatch ? bitrateMatch[1] : null; // Keep it as a string

    if (timestamp && bitrate !== null) {
      extractedData.push({ timestamp, bitrate }); // Store the full data
      bitrateData.push(bitrate); // Store only the bitrate value
    }
  });

  indexArray = Array.from({ length: extractedData.length }, (_, i) => i);
  return { extractedData, bitrateData, indexArray }; // Return all three arrays
}
function refreshBeforeUpload() {
  location.reload(); // Reload the entire page
}
function createExcelFile(filePath) {
  var xlsxChart = new XLSXChart();
  console.log("Bitrate: ",  bitrateData)
  //Bitrate object
  const bitrateObject = {
    Bitrate: bitrateData.reduce((acc, bitrate, index) => {
      acc[index] = parseFloat(bitrate); // Store bitrate values as numbers
      return acc;
    }, {})
  };
  // Options object to configure the chart and data
  var opts = {
    charts: [
      {
        position: {
          fromColumn: 0,
          toColumn: 25,
          fromRow: 1,
          toRow: 24,
        },
        customColors: {
          points: {
              "Bitrate": {
                  "Birate": 'ff0000',
              },
          },
          series: {
              "Bitrate": {
                  fill: 'ff0000',
                  line: 'ff0000',
              }
          }
        },
        chart: 'line',
        titles: ['Bitrate'],
        fields: indexArray,
        data: bitrateObject,
        chartTitle: 'Bitrate', //Special
        lineWidth: 0.2,
      }
    ]
  };
  // Generate the chart and save the Excel file
  xlsxChart.generate(opts, function (err, data) {
      const loadingOverlay = document.getElementById("loadingOverlay");
      if (err) {
          console.error(err);
      } else {
          loadingOverlay.style.display = "flex";
          fs.writeFileSync(filePath, data);
          loadingOverlay.style.display = "none";
          executePythonScript(filePath);
          showAlert("ファイルダウンロードを保存しました");
          console.log('Excel file with column chart created successfully at:', filePath);
      }
  });
    // Execute Python script after Excel file generation
    const { exec } = require('child_process');
    function executePythonScript(filePath) {
      // Execute the Python script
      const pythonScriptPath = 'python.py'; // Update with the actual name of your Python script
      exec(`python ${pythonScriptPath} "${filePath}"`, (error, stdout, stderr) => {
          if (error) {
              console.error(`Error executing Python script: ${error}`);
              return;
          }
          console.log(`Python script output: ${stdout}`);
      });
    }
}
