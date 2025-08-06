<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel ↔ Google Sheets Formula Converter</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS library for reading and writing Excel files -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Inter', sans-serif;
        }
        .table-cell {
            padding: 12px 16px;
            border: 1px solid #e2e8f0;
            font-size: 0.875rem;
            white-space: pre-wrap;
            word-break: break-all;
        }
        .table-header {
            background-color: #f8fafc;
            font-weight: 600;
            color: #475569;
        }
        .copy-btn {
            cursor: pointer;
            opacity: 0.5;
            transition: opacity 0.2s;
        }
        .copy-btn:hover {
            opacity: 1;
        }
    </style>
</head>
<body class="bg-gray-50 text-gray-800">

    <div class="container mx-auto p-4 sm:p-6 md:p-8 max-w-7xl">
        <!-- Header Section -->
        <header class="text-center mb-8">
            <h1 class="text-3xl sm:text-4xl font-bold text-gray-900">Formula Converter</h1>
            <p class="mt-2 text-lg text-gray-600">Translate between Excel and Google Sheets functions instantly.</p>
        </header>

        <!-- Main Application Card -->
        <div class="bg-white rounded-xl shadow-lg p-6 sm:p-8 border border-gray-200">
            
            <!-- Step 1: Upload -->
            <div class="mb-6">
                <h2 class="text-xl font-semibold text-gray-700 mb-3">1. Upload your file</h2>
                <div class="flex items-center justify-center w-full">
                    <label for="file-upload" class="flex flex-col items-center justify-center w-full h-48 border-2 border-gray-300 border-dashed rounded-lg cursor-pointer bg-gray-50 hover:bg-gray-100 transition-colors">
                        <div class="flex flex-col items-center justify-center pt-5 pb-6">
                            <svg class="w-10 h-10 mb-4 text-gray-500" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 20 16"><path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 13h3a3 3 0 0 0 0-6h-.025A5.56 5.56 0 0 0 16 6.5 5.5 5.5 0 0 0 5.207 5.021C5.137 5.017 5.071 5 5 5a4 4 0 0 0 0 8h2.167M10 15V6m0 0L8 8m2-2 2 2"/></svg>
                            <p class="mb-2 text-sm text-gray-500"><span class="font-semibold">Click to upload</span> or drag and drop</p>
                            <p class="text-xs text-gray-500">XLSX file</p>
                        </div>
                        <input id="file-upload" type="file" class="hidden" accept=".xlsx"/>
                    </label>
                </div>
                <p id="file-name" class="mt-2 text-center text-sm text-gray-500"></p>
            </div>

            <!-- Step 2: Choose Conversion -->
            <div class="mb-6">
                <h2 class="text-xl font-semibold text-gray-700 mb-3">2. Choose conversion direction</h2>
                <fieldset class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <label for="to-google" class="flex items-center p-4 border border-gray-300 rounded-lg cursor-pointer hover:bg-blue-50 hover:border-blue-400 transition-all">
                        <input type="radio" id="to-google" name="conversion-direction" value="to-google" class="h-5 w-5 text-blue-600 border-gray-300 focus:ring-blue-500" checked>
                        <span class="ml-3 text-md font-medium text-gray-700">Excel → Google Sheets</span>
                    </label>
                    <label for="to-excel" class="flex items-center p-4 border border-gray-300 rounded-lg cursor-pointer hover:bg-green-50 hover:border-green-400 transition-all">
                        <input type="radio" id="to-excel" name="conversion-direction" value="to-excel" class="h-5 w-5 text-green-600 border-gray-300 focus:ring-green-500">
                        <span class="ml-3 text-md font-medium text-gray-700">Google Sheets → Excel</span>
                    </label>
                </fieldset>
            </div>

            <!-- Step 3: Convert -->
            <div class="text-center">
                 <button id="convert-btn" class="w-full sm:w-auto bg-blue-600 text-white font-bold py-3 px-8 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-300 transition-all disabled:bg-gray-400 disabled:cursor-not-allowed" disabled>
                    Convert
                </button>
            </div>
        </div>

        <!-- Results Section -->
        <div id="results-container" class="mt-8 hidden">
            <h2 class="text-2xl font-bold text-gray-900 mb-4">Conversion Results</h2>
            <div id="results-output" class="bg-white rounded-xl shadow-lg border border-gray-200 overflow-hidden">
                <!-- Results will be injected here -->
            </div>
        </div>
        
        <!-- Message Box for copy notifications -->
        <div id="message-box" class="fixed bottom-5 right-5 bg-gray-900 text-white py-2 px-4 rounded-lg shadow-xl transition-opacity duration-300 opacity-0">
            Copied to clipboard!
        </div>

    </div>

    <script>
        // --- DOM Element References ---
        const fileUpload = document.getElementById('file-upload');
        const fileNameDisplay = document.getElementById('file-name');
        const convertBtn = document.getElementById('convert-btn');
        const resultsContainer = document.getElementById('results-container');
        const resultsOutput = document.getElementById('results-output');
        const messageBox = document.getElementById('message-box');
        let workbook = null;

        // --- Mappings for Formula Conversions ---
        // This is a simplified mapping. A real-world version would be much more extensive.
        const excelToGoogleMap = {
            'CONCAT': 'CONCATENATE',
            'TEXTJOIN': 'TEXTJOIN', // Same name, but good to have it explicit
            'XLOOKUP': 'XLOOKUP', // Now supported in Google Sheets
            'IFS': 'IFS',
            'MAXIFS': 'MAXIFS',
            'MINIFS': 'MINIFS',
            // Add more Excel-specific functions that have Google Sheets equivalents
        };

        const googleToExcelMap = {
            'CONCATENATE': 'CONCAT',
            'QUERY': null, // No direct equivalent in Excel, mark as untranslatable
            'IMPORTRANGE': null,
            'GOOGLEFINANCE': null,
            // Add more Google Sheets-specific functions
        };

        // --- Event Listeners ---
        fileUpload.addEventListener('change', handleFile, false);
        convertBtn.addEventListener('click', convertFile, false);

        // --- Functions ---

        /**
         * Handles the file upload event.
         * Reads the file using SheetJS and enables the convert button.
         */
        function handleFile(e) {
            const files = e.target.files;
            if (files.length === 0) return;

            const file = files[0];
            fileNameDisplay.textContent = `Selected file: ${file.name}`;
            convertBtn.disabled = false;

            const reader = new FileReader();
            reader.onload = (event) => {
                const data = new Uint8Array(event.target.result);
                // Store the workbook object for later use
                workbook = XLSX.read(data, { type: 'array', cellFormula: true });
            };
            reader.readAsArrayBuffer(file);
        }

        /**
         * Main function to trigger the conversion process.
         */
        function convertFile() {
            if (!workbook) {
                alert("Please upload a file first.");
                return;
            }

            const direction = document.querySelector('input[name="conversion-direction"]:checked').value;
            const conversionMap = direction === 'to-google' ? excelToGoogleMap : googleToExcelMap;
            
            let resultsHtml = `<table class="w-full">
                                 <thead>
                                   <tr>
                                     <th class="table-cell table-header text-left">Sheet</th>
                                     <th class="table-cell table-header text-left">Cell</th>
                                     <th class="table-cell table-header text-left">Original Formula</th>
                                     <th class="table-cell table-header text-left">Converted Formula</th>
                                   </tr>
                                 </thead>
                               <tbody>`;
            let hasFormulas = false;

            // Iterate over each sheet in the workbook
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                // Iterate over each cell in the sheet
                for (const cellAddress in worksheet) {
                    // Look for cells that have formulas (cell.f property)
                    if (worksheet[cellAddress].f) {
                        hasFormulas = true;
                        const originalFormula = worksheet[cellAddress].f;
                        const convertedFormula = translateFormula(originalFormula, conversionMap);

                        resultsHtml += `<tr>
                                          <td class="table-cell">${sheetName}</td>
                                          <td class="table-cell font-mono">${cellAddress}</td>
                                          <td class="table-cell font-mono">${originalFormula}</td>
                                          <td class="table-cell font-mono">${convertedFormula.html}
                                            ${convertedFormula.isTranslatable ? 
                                                `<span class="ml-2 copy-btn" onclick="copyToClipboard('${CSS.escape(convertedFormula.text)}')">
                                                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16"><path d="M4 1.5H3a2 2 0 0 0-2 2V14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V3.5a2 2 0 0 0-2-2h-1v1h1a1 1 0 0 1 1 1V14a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3.5a1 1 0 0 1 1-1h1v-1z"/><path d="M9.5 1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-3a.5.5 0 0 1-.5-.5v-1a.5.5 0 0 1 .5-.5h3zM-1 8a.5.5 0 0 1 .5-.5h15a.5.5 0 0 1 0 1H-1.5A.5.5 0 0 1-1 8z"/></svg>
                                                </span>` 
                                                : ''
                                            }
                                          </td>
                                        </tr>`;
                    }
                }
            });
            
            resultsHtml += `</tbody></table>`;

            if (!hasFormulas) {
                resultsOutput.innerHTML = `<div class="p-8 text-center text-gray-500">No formulas found in the uploaded file.</div>`;
            } else {
                resultsOutput.innerHTML = resultsHtml;
            }
            
            resultsContainer.classList.remove('hidden');
            // Scroll to results
            resultsContainer.scrollIntoView({ behavior: 'smooth' });
        }

        /**
         * Translates a single formula string based on the provided mapping.
         * This is a simple regex-based approach. A more robust solution would use a proper formula parser (AST).
         * @param {string} formula - The formula to translate.
         * @param {object} map - The conversion dictionary.
         * @returns {{html: string, text: string, isTranslatable: boolean}} - The translated formula object.
         */
        function translateFormula(formula, map) {
            // Regex to find function names (e.g., SUM, VLOOKUP, etc.)
            const functionRegex = /[A-Z][A-Z0-9\.]*\(/ig;
            let isTranslatable = true;

            const translated = formula.replace(functionRegex, (match) => {
                const functionName = match.slice(0, -1).toUpperCase(); // Get name, remove '(', and uppercase
                if (map.hasOwnProperty(functionName)) {
                    const newName = map[functionName];
                    if (newName === null) {
                        isTranslatable = false;
                        return `<span class="bg-red-100 text-red-700 px-2 py-1 rounded">${functionName}</span>(`;
                    }
                    return `<span class="bg-green-100 text-green-700 px-2 py-1 rounded">${newName}</span>(`;
                }
                return match; // Return original if not in map
            });

            // Create a plain text version for the clipboard
            const plainText = translated.replace(/<[^>]*>/g, '');

            return {
                html: translated,
                text: plainText,
                isTranslatable: isTranslatable
            };
        }
        
        /**
         * Copies the provided text to the user's clipboard.
         * @param {string} text - The text to copy.
         */
        function copyToClipboard(text) {
            // A fallback for navigator.clipboard which might not work in all contexts
            const textArea = document.createElement("textarea");
            textArea.value = text;
            document.body.appendChild(textArea);
            textArea.focus();
            textArea.select();
            try {
                document.execCommand('copy');
                showMessageBox();
            } catch (err) {
                console.error('Fallback: Oops, unable to copy', err);
            }
            document.body.removeChild(textArea);
        }

        /**
         * Shows a temporary notification message.
         */
        function showMessageBox() {
            messageBox.classList.remove('opacity-0');
            setTimeout(() => {
                messageBox.classList.add('opacity-0');
            }, 2000);
        }

    </script>
</body>
</html>
