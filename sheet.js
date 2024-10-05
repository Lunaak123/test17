let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Load the first sheet
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell; // Print "NULL" for null values
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply the selected operations and update the table
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and operation columns.');
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());
    filteredData = data.filter(row => {
        const primaryValue = row[primaryColumn];

        if (primaryValue === null && operation === 'null') {
            if (operationType === 'and') {
                return operationColumns.every(col => row[col] === null);
            } else {
                return operationColumns.some(col => row[col] === null);
            }
        } else if (primaryValue !== null && operation === 'not-null') {
            if (operationType === 'and') {
                return operationColumns.every(col => row[col] !== null);
            } else {
                return operationColumns.some(col => row[col] !== null);
            }
        }

        return true;
    });

    displaySheet(filteredData);
}

// Event listeners for Apply button
document.getElementById('apply-operation').addEventListener('click', applyOperation);

// Event listener for Download button to open modal
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

// Event listener to close the modal
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Event listener for Download confirmation
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value || 'download';
    const fileFormat = document.getElementById('file-format').value;

    // Export the filtered data in the chosen format
    if (fileFormat === 'xlsx') {
        const ws = XLSX.utils.json_to_sheet(filteredData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (fileFormat === 'csv') {
        const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        downloadFile(csv, `${filename}.csv`, 'text/csv');
    } else if (fileFormat === 'pdf') {
        // Handle PDF download
        const pdfContent = generatePDFContent(filteredData);
        const pdfBlob = new Blob([pdfContent], { type: 'application/pdf' });
        downloadFile(pdfBlob, `${filename}.pdf`, 'application/pdf');
    } else if (fileFormat === 'jpg' || fileFormat === 'jpeg') {
        // Handle image download (JPG or JPEG)
        const imageBlob = generateImageBlob();
        downloadFile(imageBlob, `${filename}.${fileFormat}`, `image/${fileFormat}`);
    }

    document.getElementById('download-modal').style.display = 'none'; // Close modal after download
});

// Function to download the file
function downloadFile(content, filename, mimeType) {
    const link = document.createElement('a');
    const url = URL.createObjectURL(new Blob([content], { type: mimeType }));
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Load the Excel file initially
loadExcelSheet('path_to_your_excel_file.xlsx');
