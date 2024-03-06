function nametoDativeCase(name) {
    name = name.toLowerCase();
    name = name.charAt(0).toUpperCase() + name.slice(1);

    if (name.endsWith('is')) {
        name = name.slice(0, -2) + 'im';
    } else if (!name.endsWith('is') && name.endsWith('s')) {
        name = name.slice(0, -1) + 'am';
    } else if (!name.endsWith('is') && !name.endsWith('s') && name.endsWith('o')) {
    } else if (!name.endsWith('is') && !name.endsWith('s') && name.endsWith('e')) {
        name = name.slice(0, -1) + 'ei';
    } else if (!name.endsWith('is') && !name.endsWith('s') && name.endsWith('a')) {
        name = name.slice(0, -1) + 'ai';
    }

    return name;
}

function surnametoDativeCase(surname, gender) {
    surname = surname.toLowerCase();
    surname = surname.charAt(0).toUpperCase() + surname.slice(1);

    if (surname.endsWith('is')) {
        surname = surname.slice(0, -2) + 'im';
    } else if (surname.endsWith('š')) {
        surname = surname.slice(0, -1) + 'am';
    } else if (surname.endsWith('o')) {
    } else if (surname.endsWith('s')) {
        surname = surname.slice(0, -1) + 'am';
    } else if (surname.endsWith('e')) {
        surname = gender === 'male' ? surname.slice(0, -1) + 'em' : surname.slice(0, -1) + 'ei';
    } else if (surname.endsWith('a')) {
        surname = gender === 'male' ? surname.slice(0, -1) + 'am' : surname.slice(0, -1) + 'ai';
    }

    return surname;
}

function detectGender(name) {
    name = name.toLowerCase();
    if (name.endsWith('is') || name.endsWith('s') || name.endsWith('o')) {
        return 'male';
    } else if (name.endsWith('a') || name.endsWith('e')) {
        return 'female';
    } else {
        return 'other';
    }
}

function checkRequiredFields() {
    var name = document.getElementById('name').value.trim();
    var surname = document.getElementById('surname').value.trim();
    var classValue = document.getElementById('class').value.trim();
    var award = document.getElementById('award').value.trim();
    var subject = document.getElementById('subject').value.trim();
    var olimpnr = document.getElementById('olimpnr').value.trim();

    if (name === '' || surname === '' || classValue === '' || award === '' || subject === '') {
        alert('Please fill in all required fields.');
        return false;
    }

    return true;
}
function addRow() {
    // Retrieve input values
    var name = document.getElementById("name").value.trim();
    var surname = document.getElementById("surname").value.trim();
    var gender = detectGender(name); // Determine gender based on the name
    var selectedClass = document.getElementById("class").value.trim();
    var selectedAward = document.getElementById("award").value.trim();
    var selectedSubject = document.getElementById("subject").value.trim();
    var selectedOlimpNr = document.getElementById("olimpnr").value.trim();
    var selectedAward2 = document.getElementById("award2").value.trim();
    var selectedSubject2 = document.getElementById("subject2").value.trim();
    var selectedOlimpNr2 = document.getElementById("olimpnr2").value.trim();
    var selectedAward3 = document.getElementById("award3").value.trim();
    var selectedSubject3 = document.getElementById("subject3").value.trim();
    var selectedOlimpNr3 = document.getElementById("olimpnr3").value.trim();

    // Check if the name field is empty
    if (name === '') {
        alert('Please enter a name.');
        return; // Exit the function early if the name field is empty
    }

    // Assuming selectedAward, selectedAward2, selectedAward3 are defined elsewhere

    // Check if selectedAward is equal to "atzinība"
    if (selectedAward !== "Atzinība") {
        // Modify selectedAward by replacing the last letter 'a' with 'u'
        var modifiedAward = selectedAward.replace(/a$/, 'u');
    }else{
        var modifiedAward = selectedAward;
    }

    // Check if selectedAward2 is equal to "atzinība"
    if (selectedAward2 !== "Atzinība") {
        // Modify selectedAward2 by replacing the last letter 'a' with 'u'
        var modifiedAward2 = selectedAward2.replace(/a$/, 'u');
    }else{
        var modifiedAward2 = selectedAward2;
    }

    // Check if selectedAward3 is equal to "atzinība"
    if (selectedAward3 !== "Atzinība") {
        // Modify selectedAward3 by replacing the last letter 'a' with 'u'
        var modifiedAward3 = selectedAward3.replace(/a$/, 'u');
    }else{
        var modifiedAward3 = selectedAward3;
    }


    // Append a new row to the table
    var tableRef = document.getElementById('tblexportData').getElementsByTagName('tbody')[0];

    // Insert a new row for the data
    var newRow = tableRef.insertRow(tableRef.rows.length);

    // Insert cells into the new row
    var nameCell = newRow.insertCell(0);
    var surnameCell = newRow.insertCell(1);
    var classCell = newRow.insertCell(2);
    var awardCell = newRow.insertCell(3);
    var subjectCell = newRow.insertCell(4);
    var olimpnrCell = newRow.insertCell(5);
    var award2Cell = newRow.insertCell(6);
    var subject2Cell = newRow.insertCell(7);
    var olimpnr2Cell = newRow.insertCell(8);
    var award3Cell = newRow.insertCell(9);
    var subject3Cell = newRow.insertCell(10);
    var olimpnr3Cell = newRow.insertCell(11);
    var recipientCell = newRow.insertCell(12);
    var dateCell = newRow.insertCell(13);

    // Add input values to the cells
    nameCell.appendChild(document.createTextNode(name));
    surnameCell.appendChild(document.createTextNode(surname));
    classCell.appendChild(document.createTextNode(selectedClass));
    awardCell.appendChild(document.createTextNode(selectedAward));
    subjectCell.appendChild(document.createTextNode(selectedSubject));
    olimpnrCell.appendChild(document.createTextNode(selectedOlimpNr));
    award2Cell.appendChild(document.createTextNode(selectedAward2));
    subject2Cell.appendChild(document.createTextNode(selectedSubject2));
    olimpnr2Cell.appendChild(document.createTextNode(selectedOlimpNr2));
    award3Cell.appendChild(document.createTextNode(selectedAward3));
    subject3Cell.appendChild(document.createTextNode(selectedSubject3));
    olimpnr3Cell.appendChild(document.createTextNode(selectedOlimpNr3));

    // Create hidden input fields for modified values and recipient (visible only in Excel)
    var hiddenAward = document.createElement('input');
    hiddenAward.type = 'hidden';
    hiddenAward.value = modifiedAward;
    awardCell.appendChild(hiddenAward);

    var hiddenAward2 = document.createElement('input');
    hiddenAward2.type = 'hidden';
    hiddenAward2.value = modifiedAward2;
    award2Cell.appendChild(hiddenAward2);

    var hiddenAward3 = document.createElement('input');
    hiddenAward3.type = 'hidden';
    hiddenAward3.value = modifiedAward3;
    award3Cell.appendChild(hiddenAward3);

    // Create hidden input for first subject
    var hiddenSubject = document.createElement('input');
    hiddenSubject.type = 'hidden';
    if (selectedSubject !== '') {
        // If the selected subject is not empty
        if (selectedSubject.toLowerCase() === 'sports') {
            // If the selected subject is "sports", change the last 's' to 'a'
            hiddenSubject.value = selectedSubject.slice(0, -1) + 'ā';
        } else {
            // Add 'S' at the end for all other subjects
            hiddenSubject.value = selectedSubject + 's';
        }
    } else {
        hiddenSubject.value = selectedSubject;
    }
    subjectCell.appendChild(hiddenSubject);

    // Create hidden input for second subject
    var hiddenSubject2 = document.createElement('input');
    hiddenSubject2.type = 'hidden';
    if (selectedSubject2 !== '') {
        // If the selected subject is not empty
        if (selectedSubject2.toLowerCase() === 'sports') {
            // If the selected subject is "sports", change the last 's' to 'a'
            hiddenSubject2.value = selectedSubject2.slice(0, -1) + 'ā';
        } else {
            // Add 'S' at the end for all other subjects
            hiddenSubject2.value = selectedSubject2 + 's';
        }
    } else {
        hiddenSubject2.value = selectedSubject2;
    }
    subject2Cell.appendChild(hiddenSubject2);

    // Create hidden input for third subject
    var hiddenSubject3 = document.createElement('input');
    hiddenSubject3.type = 'hidden';
    if (selectedSubject3 !== '') {
        // If the selected subject is not empty
        if (selectedSubject3.toLowerCase() === 'sports') {
            // If the selected subject is "sports", change the last 's' to 'a'
            hiddenSubject3.value = selectedSubject3.slice(0, -1) + 'ā';
        } else {
            // Add 'S' at the end for all other subjects
            hiddenSubject3.value = selectedSubject3 + 's';
        }
    } else {
        hiddenSubject3.value = selectedSubject3;
    }
    subject3Cell.appendChild(hiddenSubject3);



    // Create hidden input field for recipient
    var hiddenRecipient = document.createElement('input');
    hiddenRecipient.type = 'hidden';
    hiddenRecipient.value = gender === 'male' ? 'skolniekam' : 'skolniecei';
    recipientCell.appendChild(hiddenRecipient);

    // Get today's date
    var today = new Date();

    // Get the day, month, and year
    var day = today.getDate();
    var month = today.getMonth() + 1; // Month is zero-based, so we add 1
    var year = today.getFullYear();

    // Format day and month with leading zeros if necessary
    day = (day < 10) ? '0' + day : day;
    month = (month < 10) ? '0' + month : month;

    var datums = day + '/' + month + '/' + year;

    // Create hidden input field for date
    var hiddenDate = document.createElement('input');
    hiddenDate.type = 'hidden';
    hiddenDate.value = datums;
    dateCell.appendChild(hiddenDate);

    // Clear all input fields after adding a row
    document.getElementById("myForm").reset();
    
    // Add buttons to the newly created row
    addButtonsToRow(newRow);
}


function importFromExcel() {
    var input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx, .xls';
    input.multiple = true; // Allow multiple file selection
    
    input.onchange = function(event) {
        var files = event.target.files;

        // Function to handle file reading
        function readFile(file) {
            return new Promise((resolve, reject) => {
                var reader = new FileReader();
                reader.onload = function(e) {
                    var data = new Uint8Array(e.target.result);
                    var workbook = XLSX.read(data, { type: 'array' });
                    var sheetName = workbook.SheetNames[0];
                    var sheet = workbook.Sheets[sheetName];
                    
                    // Set the range to include only up to column L (0-indexed)
                    var range = XLSX.utils.decode_range(sheet['!ref']);
                    range.e.c = 13; // Column N (index 13, 0-indexed)
                    sheet['!ref'] = XLSX.utils.encode_range(range);

                    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Include headers
                    // Exclude the first row (headers)
                    jsonData.shift();
                    resolve(jsonData);
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        }

        // Array to store promises for each file reading
        var promises = [];
        
        // Read each file and store its promise
        for (var i = 0; i < files.length; i++) {
            promises.push(readFile(files[i]));
        }

        // Resolve all promises
        Promise.all(promises)
            .then(data => {
                // Flatten the array of arrays
                var importedData = data.flat();
                // Append imported data to the table
                appendDataToTable(importedData);
            })
            .catch(error => {
                console.error("Error importing files:", error);
            });
    };

    input.click();
}
function addHeaders(worksheet) {
    // Find the first available column index
    let column = 1;
    while (worksheet.getCell(1, column).value) {
        column++;
    }
    // Set the value of the "Saņēmējs" column header
    worksheet.getCell(1, column).value = 'Saņēmējs';
    // Set the value of the "Datums" column header
    worksheet.getCell(1, column + 1).value = 'Datums';
}

function appendDataToTable(importedData) {
    var tableBody = document.querySelector('#tblexportData tbody');

    importedData.forEach(rowData => {
        var row = document.createElement('tr');

        // Loop through each cell data in the row
        rowData.forEach(cellData => {
            var cell = document.createElement('td');
            cell.textContent = cellData;
            row.appendChild(cell);
        });

        // Add Delete button
        var deleteButton = document.createElement('button');
        deleteButton.textContent = 'Delete';
        deleteButton.className = 'btn btn-danger delete-btn';
        deleteButton.onclick = function() {
            deleteRow(this); // Call deleteRow function passing the button element
        };

        var deleteCell = document.createElement('td');
        deleteCell.appendChild(deleteButton);
        row.appendChild(deleteCell);

        tableBody.appendChild(row);
    });
}
function exportToExcel(tableID, filename) {
    var table = document.getElementById(tableID);
    var wsData = [];

    // Check if the table already contains headers
    var hasHeaders = table.rows[0].cells.length > 0 && table.rows[0].cells[0].innerText.trim() !== '';

    // Create a workbook and worksheet
    var workbook = new ExcelJS.Workbook();
    var worksheet = workbook.addWorksheet('Sheet 1');

    // Iterate through rows and cells
    for (var i = 0; i < table.rows.length; i++) {
        var rowData = [];
        for (var j = 0; j < table.rows[i].cells.length; j++) {
            // Skip column "O"
            if (j === 14) continue;

            var cell = table.rows[i].cells[j];
            var cellData = '';
            // Check if the cell contains a hidden input field for modified award value
            var hiddenInput = cell.querySelector('input[type="hidden"]');
            if (hiddenInput) {
                cellData = hiddenInput.value; // Retrieve the modified award value from the hidden input field
            } else {
                cellData = cell.innerText.trim(); // Get the text content of the cell
            }
            
            // Convert name and surname to dative case if it's not an award field
            if (j === 0) { // Name column
                cellData = nametoDativeCase(cellData);
            } else if (j === 1) { // Surname column
                cellData = surnametoDativeCase(cellData, detectGender(cellData));
            }
            
            rowData.push(cellData);
        }
        wsData.push(rowData);
    }

    // Load data into the worksheet starting from row 2 if headers were added
    if (hasHeaders) {
        for (var i = 0; i < wsData.length; i++) {
            worksheet.addRow(wsData[i]);
        }
    }

    addHeaders(worksheet);

    // Auto-size columns based on content
    worksheet.columns.forEach((column, colIndex) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const textLength = cell.value ? cell.value.toString().length : 0;
            maxLength = Math.max(maxLength, textLength);
        });
        column.width = maxLength < 10 ? 10 : maxLength + 2;
    });

    // Save the workbook
    workbook.xlsx.writeBuffer().then((buffer) => {
        var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Create a temporary URL for the blob
        var url = window.URL.createObjectURL(blob);

        // Create a link element to trigger the download
        var link = document.createElement('a');
        link.href = url;
        link.download = filename || 'TemplateFill.xlsx'; // Default file name
        link.style.display = 'none'; // Hide the link

        // Add the link to the document body
        document.body.appendChild(link);

        // Trigger a click event on the link
        link.click();

        // Clean up
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
    });
}
