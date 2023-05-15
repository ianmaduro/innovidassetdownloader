document.getElementById('excelFileUpload').addEventListener('change', handleFileUpload);

function handleFileUpload(event) {
    var files = event.target.files;
    var f = files[0];

    var reader = new FileReader();
    reader.onload = function(e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, {type: 'array'});

        var worksheet = workbook.Sheets[workbook.SheetNames[0]];

        var json = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        displayDataInTable(json);
    };
    reader.readAsArrayBuffer(f);
}

function displayDataInTable(data) {
    var table = document.getElementById('excelDataTable');

    while(table.firstChild) {
        table.removeChild(table.firstChild);
    }

    data.forEach(function(rowData, index) {
        if (rowData.length === 0) {
            return;
        }

        var row = document.createElement('tr');

        rowData.forEach(function(cellData, cellIndex) {
            var cell = document.createElement('td');
            cell.appendChild(document.createTextNode(cellData));
            row.appendChild(cell);
        });

        if (index !== 0) {
            var cell = document.createElement('td');
            var checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = 'checkbox-' + index;
            cell.appendChild(checkbox);
            row.appendChild(cell);
        }

        table.appendChild(row);
    });
}

document.getElementById('downloadButton').addEventListener('click', function() {
    var dataRows = document.getElementById('excelDataTable').querySelectorAll('tr:not(:first-child)');
    var downloadCount = 0;
    dataRows.forEach(function(row, index) {
        if (downloadCount >= 6) return;

        var checkbox = row.querySelector('input[type=checkbox]');
        var statusCell = row.querySelector('td:last-child');

        // Skip if the row doesn't have a checkbox or already has a "Complete" cell
        if (!checkbox || (statusCell && statusCell.textContent === 'Complete')) {
            return;
        }

        if (checkbox.checked) {
            downloadCount++;
            var urlCell = row.querySelector('td:nth-child(4)');
            if (urlCell) {
                var url = urlCell.textContent;
                window.open(url, '_blank');
                var cell = document.createElement('td');
                cell.appendChild(document.createTextNode('Complete'));
                row.appendChild(cell);
            }
        }
    });
});

document.getElementById('templateDownloadButton').addEventListener('click', function() {
    var url = 'https://drive.google.com/uc?export=download&id=1MC_m4e1-hLRQ2nkEvH9ZZtxqEM_qLUdN';
    window.open(url, '_blank');
});

document.getElementById('BatchRenamerTool').addEventListener('click', function() {
    var url = 'https://drive.google.com/uc?export=download&id=169tuB-wbIT8gkbs1-8VdO4X0bfmwqqyU';
    window.open(url, '_blank');
});

document.getElementById('excelFileUploadCreative').addEventListener('change', handleFileUploadCreative);

function handleFileUploadCreative(event) {
    var files = event.target.files;
    var f = files[0];

    var reader = new FileReader();
    reader.onload = function(e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, {type: 'array'});

        var worksheet = workbook.Sheets[workbook.SheetNames[0]];

        var json = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        displayDataInTableCreative(json);
    };
    reader.readAsArrayBuffer(f);
}

function displayDataInTableCreative(data) {
    var table = document.getElementById('excelDataTableCreative');

    while(table.firstChild) {
        table.removeChild(table.firstChild);
    }

    data.forEach(function(rowData, rowIndex) {
        if (rowData.length === 0) {
            return;
        }

        var row = document.createElement('tr');

        // Display columns 5 to 10
        rowData.slice(4, 11).forEach(function(cellData, cellIndex) {
            var cell = document.createElement('td');
            cell.appendChild(document.createTextNode(cellData));
            row.appendChild(cell);
        });

        if(rowIndex === 0) {
            // Add a "Status" header in the first row
            var header = document.createElement('th');
            header.appendChild(document.createTextNode('Download Status'));
            row.appendChild(header);
        } else {
            // Add checkbox to rows starting from row 2
            var cell = document.createElement('td');
            var checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = 'checkboxCreative-' + rowIndex;
            cell.appendChild(checkbox);
            row.appendChild(cell);

            // Add a "Status" column
            var statusCell = document.createElement('td');
            statusCell.appendChild(document.createTextNode('')); // initially empty
            statusCell.id = 'statusCreative-' + rowIndex;
            row.appendChild(statusCell);
        }

        table.appendChild(row);
    });
}

function openTab(evt, tabName) {
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }
    tablinks = document.getElementsByClassName("tablinks");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active", "");
    }
    document.getElementById(tabName).style.display = "block";
    evt.currentTarget.className += " active";
}

document.getElementsByClassName("tablinks")[0].click();


document.getElementById('downloadButtonCreative').addEventListener('click', function() {
    var dataRows = document.getElementById('excelDataTableCreative').querySelectorAll('tr:not(:first-child)');
    var urlsToOpen = [];
    var downloadCount = 0;

    dataRows.forEach(function(row, index) {
        if (downloadCount >= 6) return;

        var checkbox = row.querySelector('input[type=checkbox]');
        var statusCell = row.querySelector('td:last-child'); // assuming "Download Status" is the last column

        // Skip if the row doesn't have a checkbox or already has a "Complete" status
        if (!checkbox || (statusCell && statusCell.textContent === 'Complete')) {
            return;
        }

        if (checkbox.checked) {
            var urlCell = row.querySelector('td:nth-child(5)'); // Now it's the 5th column
            if (urlCell) {
                var url = urlCell.textContent;
                // Replace http with https
                url = url.replace('http://', 'https://');
                urlsToOpen.push(url);

                statusCell.textContent = 'Complete'; // set status cell to "Complete"

                downloadCount++;
            }
        }
    });

    if (urlsToOpen.length > 0) {
        var popupWindow = window.open("", "_blank");
        urlsToOpen.forEach(function(url) {
            popupWindow.document.write('<p><a href="' + url + '" target="_blank">' + url + '</a></p>');
        });
    }
});
