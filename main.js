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
