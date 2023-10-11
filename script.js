document.addEventListener('DOMContentLoaded', function() {
    const taskTable = document.getElementById('taskTable');
    taskTable.innerHTML = '';

    const excelFilePath = './tasks.xlsx';

    const xhttp = new XMLHttpRequest();
    xhttp.open('GET', excelFilePath, true);
    xhttp.responseType = 'arraybuffer';

    xhttp.onload = function(e) {
        const arraybuffer = xhttp.response;
        const data = new Uint8Array(arraybuffer);
        const workbook = XLSX.read(data, {
            type: 'array'
        });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1
        });

        const headers = jsonData[0];
        const headerRow = document.createElement('tr');
        headers.forEach(function(headerText) {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });
        taskTable.appendChild(headerRow);

        for (let i = 1; i < jsonData.length; i++) {
            const rowData = jsonData[i];
            const row = document.createElement('tr');
            rowData.forEach(function(cellData) {
                const td = document.createElement('td');
                // Check if the cell data contains 'https:' indicating it may have links
                if (cellData.toString().includes('https:')) {
                    // Split the cell data by comma to get an array of potential links
                    const potentialLinks = cellData.toString().split(','); 
                    potentialLinks.forEach(function(linkText, index) {
                        linkText = linkText.trim(); // Trim any whitespace
                        // Check if this piece really is a link
                        if (linkText.startsWith('https:')) {
                            const link = document.createElement('a');
                            link.href = linkText;
                            link.textContent = linkText;
                            link.target = '_blank';
                            link.className = 'bullet-link'; // Added class for styling
                            td.appendChild(link);
                            
                            // Add a line break after each link except the last one
                            if (index !== potentialLinks.length - 1) {
                                td.appendChild(document.createElement('br'));
                            }
                        } else {
                            td.appendChild(document.createTextNode(linkText));
                        }
                    });
                } else {
                    td.textContent = cellData;
                }
                row.appendChild(td);
            });
            
            taskTable.appendChild(row);
        }

        // Scroll to the last updated task
        taskTable.lastChild.scrollIntoView();
    };

    xhttp.send();
});

function scrollToTop() {
    const container = document.querySelector('.table-container');
    container.scrollTo({
        top: 0,
        behavior: 'smooth'
    });
}

function scrollToBottom() {
    const container = document.querySelector('.table-container');
    container.scrollTo({
        top: container.scrollHeight,
        behavior: 'smooth'
    });
}