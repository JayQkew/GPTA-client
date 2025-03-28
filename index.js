const generateTableBtn = document.querySelector('.generate-table');
const downloadBtn = document.querySelector('.download-table');
const filterBtnContainer = document.querySelector('.filter-btns-container');

const filterBtns = document.querySelectorAll('.filter-btns-container > button');
const onfindApkFilterBtn = document.getElementById('.on-findApk');
const requireUpdateFilterBtn = document.getElementById('.require-update');
const removeFilterBtn = document.getElementById('.remove-fiters');

let allApps;
let fapkApps;

generateTableBtn.addEventListener('click', async (e) => {
    e.preventDefault();

    const dbApps = await getDbApps();
    const findApkApps = await getFindApkApps();
    // loop through dbApps and make a new array that has dbApps with findApk apps
    allApps = dbApps.map(app => {
        const findApk = findApkApps.find(a => a.packageName == app.package_name);
        return {
            name: app.name,
            packageName: app.package_name,
            lastUpdated: app.last_updated ? app.last_updated : 'new release',
            gpVersion: app.version,
            findApkVersion: findApk ? findApk.versionName : null
        }
    });

    fapkApps = findApkFilter();

    generateTable(allApps);
})

downloadBtn.addEventListener('click', exportTableToExcel);

filterBtns.forEach(btn => {
    btn.addEventListener('click', (e) => {
        e.preventDefault();
        filterBtns.forEach(b => b.classList.remove('selected-filter'));
        btn.classList.add('selected-filter');

        if(btn.classList.contains('on-findApk')){
            generateTable(fapkApps);
            console.log('find apk filter pressed')
        } else if (btn.classList.contains('require-update')){
            generateTable(requireUpdateFilter());
            console.log('require update filter pressed')
        } else{
            generateTable(allApps);
            console.log('no filter pressed')
        }
    });
});

//TODO: make request to supabase API
async function getDbApps(){
    try {
        const res = await fetch('/db-apps');
        if(!res.ok) throw new Error('Failed to get db apps');
        const data = await res.json();
        return data.data;
    } catch (error){
        console.error("Error getting database apps: ", error);
    }
}

//TODO: make request directly to findApk
async function getFindApkApps(){
    try{
        const res = await fetch('/findApk-apps');
        if(!res.ok) throw new Error('Failed to get findApk apps');
        const data = await res.json();
        return data.data;
    } catch (error){
        console.error('Error getting findApk apps: ', error);
    }
}

//#region Table Functions
function generateTable(data) {
    const table = document.getElementById('top-apps-table');
    table.innerHTML = '';

    const thead = document.createElement('thead');
    const hRow = document.createElement('tr');

    const cols = Object.keys(data[0]);

    cols.forEach((c, index) => {
        const th = document.createElement('th');
        const formattedc = c.replace(/([A-Z])/g, ' $1').trim();
        th.textContent = formattedc.charAt(0).toUpperCase() + formattedc.slice(1);
        th.setAttribute('data-column', c);
        th.setAttribute('data-order', 'desc');

        if (c === "Updated") {
            th.id = "update-header";
        }

        th.addEventListener('click', () => {
            sortTableByColumn(table, index, th);
        });

        hRow.appendChild(th);
    });

    thead.appendChild(hRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    data.forEach(row => {
        const tr = document.createElement('tr');
        cols.forEach(c => {
            const td = document.createElement('td');

            if (c === "Url") {
                const link = document.createElement('a');
                link.href = row[c]; // Set href attribute
                link.textContent = "Open Link"; // Text inside the link
                link.target = "_blank"; // Open in new tab
                link.rel = "noopener noreferrer"; // Security best practice
                td.appendChild(link);
            } else {
                td.textContent = row[c];
            }

            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    document.querySelector('.download-table').disabled = false;
    document.querySelector('.filter-btns').classList.remove('hidden-v');
}

function exportTableToExcel(){
    const table = document.getElementById('top-apps-table'); // Assuming you're working with an HTML table

    // Convert the data into an array first, if needed
    const data = [];
    const rows = table.rows;
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowData = [];
        for (let j = 0; j < row.cells.length; j++) {
            let cell = row.cells[j].innerText || row.cells[j].textContent;
            if (j === 3) { 
                cell = cell.toString();
            }
            rowData.push(cell);
        }
        data.push(rowData);
    }

    // Now convert the modified data into a sheet
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Create a workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "TopApps");

    XLSX.writeFile(wb, `top_apps_${Date.now()}.xlsx`);
}

function sortTableByColumn(table, columnIndex, header) {
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));

    const order = header.getAttribute('data-order');
    const isAscending = order === 'asc';

    rows.sort((rowA, rowB) => {
        let cellA = rowA.children[columnIndex].textContent.trim();
        let cellB = rowB.children[columnIndex].textContent.trim();

        // Check if the column is a date
        const dateRegex = /\b\w{3}\s\d{2},\s\d{4}\b/; // Matches "Jan 01, 2024"
        if (dateRegex.test(cellA) && dateRegex.test(cellB)) {
            cellA = new Date(cellA);
            cellB = new Date(cellB);
        } else if (!isNaN(cellA) && !isNaN(cellB)) {
            // Convert numbers if the column contains numeric values
            cellA = parseFloat(cellA);
            cellB = parseFloat(cellB);
        }

        return isAscending ? cellA - cellB : cellB - cellA;
    });

    // Toggle order for next click
    header.setAttribute('data-order', isAscending ? 'desc' : 'asc');

    // Re-add sorted rows to tbody
    tbody.innerHTML = '';
    rows.forEach(row => tbody.appendChild(row));
}

function findApkFilter(){
    const filteredApps = allApps.filter(app => {
        if(app.findApkVersion != null) return app;
    });
    return filteredApps;
}

function requireUpdateFilter(){
    const filteredApps = fapkApps.filter(app => {
        if(app.gpVersion != app.findApkVersion) return app;
    });
    return filteredApps;
}

function noFilter(){
    return allApps;
}
//#endregion

