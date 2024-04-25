const fs = require('fs');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

// Read the HTML file
const html = fs.readFileSync('index.html', 'utf8');

// Parse HTML and extract tables
const $ = cheerio.load(html);
const tables = $('table');

// Create a new Excel workbook
const workbook = XLSX.utils.book_new();

// Function to parse tables and add them to the workbook
const parseTablesToExcel = (tables, workbook) => {
    tables.each((index, element) => {
        const table = $(element);
        const tableData = [];

        // Iterate through table rows
        table.find('tr').each((rowIndex, row) => {
            const rowData = [];
            // Iterate through row cells
            $(row).find('td').each((cellIndex, cell) => {
                rowData.push($(cell).text().trim());
            });
            tableData.push(rowData);
        });

        // Convert table data to worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(tableData);
        // Add worksheet to workbook with tab name
        XLSX.utils.book_append_sheet(workbook, worksheet, `Table${index + 1}`);
    });
};

// Parse tables and add them to the workbook
parseTablesToExcel(tables, workbook);

// Write workbook to Excel file
XLSX.writeFile(workbook, 'output.xlsx', { bookType: 'xlsx' });
