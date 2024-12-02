fetch('example.xlsx')
    .then(response => response.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const htmlTable = XLSX.utils.sheet_to_html(firstSheet, { editable: false });
        document.getElementById('excelTable').innerHTML = htmlTable;
    })
    .catch(err => {
        console.error("Error loading Excel file:", err);
        document.getElementById('excelTable').innerHTML = "<p>Failed to load the Excel file.</p>";
    });
