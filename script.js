fetch('example.xlsx')
    .then(response => response.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const merges = firstSheet["!merges"] || [];
        const htmlTable = XLSX.utils.sheet_to_html(firstSheet, { editable: false });

        const tableContainer = document.getElementById('excelTable');
        tableContainer.innerHTML = htmlTable;

        // Handle merged cells
        merges.forEach(merge => {
            const table = tableContainer.querySelector("table");
            const startRow = merge.s.r + 1; // Row start (0-indexed to 1-indexed)
            const startCol = merge.s.c + 1; // Column start
            const endRow = merge.e.r + 1;   // Row end
            const endCol = merge.e.c + 1;   // Column end

            const cell = table.rows[startRow - 1].cells[startCol - 1];
            cell.colSpan = endCol - startCol + 1;
            cell.rowSpan = endRow - startRow + 1;

            // Remove extra cells in the merge range
            for (let row = startRow; row <= endRow; row++) {
                for (let col = startCol; col <= endCol; col++) {
                    if (row !== startRow || col !== startCol) {
                        table.rows[row - 1].deleteCell(startCol - 1);
                    }
                }
            }
        });
    })
    .catch(err => {
        console.error("Error loading Excel file:", err);
    });
