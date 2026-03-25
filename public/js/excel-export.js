// Export calculated quotes to Excel format

import * as XLSX from 'xlsx';

function exportQuotesToExcel(quotes) {
    const worksheet = XLSX.utils.json_to_sheet(quotes);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Quotes');

    // Create a binary string
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

    // Save as Excel file
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'calculated_quotes.xlsx';
    a.click();
    window.URL.revokeObjectURL(url);
}

function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xff;
    }
    return buf;
}

// Example usage
// exportQuotesToExcel([{ id: 1, quote: 'Quote 1' }, { id: 2, quote: 'Quote 2' }]);
