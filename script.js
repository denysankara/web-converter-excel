function mergeSheets() {
  const input = document.getElementById('upload');
  if (!input.files[0]) {
    alert("Pilih file Excel terlebih dahulu!");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type: 'array'});

    let mergedData = [];
    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(sheet);

      mergedData = mergedData.concat(sheetData);
    });

    // Gabungkan berdasarkan 'nomer sep'
    const result = {};
    mergedData.forEach(row => {
      const sep = row['nomer sep'];
      const keterangan = row['keterangan'];

      if (result[sep]) {
        if (!result[sep].includes(keterangan)) {
          result[sep] += `, ${keterangan}`;
        }
      } else {
        result[sep] = keterangan;
      }
    });

    const finalData = Object.keys(result).map(sep => ({
      'nomer sep': sep,
      'keterangan': result[sep]
    }));

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(finalData);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Hasil Gabungan');

    const wbout = XLSX.write(newWorkbook, {bookType: 'xlsx', type: 'array'});
    const blob = new Blob([wbout], {type: 'application/octet-stream'});

    const url = URL.createObjectURL(blob);
    const downloadLink = document.getElementById('download-link');
    downloadLink.href = url;
    downloadLink.download = 'hasil-gabungan.xlsx';
    downloadLink.style.display = 'inline';
    downloadLink.textContent = 'Klik untuk mengunduh file gabungan';
  };

  reader.readAsArrayBuffer(input.files[0]);
}
