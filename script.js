let workbooks = [];

function loadExcel() {
    const excelFiles = document.getElementById('excel-files').files;
    for (let i = 0; i < excelFiles.length; i++) {
        const reader = new FileReader();
        reader.onload = function(event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            workbooks.push(workbook);
            if (workbooks.length === excelFiles.length) {
                displayFileList();
            }
        }
        reader.readAsArrayBuffer(excelFiles[i]);
    }
}

function displayFileList() {
    const excelFiles = document.getElementById('excel-files').files;
    const fileListContainer = document.getElementById('file-list');
    fileListContainer.innerHTML = '';
    workbooks.forEach((workbook, index) => {
        const fileName = excelFiles[index].name.replace(".xlsx","");
        const fileCard = document.createElement('div');
        fileCard.classList.add('file-card');
        const fileNameElement = document.createElement('div');
        fileNameElement.classList.add('file-name');
        fileNameElement.textContent = fileName;
        const downloadButton = document.createElement('button');
        downloadButton.classList.add('file-download');
        downloadButton.textContent = 'Download';
        downloadButton.onclick = function() {
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const manipulatedWorksheet = manipulateWorksheet(worksheet);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, manipulatedWorksheet, sheetName);
            const fileData = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });
            const blob = new Blob([s2ab(fileData)], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${fileName}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);
        }
        fileCard.appendChild(fileNameElement);
        // fileCard.appendChild(viewButton);
        fileCard.appendChild(downloadButton);
        fileListContainer.appendChild(fileCard);
    });
}


function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}

function manipulateWorksheet(worksheet) {
    // Convert worksheet data to array of objects
    let data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    data = data.slice(2)
    // Manipulate the data

        // Step 1: Get unique roll numbers
    const uniqueRollNumbers = [...new Set(data.map((row) => row[0]))].slice(1);

    // Step 2: Create a new array with unique roll numbers and subject headers
    const transformedData = [];
    transformedData.push([...new Set(data.map((row) => row[1]))]);
    transformedData.push([...new Set(data.map((row) => row[2]))]);
    transformedData.push(["S.No","Hall Ticket No",...transformedData[1].slice(1).map((subject) => ["I","E","T","G","GP","C"]).flat()]);

    transformedData[0] = transformedData[0].map((item, index) => {
        if (index === 0) {
          return [item, ""];
        } else {
          return [item, "", "", "","",""];
        }
      }).flat(); 

      transformedData[1] = transformedData[1].map((item, index) => {
        if (index === 0) {
          return [item, ""];
        } else {
          return [item, "", "", "","",""];
        }
      }).flat(); 

    // Step 3: Fill in the internal and external marks for each unique roll number
    for (let i = 0; i < uniqueRollNumbers.length; i++) {
      const rollNumber = uniqueRollNumbers[i];
      const marks = data.filter((row) => row[0] === rollNumber).map((row) => [row[3], row[4],row[5],row[6],row[7],row[8]]);

      transformedData.push([String(i+1),rollNumber, ...marks.flat()]);
    }
  
    console.table(transformedData)

    // Convert the manipulated data back to worksheet format
    // const headers = XLSX.utils.decode_range(worksheet['!ref']).e.c + 1;
    const newWorksheet = XLSX.utils.json_to_sheet(transformedData); // const newWorksheet = XLSX.utils.json_to_sheet(data, { header: [], skipHeader: true });
    // XLSX.utils.sheet_add_aoa(newWorksheet, [[null, null, 'New Column']], { origin: -1 });
    // XLSX.utils.sheet_set_range_style(newWorksheet, XLSX.utils.decode_range(`A${headers + 1}:C${headers + data.length}`), { font: { bold: true } });
    // newWorksheet["!merge"] = [{s:{r:1,c:0},e:{r:1,c:1}},{s:{r:2,c:0},e:{r:2,c:1}}]
    return newWorksheet;
  }
  
