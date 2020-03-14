const initialRowNum = 10;
const initialColNum = 5;

// global variables
let sheetData = initSheetData(initialRowNum, initialColNum);
let selectedRow = -1;
let selectedCol = -1;
let rowNum = initialRowNum;
let colNum = initialColNum;

// get dom elements
const btnAddRow = document.querySelector('#btn-add-row');
const btnAddCol = document.querySelector('#btn-add-col');
const btnDelRow = document.querySelector('#btn-del-row');
const btnDelCol = document.querySelector('#btn-del-col');
const btnImport = document.querySelector('#btn-import');
const btnLoad = document.querySelector('#btn-load');
const btnExport = document.querySelector('#btn-export');
const inputUpload = document.querySelector('#input-upload');
const tableBody = document.querySelector('#sheet tbody');
const selection = document.querySelector('#selection');
// create an input element
const editInput = document.createElement('input');
// add a css class to the input element
editInput.classList.add('edit-input');

// function used to make ajax requests, can pull data from backend node server
// this function is async and return a Promise object, so it can be called by 'ajax(...).then(data=>{})' or 'data = await ajax(...)'
function ajax(url, type, data, headers = {'Content-Type': 'application/x-www-form-urlencoded'}) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();

    xhr.open(type, url);
    if (data && headers['Content-Type'] === 'application/x-www-form-urlencoded') {
      data = urlEncode(data);
    } else if (data && headers['Content-Type'] === 'application/json;charset=UTF-8') {
      data = JSON.stringify(data);
    }
    // set request headers if have
    if (headers) {
      Object.keys(headers).map((key) => {
        xhr.setRequestHeader(key, headers[key]);
      });
    }
    xhr.onload = function() {
      // if get correct response code, resolve the response data, otherwise reject it (throw an exception)
      if (xhr.status === 200) {
        resolve(xhr.responseText);
      } else {
        const e = {};
        e.message = xhr.status;
        reject(e);
      }
    };
    // if request method is 'POST' and data is not empty, send the data, otherwise send(request) without data 
    if (type === 'POST' && data) {
      xhr.send(data);
    } else {
      xhr.send();
    }
  });
}

// initialize sheet data with empty strings
// should return a row * col 2-dimension array
function initSheetData(row, col) {
  let sheetData = [];
  const defaultValue = '';
  for (let i = 0; i < row; i++) {
    let rowData = [];
    for (let j = 0; j < col; j++) {
      rowData.push(defaultValue);
    }
    sheetData.push(rowData);
  }
  return sheetData;
} 

// update buttons (add row button, del row button...) status above sheet
function updateButtonStatus() {
  // selectedCol, selectedRow is set to -1 while no row and no col selected. we should set corresponding button disabled
  if (selectedCol > 0) {
    btnAddCol.disabled = false;
    btnDelCol.disabled = false;
  } else {
    btnAddCol.disabled = true;
    btnDelCol.disabled = true;
  }
  if (selectedRow > 0) {
    btnAddRow.disabled = false;
    btnDelRow.disabled = false;
  } else {
    btnAddRow.disabled = true;
    btnDelRow.disabled = true;
  }
  if (colNum <= 1) {
    btnDelCol.disabled = true;
  }
  if (rowNum <= 1) {
    btnDelRow.disabled = true;
  }
}

// draw select rect by it's rectangle position(top, left, right, bottom)
function drawSelectArea(rect) {
  selection.style.top = rect.top + 'px';
  selection.style.left = rect.left + 'px';
  selection.style.height = rect.bottom - rect.top + 'px';
  selection.style.width = rect.right - rect.left + 'px';
  selection.style.display = 'block';
}
// render sheet cells to page
function renderSheet() {
  // reset sheet by removing all cells
  tableBody.innerHTML = '';
  // create a html fragment
  const htmlFrag = document.createDocumentFragment();
  for (let i = 0; i < rowNum + 1; i++) {
    // create table row
    const tr = document.createElement('tr');
    tr.setAttribute('data-row', i);
    for (let j = 0; j < colNum + 1; j++) {
      // create table cell
      const td = document.createElement('td');
      // set first col and first row as sheet header
      if (i === 0 && j === 0) {
        td.classList.add('sheet-header');
      } else if (i === 0 && j !== 0) {
        td.classList.add('sheet-header');
        // column header A-Z
        td.innerText = String.fromCharCode(64 + j);
      } else if (i !== 0 && j === 0) {
        td.classList.add('sheet-header');
        td.innerText = i;
      // render real cells
      } else {
        const value = sheetData[i-1][j-1];
        // if cell value starts with '=', resolve it using computeFormula function
        if (value.startsWith('=')) {
          td.innerText = computeFormula(value.replace(/ /, '').toUpperCase(), [posToCellAddress([i-1, j-1])]);
        // else set the value directly to cell
        } else {
          td.innerText = value;
        }
        
      }
      td.setAttribute('data-id', i + '|' + j);
      td.setAttribute('data-row', i);
      td.setAttribute('data-col', j);
      // append cell to row
      tr.appendChild(td);
    }
    // append row to fragment
    htmlFrag.appendChild(tr);
  }
  // append fragment to document
  tableBody.appendChild(htmlFrag);
  updateButtonStatus();
}

// transform all values in sheet data to plain text.
function getTransformSheet() {
  const transformedSheetData = [];
  for (let i = 0; i < rowNum; i++) {
    const rowData = [];
    for (let j = 0; j < colNum; j++) {
      const value = sheetData[i][j];
      if (value.startsWith('=')) {
        rowData.push(computeFormula(value.replace(/ /, '').toUpperCase()));
      } else {
        rowData.push(value);
      }
    }
    transformedSheetData.push(rowData);
  }
  return transformedSheetData;
}

// export csv data
function exportCSV() {
  // stringify table data
  const transformedSData = getTransformSheet();
  let stringedData = '';
  for (let rowData of transformedSData) {
    stringedData += rowData.join('\t') + '\n';
  }
  stringedData = stringedData.slice(0, -1);
  // download file
  const blob = new Blob(["\uFEFF" + stringedData], { type: 'text/csv;charset=gb2312;' });
  const a = document.createElement('a');
  // set file name
  a.download = 'sheet_' + +new Date;
  a.href = URL.createObjectURL(blob);
  a.click();
}

// render the selection
function renderSelection(row, col) {
  const topHeaderCells = document.querySelectorAll(`td[data-row="0"]`);
  const leftHeaderCells = document.querySelectorAll(`td[data-col="0"]`);

  let rect = {
    left: 0,
    top: 0,
    right: 0,
    bottom: 0
  };
  // if click column header select the entire column
  if (row === 0 && col !== 0) {
    selectedRow = -1;
    selectedCol = col;
    
    for (let cell of topHeaderCells) {
      if (cell.dataset.col !== selectedCol + '') {
        cell.classList.remove('sel-gray');
      } else {
        cell.classList.add('sel-gray');
      }
    }
    for (let cell of leftHeaderCells) {
      cell.classList.remove('sel-gray');
    }
    
    const startCell = document.querySelector(`td[data-id="1|${selectedCol}"]`);
    const endCell = document.querySelector(`td[data-id="${rowNum}|${selectedCol}"]`);
    const { left, top } = startCell.getBoundingClientRect();
    const { right, bottom } = endCell.getBoundingClientRect();
    rect = { left, top, right, bottom };
    selection.classList.add('sel-blue');
  // if click row header select the entire row
  } else if (row !== 0 && col === 0) {
    selectedRow = row;
    selectedCol = -1;

    for (let cell of topHeaderCells) {
      cell.classList.remove('sel-gray');
    }
    for (let cell of leftHeaderCells) {
      if (cell.dataset.row !== selectedRow + '') {
        cell.classList.remove('sel-gray');
      } else {
        cell.classList.add('sel-gray');
      }
    }

    const startCell = document.querySelector(`td[data-id="${selectedRow}|1"]`);
    const endCell = document.querySelector(`td[data-id="${selectedRow}|${colNum}"]`);
    const { left, top } = startCell.getBoundingClientRect();
    const { right, bottom } = endCell.getBoundingClientRect();
    rect = { left, top, right, bottom };
    selection.classList.add('sel-blue');
  // if click the first cell, reset selection
  } else if (row === 0 && col === 0) {
    selectedRow = -1;
    selectedCol = -1;
    for (let cell of topHeaderCells) {
      cell.classList.remove('sel-gray');
    }
    for (let cell of leftHeaderCells) {
      cell.classList.remove('sel-gray');
    }
    console.log('to do');
  // if click the real cell, just select the cell
  } else {
    selectedRow = -1;
    selectedCol = -1;
    for (let cell of topHeaderCells) {
      cell.classList.remove('sel-gray');
    }
    for (let cell of leftHeaderCells) {
      cell.classList.remove('sel-gray');
    }
    const selectCell = document.querySelector(`td[data-id="${row}|${col}"]`);
  
    rect = selectCell.getBoundingClientRect();
    selection.classList.remove('sel-blue');
  }
  // actual draw selection rect
  drawSelectArea(rect);
}
// check if a char is upper-cased alphabet
function isUpperAlphabet(ch) {
  const charCode = ch.charCodeAt();
  if (charCode >= 65 && charCode <= 90) {
    return true;
  } else {
    return false;
  }
}
// check if a string is a number
function isNumber(str) {
  return !isNaN(str);
}
function cellAddressToPos(cellAddress) {
  const col = cellAddress[0].charCodeAt() - 65;
  const row = parseInt(cellAddress.slice(1)) - 1;
  return [row, col];
}
function posToCellAddress(pos) {
  return String.fromCharCode(pos[1] + 65) + (pos[0] + 1);
}
// function to resolve formula
function computeFormula(exp, referrers=[]) {
  console.log(exp, referrers);
  // check if it is a sum expression
  const m = exp.match(/SUM\((.*:.*)\)/);
  if (m) {
    const scope = m[1];
    const [start, end] = scope.split(':');
    // get the start column number, row number and end column number, row number
    const startCol = start[0].charCodeAt() - 65;
    const startRow = parseInt(start.slice(1)) - 1;
    const endCol = end[0].charCodeAt() - 65;
    const endRow = parseInt(end.slice(1)) - 1;
    // avoid circle reference 
    for (let referrer of referrers) {
      const [rRow, rCol] = cellAddressToPos(referrer);
      if (rRow >= startRow && rRow <= endRow && rCol >= startCol && rCol <= endCol) {
        return 'Invalid';
      }
    }
    // sum all the cell values in range
    let sum = 0;
    for (let i = startRow; i <= endRow; i++) {
      for (let j = startCol; j <= endCol; j++) {
        sum += parseInt(sheetData[i][j]);
      }
    }
    return sum;
  } else {
    // get operands
    const operands = exp.slice(1).split(/\+|-|\*|\//);
    let transformedExp = exp.slice(1);
    // avoid circle reference 
    for (let referrer of referrers) {
      if (operands.includes(referrer)) {
        return 'Invalid';
      }
    }
    for (let operand of operands) {
      if (isUpperAlphabet(operand[0])) {
        if(isNumber(operand.slice(1))) {
          const col = operand[0].charCodeAt() - 65;
          const row = parseInt(operand.slice(1)) - 1;
          if (row >= 0 && row < rowNum && col >= 0 && col < colNum) {
            const value = sheetData[row][col];
            // replace cell references with actual values
            if (value.startsWith('=')) {
              referrers.push(operand);
              transformedExp = transformedExp.replace(new RegExp(operand, 'g'), computeFormula(value.replace(/ /, '').toUpperCase(), referrers));
            } else {
              transformedExp = transformedExp.replace(new RegExp(operand, 'g'), value);
            }
            
          }
        }
      }
    }
    let res;
    try {
      res = eval(transformedExp);
    } catch (e) {
      res = 'Invalid';
    }
    return res;
  }
}

renderSheet();

// rxjs event observers
const tbClickStream = Rx.Observable.fromEvent(tableBody, 'click').filter(e => e.target.nodeName === 'TD');
const tbMousedownStream = Rx.Observable.fromEvent(tableBody, 'mousedown');
const tbMousemoveStream = Rx.Observable.fromEvent(tableBody, 'mousemove');
const tbMouseupStream = Rx.Observable.fromEvent(tableBody, 'mouseup');
const endEditStream = Rx.Observable.fromEvent(editInput, 'blur').merge(Rx.Observable.fromEvent(editInput, 'keyup').filter(e => e.keyCode === 13));
const addRowStream = Rx.Observable.fromEvent(btnAddRow, 'click');
const addColStream = Rx.Observable.fromEvent(btnAddCol, 'click');
const delRowStream = Rx.Observable.fromEvent(btnDelRow, 'click');
const delColStream = Rx.Observable.fromEvent(btnDelCol, 'click');
const loadDataStream = Rx.Observable.fromEvent(btnLoad, 'click');
const importStream = Rx.Observable.fromEvent(btnImport, 'click');
const exportStream = Rx.Observable.fromEvent(btnExport, 'click');
const uploadStream = Rx.Observable.fromEvent(inputUpload, 'change');
// click
tbClickStream
  .subscribe(e => {
    const td = e.target;
    const id = td.dataset.id;
    const row = parseInt(td.dataset.row);
    const col = parseInt(td.dataset.col);
    renderSelection(row, col);
    updateButtonStatus();
  });
// double click  
tbClickStream
  .bufferCount(2, 1)
  .filter(([a, b]) => a.target.dataset.id === b.target.dataset.id && !a.target.classList.contains('sheet-header'))
  .map(([e]) => e)
  .subscribe(e => {
    const row =  e.target.dataset.row;
    const col =  e.target.dataset.col;

    // const value = e.target.innerText;
    e.target.innerText = '';
    e.target.appendChild(editInput);
    editInput.value = sheetData[row-1][col-1];
    editInput.focus();
  });
// end edit
endEditStream
  .debounceTime(200)
  .subscribe(e => {
    const td = e.target.parentElement;
    const row = td.dataset.row;
    const col = td.dataset.col;
    sheetData[row-1][col-1] = e.target.value;
    renderSheet();
  });
// add a new column
addColStream
  .subscribe(e => {
    if (colNum >= 26) {
      alert('reach max column number');
      return;
    }
    for (let i = 0; i < rowNum; i++) {
      sheetData[i].splice(selectedCol, 0, '');
    }
    colNum++;
    renderSheet(); 
  });
// delete a column
delColStream
  .subscribe(e => {
    for (let i = 0; i < rowNum; i++) {
      sheetData[i].splice(selectedCol - 1, 1);
    }
    colNum--;
    renderSheet();
  });
// add a new row
addRowStream
  .subscribe(e => {
    const newRow = new Array(colNum);
    newRow.fill('');
    sheetData.splice(selectedRow, 0, newRow);
    rowNum++;
    renderSheet();
  });
// delete a row
delRowStream
  .subscribe(e => {
    sheetData.splice(selectedRow - 1, 1);
    rowNum--;
    renderSheet();
  });
exportStream.subscribe(exportCSV);
// load data from node server
loadDataStream.subscribe(async e => {
  const res = await ajax('/load_data', 'GET');
  const csvData = JSON.parse(res);
  const rlen = csvData.length;
  if (rlen < 1) {
    return;
  }
  const clen = csvData[0].length;
  if (clen < 1) {
    return;
  }
  rowNum = rlen;
  colNum = clen;
  sheetData = csvData;
  renderSheet();
});
// import data
importStream.subscribe(e => {
  inputUpload.click();
});
// import local data using FileReader
uploadStream.subscribe(e => {
  const files = e.target.files;
  const fileReader = new FileReader();
  const csvData = [];
  fileReader.readAsText(files[0], 'UTF-8');
  fileReader.onload = function(e) {
    for (let line of this.result.split('\n')) {
      const row = [];
      for (let v of line.split('\t')) {
        row.push(v);
      }
      csvData.push(row);
    }
    const rlen = csvData.length;
    if (rlen < 1) {
      return;
    }
    const clen = csvData[0].length;
    if (clen < 1) {
      return;
    }
    rowNum = rlen;
    colNum = clen;
    sheetData = csvData;
    renderSheet();
  }
});