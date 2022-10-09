const XLSX = require('xlsx');
const http = require(`./utils/httpReq`);

const totalColumnsArr = generateColumnsRange();

function generateColumnsRange(length = 16384) {
  const columns = [];

  const nextChar = (c) => (c ? String.fromCharCode(c.charCodeAt(0) + 1) : 'A');

  const nextCol = (s) =>
    s.replace(/([^Z]?)(Z*)$/, (_, a, z) => nextChar(a) + z.replace(/Z/g, 'A'));

  for (let i = 0, s = ''; i < length; i++) {
    s = nextCol(s);
    columns.push(s);
  }

  return columns;
}

function findColumn(arr, i) {
  const column = arr[i]
    .split(``)
    .filter((el) => el.match(/[a-z]/i))
    .join(``);

  return column;
}

function findRow(arr, i) {
  const row = arr[i]
    .split(``)
    .filter((el) => el.match(/[0-9]/i))
    .join(``);

  return row;
}

exports.calculateTargetColumns = function (cell, columnsArr) {
  const startCell = [cell.toUpperCase()],
    startColumn = findColumn(startCell, 0),
    targetColumnsArr = totalColumnsArr.slice(
      totalColumnsArr.indexOf(startColumn),
      totalColumnsArr.indexOf(startColumn) + columnsArr.length
    );

  return targetColumnsArr;
};

exports.calculateTargetRows = function (cell, rowsObj) {
  const startCell = [cell.toUpperCase()];

  return {
    firstRow: +findRow(startCell, 0),
    lastRow: +findRow(startCell, 0) + rowsObj.lastRow - rowsObj.firstRow,
  };
};

exports.generateColumnsArray = function (cellRange) {
  const rangeData = cellRange.toUpperCase().split(`:`),
    firstColumn = findColumn(rangeData, 0),
    lastColumn = findColumn(rangeData, 1),
    columnArr = totalColumnsArr.slice(
      totalColumnsArr.indexOf(firstColumn),
      totalColumnsArr.indexOf(lastColumn) + 1
    );

  return columnArr;
};

exports.getFirstAndLastRow = function (cellRange) {
  const rangeData = cellRange.toUpperCase().split(`:`);

  return { firstRow: +findRow(rangeData, 0), lastRow: +findRow(rangeData, 1) };
};

exports.generateCellRangeArray = function (columnsArr, rowsObj) {
  const cellArr = [];

  for (let i = rowsObj.firstRow; i <= rowsObj.lastRow; i++) {
    columnsArr.forEach((el) => cellArr.push(el + i));
  }

  return cellArr;
};

exports.getSeverURL = function (url) {
  const serverURL =
    url.slice(0, url.indexOf(`/`) + 2) +
    url.slice(url.indexOf(`/`) + 2).split(`/`)[0];

  return serverURL;
};

exports.getFileID = function (url) {
  const fileID =
    url.indexOf(`&`) !== -1
      ? url.slice(url.indexOf(`=`) + 1, url.indexOf(`&`))
      : url.slice(url.indexOf(`=`) + 1, url.length + 1);

  return fileID;
};

exports.getToken = async function (serverURL, userData) {
  try {
    const respData = await http.sendData(
      `post`,
      `${serverURL}/api/2.0/authentication`,
      { 'Content-Type': 'application/json' },
      userData
    );

    return respData.token;
  } catch (error) {
    console.log(error);
  }
};

exports.getCellData = async function (serverURL, fileID, token, worksheet) {
  try {
    const respData = await http.downloadData(
      `get`,
      `${serverURL}/Products/Files/HttpHandlers/filehandler.ashx?action=download&fileid=${fileID}`,
      {
        Authorization: `Bearer ${token}`,
      },
      'arraybuffer'
    );

    const wb = XLSX.read(respData, {
      cellNF: true,
    });

    const ws = wb.Sheets[worksheet];

    return ws;
  } catch (error) {
    console.log(error);
  }
};

function correctNumberFormat(format) {
  if (format.includes(`* #,##0.00`)) {
    return `_(${format.slice(
      format.indexOf(`$`) + 1,
      format.indexOf(`$`) + 2
    )}* #,##0.00_)`;
  } else if (format.includes(`#,##0.00`)) {
    return `${format.slice(
      format.indexOf(`$`) + 1,
      format.indexOf(`$`) + 2
    )}#,##0.00`;
  } else {
    return format;
  }
}

exports.generateCellsObj = function (cellArr, cellData, targetCellArr) {
  const cellsObj = {};

  cellArr.forEach((el, i) => {
    cellsObj[targetCellArr[i]] = {};
  });
  cellArr
    .map((cell) => (cellData[cell] ? cellData[cell][`v`] : ``))
    .forEach((el, i) => (cellsObj[targetCellArr[i]][`value`] = el));

  cellArr
    .map((cell) => (cellData[cell] ? cellData[cell][`z`] : ``))
    .forEach(
      (el, i) =>
        (cellsObj[targetCellArr[i]][`format`] = correctNumberFormat(el))
    );

  return cellsObj;
};
