const model = require('./model');
const catchAsync = require(`./utils/catchAsync`);

exports.getCellsValues = catchAsync(async function (req, res, next) {
  const userData = { userName: req.body.login, password: req.body.password },
    worksheet = req.body.worksheet,
    cellRange = req.body.range,
    targetCell = req.body.targetCell,
    token = req.body.token,
    serverURL = model.getSeverURL(req.body.urlToFile),
    fileID = model.getFileID(req.body.urlToFile),
    // token = await model.getToken(serverURL, userData),
    cellData = await model.getCellData(serverURL, fileID, token, worksheet);

  const columnsArr = model.generateColumnsArray(cellRange),
    rowsObj = model.getFirstAndLastRow(cellRange),
    cellArr = model.generateCellRangeArray(columnsArr, rowsObj);

  const targetColumnArr = model.calculateTargetColumns(targetCell, columnsArr),
    targetRowsObj = model.calculateTargetRows(targetCell, rowsObj),
    targetCellArr = model.generateCellRangeArray(
      targetColumnArr,
      targetRowsObj
    );

  const cellsObj = model.generateCellsObj(cellArr, cellData, targetCellArr);

  res.status(200).json({
    status: `success`,
    data: {
      cellsObj,
    },
  });
});
