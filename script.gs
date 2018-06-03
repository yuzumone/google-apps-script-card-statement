function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [
    {name: "A", functionName: "A"},
    {name: "B", functionName: "B"},
    {name: "Total", functionName: "total"}
  ];
  ss.addMenu("処理", menu)
}

function A() {
  summary('A')
}

function B() {
  summary('B')
}

function summary(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName(sheetName)
  var values = sheet.getRange(1, 2, sheet.getLastRow(), 2).getValues()
  var attr = {}
  for each(var v in values) {
    name = v[0]
    price = v[1]
    if (attr[name] == null) {
      attr[name] = price
    } else {
      attr[name] += price
    }
  }
  var data = []
  for (key in attr) {
    data.push([key, attr[key]])
  }
  var range = sheet.getRange(1, 5, data.length, 2)
  range.setValues(data)
  range.sort({column: 6, ascending: false})

  // graph
  var chart = sheet.newChart()
  .addRange(range)
  .setChartType(Charts.ChartType.COLUMN)
  .setOption('title', sheetName)
  .setOption('legend.position', 'none')
  .setPosition(1, 7, 10, 10)
  .build()
  sheet.insertChart(chart)

  // upload to slack
  var image = chart.getAs('image/png').setName(sheetName + ".png")
  var title = Utilities.formatDate(new Date(), 'JST', 'yyyyMM') + 'の' + sheetName
  upload(image, title)
}

function total() {
  var names = ['A', 'B']
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var total = ss.getSheetByName('total')
  var data = []
  for each(var name in names) {
    var sheet = ss.getSheetByName(name)
    var lastRow = sheet.getLastRow()
    var values = sheet.getRange(1, 1, lastRow, 3).getValues()
    for each(var v in values) {
      data.push(v)
    }
  }
  var tableRange = total.getRange(1, 1, data.length, 3);
  tableRange.setValues(data)
  tableRange.sort({column: 1, ascending: true})

  var attr = {}
  var values = total.getRange(1, 2, total.getLastRow(), 2).getValues()
  for each(var v in values) {
    name = v[0];
    price = v[1];
    if (attr[name] == null) {
      attr[name] = price
    } else {
      attr[name] += price
    }
  }
  var d = []
  for (key in attr) {
    d.push([key, attr[key]])
  }
  var columnRange = total.getRange(1, 5, d.length, 2)
  columnRange.setValues(d)
  columnRange.sort({column: 6, ascending: false})

  var chart = total.newChart()
  .addRange(columnRange)
  .setChartType(Charts.ChartType.COLUMN)
  .setOption('title', 'Total')
  .setOption('legend.position', 'none')
  .setPosition(1, 7, 10, 10)
  .build()
  total.insertChart(chart);
  var date = Utilities.formatDate(new Date(), 'JST', 'yyyyMM')
  var image = chart.getBlob().getAs('image/png').setName(date + '.png')
  upload(image, date)

  var text = ''
  for each(var v in tableRange.getValues()) {
    var date = Utilities.formatDate(v[0], 'JST', 'yyyy/MM/dd')
    var name = v[1]
    var value = v[2]
    text += date + '\t' + name + '\t' + value + '\n'
  }
  var blob = Utilities.newBlob('', 'text', 'total.txt').setDataFromString(text, 'utf-8')
  upload(blob, 'Total')
}

function upload(file, title) {
 var res = UrlFetchApp.fetch('https://slack.com/api/files.upload', {
    "method" : "post",
    "payload" : {
      token: 'YOUR_SLACK_TOKEN',
      file: file,
      filename: file.getName(),
      channels: 'YOUR_CHANNELS_NAME',
      title: title
    }
  });
}

function copy() {
  var date = new Date()
  var year = date.getYear()
  var month = date.getMonth()
  var templateFile = DriveApp.getFileById('YOUR_TMPLATE_FILE_ID')
  var OutputFolder = DriveApp.getFolderById('YOUR_FOLDER_ID')
  var OutputFileName = Utilities.formatDate(new Date(year, month, 0), 'JST', 'yyyyMM')
  templateFile.makeCopy(OutputFileName, OutputFolder)
  clearAll()
}

function clearAll() {
  var names = ['A', 'B', 'Total']
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  for each(var name in names) {
    var sheet = ss.getSheetByName(name)
    for each(var chart in sheet.getCharts()) {
      sheet.removeChart(chart)
    }
    sheet.clear()
  }
}
