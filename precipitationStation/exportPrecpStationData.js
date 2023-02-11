import fs from 'fs'
import xlsx from 'node-xlsx'

// const data = JSON.parse(fs.readFileSync('./result'))

function PrecpDataExporter() {
}

// 之後再抽成 class = =，可能新版+esm，語法有些更動所以用 function Class 的寫法一直有問題
PrecpDataExporter.TITLE_MAPPING = {
  DataYearMonth: '觀測時間',
  StationPressure: {
    Mean: '測站平均氣壓',
    Maximum: '測站最高氣壓',
    MaximumTime: '測站最高氣壓時間',
    Minimum: '測站最低氣壓',
    MinimumTime: '測站最低氣壓時間'
  },
  AirTemperature: {
    Mean: '平均氣溫',
    Maximum: '最高氣溫',
    MaximumTime: '最低氣溫時間',
    Minimum: '最低氣溫',
    MinimumTime: '最低氣溫時間'
  },
  WindSpeed: {
    Mean: '風速'
  },
  WindDirection: {
    Prevailing: '風向'
  },
  PeakGust: {
    Maximum: '最大陣風',
    Direction: '最大陣風風向',
    MaximumTime: '最大陣風風速時間'
  },
  Precipitation: {
    Accumulation: '累積降水量',
    PrecipitationDays: '降水日數',
    DailyMaximum: '最大日降水量',
    DailyMaximumDate: '最大降水量時刻表'
  },
  RelativeHumidity: {
    Mean: '平均相對濕度'
  }
}

// 初始化年平均雨量，預期每年資料皆有12筆月雨量
PrecpDataExporter.initYearMeanPrecp = function () {
  return {
    monthCount: 12,  // data 後面處理
    monthPrecp: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  // 值來自 {Precipitation: Accumulation}
  }
}

PrecpDataExporter.createHeaders = function () {
  return _extractObjectValue(this.TITLE_MAPPING)
}

// 單獨處理某個月的資料(取出要寫入Excel的資料)
PrecpDataExporter.createStationDataRow = function (monthData) {
  return _extractObjectValue(monthData, this.TITLE_MAPPING)
}

// 取出當月份的累積雨量
PrecpDataExporter.getMonthAccumPrecp = function (monthData) {
  const filter = {
    Precipitation: {
      Accumulation: '累積降水量'
    }}
  const accumPrecp = _extractObjectValue(monthData, filter)
  return accumPrecp?.[0] || 0
}

PrecpDataExporter.buildExcel = function (yearData, year, stationName) {
  let dataRows = []
  dataRows.push(this.createHeaders())

  for (const monthData of yearData) {
    dataRows.push(this.createStationDataRow(monthData))
  }

  // 製作 Excel/Csv
  const buffer = xlsx.build([{name: `${stationName}`, data: dataRows}])
  fs.writeFileSync(`./${stationName}_${year}.xlsx`, buffer)
}

PrecpDataExporter.calAccumPrecpData = function (yearData, year) {
  let yearMeanPrecp = this.initYearMeanPrecp()

  let monthIndex = 0
  for (const monthData of yearData) {
    yearMeanPrecp.monthPrecp[monthIndex] = this.getMonthAccumPrecp(monthData)
    monthIndex += 1
  }
  yearMeanPrecp.total = yearMeanPrecp.monthPrecp.reduce((accum, currVal) => accum + currVal, 0);
  yearMeanPrecp.year = year
  return yearMeanPrecp
}

// TODO: 之後拆出去 common lib
// 取物件的值(最底層)，可依照樣板key進行過濾
function _extractObjectValue(obj, filter = obj) {
  let values = []
  for (let key of Object.keys(filter)) {
    let filterValue = JSON.parse(JSON.stringify(filter[key]))
    let objectValue = JSON.parse(JSON.stringify(obj[key]))
    // 若 filterValue 為物件，表示還要再進入取值
    if (filterValue && !Array.isArray(filterValue) && typeof filterValue === 'object') {
      objectValue = _extractObjectValue(objectValue, filterValue)
    }

    // 若物件本身沒有該欄位則跳過
    if (objectValue === undefined) continue
    values = values.concat(objectValue)
  }
  return values
}

export default PrecpDataExporter