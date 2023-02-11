const TITLE_MAPPING = {
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
_extractObjectValue(TITLE_MAPPING)

function _extractObjectValue(obj, filter = obj) {
  let values = []
  for (let key of Object.keys(filter)) {
    let filterValue = JSON.parse(JSON.stringify(filter[key]))
    let objectValue = JSON.parse(JSON.stringify(obj[key]))
    // 若 filterValue 為物件，表示還要再進入取值
    if (filterValue && !Array.isArray(filterValue) && typeof filterValue === 'object') {
      objectValue = _extractObjectValue(objectValue, filterValue)
    }
    if (objectValue === undefined) continue
    values = values.concat(objectValue)
  }
  return values
}