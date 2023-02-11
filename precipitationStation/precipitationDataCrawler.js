import fs from 'fs'
import fetch from 'node-fetch'
import precpStationsData from './params/precpStationsData.js'
import PrecpDataExporter from './exportPrecpStationData.js'

const {stations, queryInfo, startYear, endYear} = precpStationsData

// parse params
const stationsCount = stations.length
const yearsCount = endYear - startYear + 1
const totalReqCount = stationsCount * yearsCount

// POST request
for (let station of stations) {
  let year = startYear
  while (year <= endYear) {
    getPrecpData(station, queryInfo, year)
    // .then(result => fs.writeFileSync('./result', JSON.stringify(result), 'utf8'))
    .then(response => {  // note: 直接拿 year 進到 then 時因為非同步，所以會直接跳成endYear+1
      // console.log(response);
      if (response.metadata.count !== 0) {
        const stationId = response.data[0].StationID
        const yearData = response.data[0].dts
        PrecpDataExporter.buildExcel(yearData, response.year, stationId)
        const accumPrecp = PrecpDataExporter.calAccumPrecpData(yearData, response.year)
        console.log(accumPrecp);
      }
    })
    year += 1
  }
}

// POST req params
async function getPrecpData(station, queryInfo, queryYear) {
  const url = `https://codis.cwb.gov.tw/api/station`

  const params = createFormData()
  return fetch(url, {
    method: 'POST',
    // headers: {'Content-Type': 'x-www-form-urlencoded'},  // 當body為URLSearchParams物件時，該項似乎會自動設定，多設定多錯
    body: params
  })
  .then(res => res.json())  // return json() 的結果
  .then(data => {
    data.year = queryYear
    return data
  })

  function createFormData() {
    const params = new URLSearchParams()
    params.append("type", queryInfo.tableType)
    params.append("stn_ID", station.id)
    params.append("stn_type", station.type)
    params.append("start", `${queryYear}-01-01T00:00:00`)
    params.append("end", `${queryYear}-12-31T00:00:00`)

    return params
  }
}


