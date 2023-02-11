// startDate, endDate的年份會替換掉
const startYear = 2000
const endYear = 2022

const stations =[
  // {
  //   name: '關廟',
  //   id: 'C0X170',
  //   type: 'auto_C0'
  // },
  // {
  //   name: '大寮',
  //   id: 'C0V730',
  //   type: 'auto_C0'
  // },
  // {
  //   name: '崎頂',
  //   id: 'C0O960',
  //   type: 'auto_C0'
  // },
  {
    name: '沙崙',
    id: 'C1N001',
    type: 'auto_C1'
  },
]

const queryInfo = {
  tableType: 'table_year',
  startDate: `2022-01-01T00:00:00`,
  endDate: `2021-12-31T00:00:00`
}

export default {stations, queryInfo, startYear, endYear}