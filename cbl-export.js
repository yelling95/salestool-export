const XLSX = require('excel4node')
const request = require('request')
const _ = require('lodash')
const moment = require('moment')
const timezone = require('moment-timezone')
const columns = require('./columns')

const START = 0
const IN = 1
const END = 2

const API_URL = process.env.API_URL

const GET_OPTION = {
  uri: '',
  qs: {}
}

const POST_OPTION = {
  uri: '',
  method: 'POST',
  form: {}
}

const { A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z } = columns()
const _COLUMS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

const DEFAULT_HEIGHT = 17

const START_ROW = 2

const SUMMARY_START_ROW = 3

const SHEET_STYLE = {
  common: {
    percent: {
      default: { numberFormat: '#.##%' },
      integer: { numberFormat: '#%' }
    },
    thuosand: {
      default: { numberFormat: '#,###' }
    },
    decimal: {
      default: { numberFormat: '0.00' },
    }
  },
  bg: {
    white: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'white',
        fgColor: 'white'
      }
    },
    gray: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'F2F2F2',
        fgColor: 'F2F2F2'
      }
    },
    antiquewhite: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'FDE9D9',
        fgColor: 'FDE9D9'
      }
    },
    lightblue: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'DCE6F1',
        fgColor: 'DCE6F1'
      }
    },
    yellow: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'yellow',
        fgColor: 'yellow'
      }
    }
  },
  data: {
    title: {
      alignment: {
        horizontal: ['left'],
        vertical: ['center']
      },
      font: {
        bold: false,
        size: 11,
        name: '맑은 고딕'
      }
    },
    empty: {
      fill: {
        type: 'pattern',
        bgColor: 'white',
        fgColor: 'white'
      },
      border: {
        left: {
          style: 'thin',
          color: 'white'
        },
        right: {
          style: 'thin',
          color: 'white'
        },
        top: {
          style: 'thin',
          color: 'white'
        },
        bottom: {
          style: 'thin',
          color: 'white'
        }
      }
    },
    header: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'D8D8D8',
        fgColor: 'D8D8D8'
      },
      alignment: {
        horizontal: ['center'],
        vertical: ['center']
      },
      font: {
        bold: true,
        size: 9,
        name: '맑은 고딕'
      },
      border: {
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    },
    body: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'white',
        fgColor: 'white'
      },
      alignment: {
        horizontal: ['center'],
        vertical: ['center']
      },
      font: {
        bold: false,
        size: 9,
        name: '맑은 고딕'
      },
      border: {
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    },
    body_top_align: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: 'white',
        fgColor: 'white'
      },
      alignment: {
        horizontal: ['center'],
        vertical: ['top']
      },
      font: {
        bold: false,
        size: 9,
        name: '맑은 고딕'
      },
      border: {
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    }
  }
}

function CBL (app) {
  app.get('/download/cbl', async function(req, res) {
    try {
      console.log('-------------------------------------------')
      console.log('1. Call Download Data ')

      GET_OPTION.uri = API_URL + '/cbl'

      request(GET_OPTION, async function (error, response, body) {
        try {
          if (response.statusCode === 200) {
            const json = JSON.parse(body)
            const data = json.data
            const workbook = new XLSX.Workbook()
            const status = [START, START]
  
            const convert = async () => {
              let summary = workbook.addWorksheet('요약')
              let cbl = workbook.addWorksheet('CBL 입력')
  
              console.log('2. Convert Sheet Data')
              console.log('2-1. Render Data Start')
              await CBL.prototype.renderData(workbook, cbl, data, status, 1)
              console.log('2-1. Render Data End')
              console.log('2-2. Render Summary Start')
              await CBL.prototype.renderSummary(workbook, summary, data, status, 0)
              console.log('2-2. Render Summary End')
            }
  
            await convert()
  
            const finInterval = setInterval(() => {
              console.log('END: ' + END)
              console.log('status 0 : ' + status[0])
              console.log('status 1 : ' + status[1])
              if (status[0] === END && status[1] === END) {
                clearInterval(finInterval)
                console.log('3. Done')
                console.log('')
                workbook.write('CBL분석_' + moment().format('YYYYMMDD') + '.xlsx', res)
              }
            }, 500)
          } else {
            res.json({ status: 'error', message: 'CBL 분석 결과가 존재하지 않습니다.' })
            console.log('[Error] CBL 분석 결과가 존재하지 않습니다.')
          }
        } catch (e) {
          res.json({ status: 'error', message: '[Error] ' + e.message })
          console.log('[Error] ' + e.message)
        }
      })
    } catch (e) {
      res.json({ status: 'error', message: '[Error] ' + e.message })
      console.log('[Error] ' + e.message)
    }
  })

  app.get('/download/cbl/:resId', async function(req, res) {
    try {
      console.log('-------------------------------------------')
      console.log('1. Call Download Data ')
      const resId = req.params.resId
      console.log(resId)

      GET_OPTION.uri = API_URL + '/cbl/' + resId

      request(GET_OPTION, async function (error, response, body) {
        try {
          if (response.statusCode === 200) {
            const json = JSON.parse(body)
            const data = json.data
            const workbook = new XLSX.Workbook()
            const status = [START, START]
  
            const convert = async () => {
              let summary = workbook.addWorksheet('요약')
              let cbl = workbook.addWorksheet('CBL 입력')
  
              console.log('2. Convert Sheet Data')
              console.log('2-1. Render Data Start')
              await CBL.prototype.renderData(workbook, cbl, data, status, 1)
              console.log('2-1. Render Data End')
              console.log('2-2. Render Summary Start')
              await CBL.prototype.renderSummary(workbook, summary, data, status, 0)
              console.log('2-2. Render Summary End')
            }
  
            await convert()
  
            const finInterval = setInterval(() => {
              console.log('END: ' + END)
              console.log('status 0 : ' + status[0])
              console.log('status 1 : ' + status[1])
              if (status[0] === END && status[1] === END) {
                clearInterval(finInterval)
                console.log('3. Done')
                console.log('')
                workbook.write('CBL분석_' + moment().format('YYYYMMDD') + `(${ data[0].resNm })` + '.xlsx', res)
              }
            }, 500)
          } else {
            res.json({ status: 'error', message: 'CBL 분석 결과가 존재하지 않습니다.' })
            console.log('[Error] CBL 분석 결과가 존재하지 않습니다.')
          }
        } catch (e) {
          res.json({ status: 'error', message: '[Error] ' + e.message })
          console.log('[Error] ' + e.message)
        }
      })
    } catch (e) {
      res.json({ status: 'error', message: '[Error] ' + e.message })
      console.log('[Error] ' + e.message)
    }
  })
}

CBL.prototype.renderSummary = async function (wb, ws, data, status, sheetIdx) {
  const style = SHEET_STYLE['data']
  const titleStyle = wb.createStyle(style.title)

  // interval 순서
  let index = 0

  // interval lock 상태
  let isLock = false
  
  // 자원이 시작하는 줄
  let row = SUMMARY_START_ROW

  // cbl 자원이 시작하는줄
  let cblRow = START_ROW

  ws.row(1).setHeight(DEFAULT_HEIGHT)
  ws.row(2).setHeight(DEFAULT_HEIGHT)
  ws.column(C).setWidth(20)
  ws.column(D).setWidth(20)
  ws.cell(2, B).string('[ 감축능력 현황 표 ]').style(titleStyle)
  ws.cell(2, O).string('분석일').style(titleStyle)
  ws.cell(2, P).string(moment().format('YYYY.MM.DD')).style(titleStyle)

  const renderInterval = setInterval(() => {
    if (!isLock) {
      if (index === data.length) {
        clearInterval(renderInterval)
        status[sheetIdx] = END
      } else {
        let res = data[index]
        if (res) {
          isLock = true
          cblRow = cblRow + res.sites.length * 5
          res.index = index

          this.renderSummaryTb(wb, ws, row, cblRow, res)
            .then(() => {
              row = row + 7
              cblRow = cblRow + 2
              isLock = false
              index++
            })
            .catch(e => {
              console.log(e)
            })
        } else {
          index++
        }
      }
    }
  }, 100)
}

CBL.prototype.renderSummaryTb = async function (wb, ws, row, cblRow, resource) {
  const cStyle = SHEET_STYLE['common']
  const bgStyle = SHEET_STYLE['bg']
  const style = SHEET_STYLE['data']
  const headerStyle = wb.createStyle(style.header)
  const bodyStyle = wb.createStyle(style.body)
  const emptyStyle = wb.createStyle(style.empty)
  const thuosandBodyStyle = wb.createStyle(_.assign({}, style.body, cStyle.thuosand.default))
  const thuosandBodyAntiqueStyle = wb.createStyle(_.assign({}, style.body, bgStyle.antiquewhite, cStyle.thuosand.default))
  const thuosandBodyYellowStyle = wb.createStyle(_.assign({}, style.body, bgStyle.yellow, cStyle.thuosand.default))

  // 테이블에 높이 설정
  for (let i = row; i < row + 7; i++) {
    ws.row(i).setHeight(DEFAULT_HEIGHT)
  }
  
  ws.cell(row, B, row + 1, B, true).string('No').style(headerStyle)
  ws.cell(row, C, row + 1, C, true).string('자원명').style(headerStyle)
  ws.cell(row, D, row + 1, D, true).string('구분').style(headerStyle)
  ws.cell(row, E, row, O, true).string('Time').style(headerStyle)
  ws.cell(row + 1, E).string('9:00').style(headerStyle)
  ws.cell(row + 1, F).string('10:00').style(headerStyle)
  ws.cell(row + 1, G).string('11:00').style(headerStyle)
  ws.cell(row + 1, H).string('12:00').style(headerStyle)
  ws.cell(row + 1, I).string('13:00').style(headerStyle)
  ws.cell(row + 1, J).string('14:00').style(headerStyle)
  ws.cell(row + 1, K).string('15:00').style(headerStyle)
  ws.cell(row + 1, L).string('16:00').style(headerStyle)
  ws.cell(row + 1, M).string('17:00').style(headerStyle)
  ws.cell(row + 1, N).string('18:00').style(headerStyle)
  ws.cell(row + 1, O).string('19:00').style(headerStyle)
  ws.cell(row, P, row + 1, P, true).string('비고').style(headerStyle)

  ws.cell(row + 2, B, row + 6, B, true).number(resource.index + 1).style(bodyStyle)
  ws.cell(row + 2, C, row + 6, C, true).string(resource.resNm).style(bodyStyle)
  ws.cell(row + 2, P, row + 6, P, true).string('').style(bodyStyle)
  ws.cell(row + 2, D).string('CBL').style(bodyStyle)
  ws.cell(row + 3, D).string('사용예정량').style(bodyStyle)
  ws.cell(row + 4, D).string('등록용량').style(bodyStyle)
  ws.cell(row + 5, D).string('감축용량').style(bodyStyle)
  ws.cell(row + 6, D).string('차이(감축량-등록용량)').style(bodyStyle)

  const timetable = _.slice(_COLUMS, 4, 15)

  await Promise.all(
    timetable.map(async (_pos, i) => {
      const pos = E + i
      const sumCbl = cblRow
      const sumGap = sumCbl + 1
      ws.cell(row + 2, pos).formula("'CBL 입력'!" + _pos + sumCbl).style(thuosandBodyStyle)
      ws.cell(row + 3, pos).string('-').style(thuosandBodyStyle)
      ws.cell(row + 4, pos).number(resource.kpxCapacity).style(thuosandBodyStyle)
      ws.cell(row + 5, pos).formula(_pos + (row + 2) + '-' + _pos + (row + 3)).style(thuosandBodyAntiqueStyle)
      ws.cell(row + 6, pos).formula("'CBL 입력'!" + _pos + sumGap).style(thuosandBodyYellowStyle)
    })
  )
}

CBL.prototype.renderData = async function (wb, ws, data, status, sheetIdx) {
  const style = SHEET_STYLE['data']
  const headerStyle = wb.createStyle(style.header)
  const bodyStyle = wb.createStyle(style.body)
  const topAlignBodyStyle = wb.createStyle(style.body_top_align)

  // interval 순서
  let index = 0

  // interval lock 상태
  let isLock = false
  
  // 자원이 시작하는 줄
  let row = START_ROW

  // 한 자원 안에 들어가는 줄 수
  let total = 0

  // Header
  ws.row(1).setHeight(DEFAULT_HEIGHT)
  ws.column(A).setWidth(20)
  ws.column(B).setWidth(30)
  ws.column(C).setWidth(15)
  ws.column(D).setWidth(20)
  ws.column(P).setWidth(20)
  ws.cell(1, A).string('참여자원명').style(headerStyle)
  ws.cell(1, B).string('참여고객명').style(headerStyle)
  ws.cell(1, C).string('월별감축가능용량').style(headerStyle)
  ws.cell(1, D).string('구분').style(headerStyle)
  ws.cell(1, E).string('09~10시').style(headerStyle)
  ws.cell(1, F).string('10~11시').style(headerStyle)
  ws.cell(1, G).string('11~12시').style(headerStyle)
  ws.cell(1, H).string('12~13시').style(headerStyle)
  ws.cell(1, I).string('13~14시').style(headerStyle)
  ws.cell(1, J).string('14~15시').style(headerStyle)
  ws.cell(1, K).string('15~16시').style(headerStyle)
  ws.cell(1, L).string('16~17시').style(headerStyle)
  ws.cell(1, M).string('17~18시').style(headerStyle)
  ws.cell(1, N).string('18~19시').style(headerStyle)
  ws.cell(1, O).string('19~20시').style(headerStyle)
  ws.cell(1, P).string('비고').style(headerStyle)

  const renderInterval = setInterval(() => {
    console.log('isLock: ' + isLock)
    if (!isLock) {
      console.log('index: ' + index)
      console.log('data length: ' + data.length)
      if (index === data.length) {
        clearInterval(renderInterval)
        status[sheetIdx] = END
      } else {
        let res = data[index]
        if (res.sites && res.sites.length > 0) {
          isLock = true
          total = res.sites.length * 5 + 2

          // 자원명 표시
          ws.cell(row, A, row + total - 1, A, true).string(res.resNm).style(topAlignBodyStyle)

          // 사업장 테이블 표시
          this.renderSiteTb(wb, ws, row, res.sites)
            .then(() => {
              row = row + res.sites.length * 5 + 2
              isLock = false
              index++
            })
            .catch(e => {
              console.log(e)
            })
        } else {
          index++
        }
      }
    }
  }, 100)
}

CBL.prototype.renderSiteTb = async function (wb, ws, row, sites) {
  const cStyle = SHEET_STYLE['common']
  const bgStyle = SHEET_STYLE['bg']
  const style = SHEET_STYLE['data']
  const bodyStyle = wb.createStyle(style.body)
  const bodyGrayStyle = wb.createStyle(_.assign({}, style.body, bgStyle.gray))
  const percentBodyStyle = wb.createStyle(_.assign({}, style.body, cStyle.percent.integer))
  const thuosandBodyStyle = wb.createStyle(_.assign({}, style.body, cStyle.thuosand.default))
  const thuosandBodyGrayStyle = wb.createStyle(_.assign({}, style.body, bgStyle.gray, cStyle.thuosand.default))
  const decimalBodyStyle = wb.createStyle(_.assign({}, style.body, cStyle.decimal.default))
  const decimalBodyBlueStyle = wb.createStyle(_.assign({}, style.body, bgStyle.lightblue, cStyle.decimal.default))
  const decimalBodyAniqueStyle = wb.createStyle(_.assign({}, style.body, bgStyle.antiquewhite, cStyle.decimal.default))
  

  await Promise.all(
    sites.map(async (site, index) => {
      const start = row + index * 5
      const end = start + 5 - 1

      const _usage = start
      const _cbl = start + 1
      const _capacity = start + 2
      const _reduce = start + 3
      const _gap = start + 4

      // 테이블에 높이 설정
      for (let i = start; i <= end; i++) {
        ws.row(i).setHeight(DEFAULT_HEIGHT)
      }

      ws.cell(start, B, end, B, true).string(site.siteNm).style(bodyStyle)
      ws.cell(start, C, end, C, true).number(site.drmsCapacity).style(thuosandBodyStyle)
      ws.cell(start, P, end, P, true).string('').style(bodyStyle)
      ws.cell(_usage, D).string('사용 예정량').style(bodyStyle)
      ws.cell(_cbl, D).string('CBL').style(bodyStyle)
      ws.cell(_capacity, D).string('월별감축가능용량 대비').style(bodyStyle)
      ws.cell(_reduce, D).string('감축용량').style(bodyStyle)
      ws.cell(_gap, D).string('차이').style(bodyStyle)

      const cbl = site.cbl
      // 2020. 05. 12 reduce는 입력하는 값
      // const reduce = site.reduce

      await Promise.all(
        cbl.map((c, i) => {
          const pos = E + i
          const _pos = _COLUMS[pos - 1]
          // const reduceKw = reduce[i]
          const reduceKw = ''
          ws.cell(_usage, pos).string('').style(bodyStyle)
          
          if (c === null) ws.cell(_cbl, pos).string('-').style(bodyStyle)
          else ws.cell(_cbl, pos).number(c).style(thuosandBodyStyle)

          ws.cell(_capacity, pos).formula((_pos + _cbl) + '-C$' + start).style(thuosandBodyStyle)
          ws.cell(_reduce, pos).string(reduceKw).style(decimalBodyBlueStyle)
          ws.cell(_gap, pos).formula((_pos + _reduce) + '-C$' + start).style(decimalBodyAniqueStyle)
        })
      )
    })
  )

  // 소계
  const start = row + sites.length * 5
  const end = start + 1
  const timetable = _.slice(_COLUMS, 4, 15)

  ws.cell(start, B, end, B, true).string('계').style(bodyGrayStyle)
  ws.cell(start, C, end, C, true).formula('SUM(C' + row + ':C' + (start - 1) + ')').style(thuosandBodyGrayStyle)
  ws.cell(start, D).string('CBL').style(bodyGrayStyle)
  ws.cell(end, D).string('차이').style(bodyGrayStyle)
  ws.cell(start, P).string('').style(bodyGrayStyle)
  ws.cell(end, P).string('').style(bodyGrayStyle)

  await Promise.all(
    timetable.map(async (_pos, i) => {
      const pos = E + i
      let _totalCbl = []
      let _totalGap = []
      await Promise.all(
        sites.map((site, index) => {
          _totalCbl.push(_pos + ((row + index * 5) + 1))
          _totalGap.push(_pos + ((row + index * 5) + 4))
        })
      )
      ws.cell(start, pos).formula('SUM(' + _totalCbl.join(',') + ')').style(bodyGrayStyle)
      ws.cell(end, pos).formula('SUM(' + _totalGap.join(',') + ')').style(bodyGrayStyle)
    })
  )
}

module.exports = CBL
