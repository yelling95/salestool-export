const express = require('express')
const app = express()
const request = require('request')
const XLSX = require('excel4node')
const _ = require('lodash')
const fs = require('fs-extra')
const moment = require('moment')
const randomstring = require('randomstring')
const readline = require('readline')

const CBL = require('./cbl-export')
const COL = require('./column')
const sample = require('./rrmse-sample')

const { google } = require('googleapis')

const { fromJS } = require('immutable')

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

const SCOPES = [
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/spreadsheets'
]

const CODE = {
  SUCCESS: 200
}

const ROOT_FOLDER = '12r4lOrOS8OeR4WFnZWhWk65ha18nRx7S'

const TOKEN_PATH = 'token.json'

const CHART_SHEET_NAME = '차트 데이터'

const SHEETS_OPTION = [
  { key: 'raw',                     name: '계량데이터(RAW)',      theme: 'data' },
  { key: 'primary',                 name: '제안서',              theme: 'detail' },
  { key: 'secondary',               name: 'RRMSE 산출 계량데이터', theme: 'data' },

  { key: 'days.usage',              name: '사용량차트',           theme: 'chart', to: 'primary', r: 10, c: COL.A, x: 10, y: 24 },
  { key: 'days.cbl.max',            name: 'CBL(MAX)차트',       theme: 'chart', to: 'primary', r: 10, c: COL.I, x: 10, y: 24 },
  { key: 'days.cbl.mid',            name: 'CBL(MID)차트',       theme: 'chart', to: 'primary', r: 10, c: COL.Q, x: 10, y: 24 },

  { key: 'seasonal.usage.winter',   name: '사용량차트_겨울',       theme: 'chart', to: 'primary', r: 14, c: COL.A, x: 10, y: 24 },
  { key: 'seasonal.cbl.max.winter', name: 'CBL(MAX)차트_겨울',   theme: 'chart', to: 'primary', r: 14, c: COL.I, x: 10, y: 24 },
  { key: 'seasonal.cbl.mid.winter', name: 'CBL(MID)차트_겨울',   theme: 'chart', to: 'primary', r: 14, c: COL.Q, x: 10, y: 24 },

  { key: 'seasonal.usage.spring',   name: '사용량차트_봄',        theme: 'chart', to: 'primary', r: 17, c: COL.A, x: 10, y: 24 },
  { key: 'seasonal.cbl.max.spring', name: 'CBL(MAX)차트_봄',    theme: 'chart', to: 'primary', r: 17, c: COL.I, x: 10, y: 24 },
  { key: 'seasonal.cbl.mid.spring', name: 'CBL(MID)차트_봄',    theme: 'chart', to: 'primary', r: 17, c: COL.Q, x: 10, y: 24 },

  { key: 'seasonal.usage.summer',   name: '사용량차트_여름',        theme: 'chart', to: 'primary', r: 20, c: COL.A, x: 10, y: 24 },
  { key: 'seasonal.cbl.max.summer', name: 'CBL(MAX)차트_여름',    theme: 'chart', to: 'primary', r: 20, c: COL.I, x: 10, y: 24 },
  { key: 'seasonal.cbl.mid.summer', name: 'CBL(MID)차트_여름',    theme: 'chart', to: 'primary', r: 20, c: COL.Q, x: 10, y: 24 },

  { key: 'seasonal.usage.autumn',   name: '사용량차트_가을',        theme: 'chart', to: 'primary', r: 23, c: COL.A, x: 10, y: 24 },
  { key: 'seasonal.cbl.max.autumn', name: 'CBL(MAX)차트_가을',    theme: 'chart', to: 'primary', r: 23, c: COL.I, x: 10, y: 24 },
  { key: 'seasonal.cbl.mid.autumn', name: 'CBL(MID)차트_가을',    theme: 'chart', to: 'primary', r: 23, c: COL.Q, x: 10, y: 24 }
]

const SHEET_STYLE = {
  common: {
    percent: {
      default: { numberFormat: '#.##%;' },
      integer: { numberFormat: '#%;' }
    },
    thuosand: {
      default: { numberFormat: '#,##0' }
    }
  },
  data: {
    header: {
      alignment: {
        horizontal: ['center'],
        vertical: ['center']
      },
      font: {
        bold: true,
        size: 11,
        name: '맑은 고딕'
      }
    },
    body: {},
    lately: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        fgColor: '#a3cbfa'
      }
    },
    latelyData: {
      fill: {
        type: 'pattern',
        patternType: 'solid',
        fgColor: '#f29f6f'
      }
    }
  },
  detail: {
    date: {
      alignment: {
        horizontal: ['right'],
        vertical: ['center']
      }
    },
    point: {
      font: {
        size: 10,
        name: '맑은 고딕',
        color: 'red'
      }
    },
    title: {
      alignment: {
        horizontal: ['left'],
        vertical: ['center']
      },
      font: {
        size: 20,
        name: '맑은 고딕'
      }
    },
    header: {
      alignment: {
        wrapText: true,
        horizontal: ['center'],
        vertical: ['center']
      },
      font: {
        bold: false,
        size: 11,
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
    importHeader: {
      alignment: {
        wrapText: true,
        horizontal: ['center'],
        vertical: ['center']
      },
      font: {
        bold: false,
        size: 11,
        name: '맑은 고딕'
      },
      border: {
        left: {
          style: 'medium',
          color: 'black'
        },
        right: {
          style: 'medium',
          color: 'black'
        },
        top: {
          style: 'medium',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    },
    subHeader: {
      alignment: {
        wrapText: true,
        horizontal: 'left',
        vertical: ['center']
      },
      font: {
        bold: true,
        size: 11,
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
      },
      fill: {
        type: 'pattern',
        patternType: 'solid',
        fgColor: '#E7E6E6'
      }
    },
    body: {
      alignment: {
        horizontal: ['center'],
        vertical: ['center']
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
    importBody: {
      alignment: {
        horizontal: ['center'],
        vertical: ['center']
      },
      border: {
        left: {
          style: 'medium',
          color: 'black'
        },
        right: {
          style: 'medium',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'medium',
          color: 'black'
        }
      }
    },
    importMidBody: {
      alignment: {
        horizontal: ['center'],
        vertical: ['center']
      },
      border: {
        left: {
          style: 'medium',
          color: 'black'
        },
        right: {
          style: 'medium',
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
    table: {
      header: {
        alignment: {
          horizontal: 'center',
          vertical: ['center']
        },
        font: {
          bold: true,
          size: 11,
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
        },
        fill: {
          type: 'pattern',
          patternType: 'solid',
          fgColor: '#B8CCE4'
        }
      },
      body: {
        alignment: {
          horizontal: 'center',
          vertical: ['center']
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
}

function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed
  const oAuth2Client = new google.auth.OAuth2 (client_id, client_secret, redirect_uris[0])
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback)
    oAuth2Client.setCredentials(JSON.parse(token))
    callback(oAuth2Client)
  })
}

function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  })
  console.log('Authorize this app by visiting this url:', authUrl)
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  })
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close()
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error retrieving access token', err)
      oAuth2Client.setCredentials(token)
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err)
        console.log('Token stored to', TOKEN_PATH)
      })
      callback(oAuth2Client)
    })
  })
}

function initialization(auth) {
  GoogleDrivers = google.drive({ version: 'v3', auth })
  GoogleSheets = google.sheets({ version: 'v4', auth })
}

function getPercentScope (value) {
  let scope = ''
  let standard = value * 100

  // 0~6% 매우 우수, 6~15% 우수, 15~30% 보통, 30% 이상 나쁨
  if (standard >= 30) {
    scope = '나쁨'
  } else if (standard >= 15) {
    scope = '보통'
  } else if (standard >= 6) {
    scope = '우수'
  } else {
    scope = '매우 우수'
  }

  return scope
}

async function exportSheetFile(fileId) {
  try {
    const result = await GoogleDrivers.files.export({
      fileId,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }, {
      responseType: 'stream'
    })

    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Export Excel [ ' + fileId + ' ]')
      return null
    } else {
      console.log('[Success] Exported Excel [ ' + fileId + ' ]')
      return result.data
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function createSheetFile(uploadDir, serial, extension) {
  try {
    const result = await GoogleDrivers.files.create({
      resource: {
        name: serial,
        mimeType: 'application/vnd.google-apps.spreadsheet',
        parents: [ROOT_FOLDER]
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: fs.createReadStream(uploadDir + serial + extension)
      }
    })
  
    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Created Excel File  [ ' + serial + ' ]')
      return null
    } else {
      console.log('[Success] Created Excel File [ ' + result.data.id + ' ]')
      const excelDetail = await GoogleSheets.spreadsheets.get({
        spreadsheetId: result.data.id
      })
      if (excelDetail.status === CODE.SUCCESS) {
        return {
          id: result.data.id,
          sheets: excelDetail.data.sheets
        }
      } else {
        return null
      }
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function getChartSeries(sheetId, scope) {
  const series = []
  await Promise.all(
    _.range(scope.row.start, scope.row.end, 1).map(row => {
      series.push({
        "series": {
          "sourceRange": {
            "sources": [
              {
                "sheetId": sheetId,
                "startRowIndex": row - 1,
                "endRowIndex": row,
                "startColumnIndex": scope.column.start - 1,
                "endColumnIndex": scope.column.end + 1
              }
            ]
          }
        },
        "targetAxis": "LEFT_AXIS",
        "dataLabel": {
          "type": "NONE",
          "textFormat": {
             "fontFamily": "Roboto"
          }
        }
      })
    })
  )
  return series
}

async function createSheetsChart(sheetData, chartScope) {
  const fileId = sheetData.id
  const sheetArray = sheetData.sheets
  const targetArray = []
  const toArray = []
  const chartArray = []

  try {
    await Promise.all(
      SHEETS_OPTION.map((option, optionIndex) => {
        if (option.theme === 'chart') {
          const toIndex = SHEETS_OPTION.findIndex(otp => otp.key === option.to)
          toArray.push(sheetArray[toIndex].properties)
          targetArray.push(_.assign(sheetArray[optionIndex].properties, {
            chartKey: option.key,
            chartTitle: option.name,
            chartRow: option.r,
            chartColumn: option.c,
            chartX: option.x,
            chartY: option.y
          }))
        }
      })
    )

    await Promise.all(
      targetArray.map(async (target, index) => {
        const to = toArray[index]
        const scope = chartScope.get(target.chartKey)
        const series = await getChartSeries(target.sheetId, scope)

        chartArray.push({
          "addChart": {
            "chart": {
              "spec": {
                "title": target.chartTitle,
                "basicChart": {
                  "chartType": "LINE",
                  "legendPosition": "NO_LEGEND",
                  "axis": [
                    {
                      "position": "BOTTOM_AXIS"
                    },
                    {
                      "position": "LEFT_AXIS"
                    }
                  ],
                  "domains": [{
                    "domain": {
                      "sourceRange": {
                        "sources": [{
                          "sheetId": target.sheetId,
                          "startRowIndex": 0,
                          "endRowIndex": 1,
                          "startColumnIndex": 0,
                          "endColumnIndex": 25
                        }]
                      }
                    }
                  }],
                  "series": series,
                  "headerCount": 1
                },
                "titleTextFormat": {
                  "fontFamily": "Roboto"
                },
                "fontName": "Roboto"
              },
              "position": {
                "overlayPosition": {
                  "anchorCell": {
                    "sheetId": to.sheetId,
                    "rowIndex": target.chartRow - 1,
                    "columnIndex": target.chartColumn - 1
                  },
                  "offsetXPixels": target.chartX,
                  "offsetYPixels": target.chartY,
                  "widthPixels": 500,
                  "heightPixels": 400
                }
              },
              "border":{
                "color":{
                   
                },
                "colorStyle":{
                   
                }
              }
            }
          }
        })
      })
    )

    const result = await GoogleSheets.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      resource: {
        requests: chartArray
      }
    })

    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Update Google Sheets [ ' + fileId + ' ]')
      return false
    } else {
      console.log('[Success] Updated Google Sheets [ ' + fileId + ' ]')
      return true
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return false
  }
}

async function updateSheetProperties(sheetData) {
  const fileId = sheetData.id
  const sheetArray = sheetData.sheets
  const propertyArray = []
  const requestArray = []

  try {
    await Promise.all(
      SHEETS_OPTION.map((option, optionIndex) => {
        if (option.theme === 'chart') {
          const filterd = fromJS(sheetArray[optionIndex].properties)
            .delete('chartKey')
            .delete('chartTitle')
            .delete('chartRow')
            .delete('chartColumn')
            .delete('chartX')
            .delete('chartY')
            .toJS()
          propertyArray.push(_.assign(filterd, {
            "hidden": true
          }))
        }
      })
    )

    await Promise.all(
      propertyArray.map(async property => {
        requestArray.push({
          "updateSheetProperties": {
            "properties": property,
            "fields": "hidden"
          }
        })
      })
    )

    const result = await GoogleSheets.spreadsheets.batchUpdate({
      spreadsheetId: fileId,
      resource: {
        requests: requestArray
      }
    })

    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Update Google Sheets [ ' + fileId + ' ]')
      return false
    } else {
      console.log('[Success] Updated Google Sheets [ ' + fileId + ' ]')
      return true
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return false
  }
}

async function serialize (data) {
  let beforeTimestamp = null
  let timestamp = null
  let temp = null
  let yyyymmdd = null
  let hh = null
  let hour = null
  let serialized = []
  let scope = { start: 0, end: 0 }

  if (data && data.length > 0) {  
    data.map((d, dIndex) => {
      yyyymmdd = String(d.timestamp).substring(0, 10)
      hh = String(d.timestamp).substring(11, 13)
      
      if (hh === '00') {
        timestamp = moment(yyyymmdd, 'YYYYMMDD').subtract(1, 'days')
        hour = 24
      } else {
        timestamp = moment(yyyymmdd, 'YYYYMMDD')
        hour = Number(hh)
      }

      if (beforeTimestamp === null || beforeTimestamp.format('YYYYMMDD') !== timestamp.format('YYYYMMDD')) {
        if (beforeTimestamp !== null) {
          serialized.push(temp)
        }
        temp = {}
        temp.timestamp = timestamp.format('YYYYMMDD')
        temp.values = []
        temp.values.push(d.value)
        
        beforeTimestamp = timestamp
        scope.start = hour
      } else {
        temp.values.push(d.value)
        scope.end = hour

        if (dIndex === (data.length - 1)) {
          serialized.push(temp)
        }
      }
    })
  }
  return { serialized, scope }
}

async function renderRaw (wb, ws, raw, style) {
  try {
    serialize(raw.data).then(result => {
      const { serialized, scope } = result
      const haederStyle = wb.createStyle(style.header)
      const bodyStyle = wb.createStyle(style.body)
      const startRow = 2
      const startColumn = 2
      
      ws.cell(1, 1).string('날짜').style(haederStyle)
      
      let colIndex = startColumn
      for (let start = scope.start; start <= scope.end; start++) {
        ws.cell(1, colIndex).string((start >= 10 ? start : '0' + start) + ':00').style(haederStyle)
        colIndex++
      }
  
      if (serialized && serialized.length > 0) {
        serialized.map((row, rowIndex) => {
          ws.cell(rowIndex + startRow, 1).string(row.timestamp).style(bodyStyle)
          row.values.map((value, valIndex) => {
            ws.cell(rowIndex + startRow, valIndex + startColumn).number(value ? value : 0).style(bodyStyle)
          })
        })
      }
    })
  } catch (e) {
    console.log(e)
    console.log('[ Error ] ' + e.message)
  }
}

async function renderSummary (wb, ws, data, style) {
  try {
    const defaultHeight = 40
    const chartHeight = 400
    const titleStyle = wb.createStyle(style.title)
    const headerStyle = wb.createStyle(style.header)
    const subHeaderStyle = wb.createStyle(style.subHeader)
    const importHeaderStyle = wb.createStyle(style.importHeader)
    const bodyStyle = wb.createStyle(style.body)
    const percentBodyStyle = wb.createStyle(Object.assign(style.body, SHEET_STYLE['common'].percent.integer))
    const thuosandBodyStyle = wb.createStyle(Object.assign(style.body, SHEET_STYLE['common'].thuosand.default))
    const importBodyStyle = wb.createStyle(style.importBody)
    const thousandImportBodyStyle = wb.createStyle(Object.assign(style.importBody, SHEET_STYLE['common'].thuosand.default))
    const thousandImportMidBodyStyle = wb.createStyle(Object.assign(style.importMidBody, SHEET_STYLE['common'].thuosand.default))
    const pointStyle = wb.createStyle(style.point)
    const dateStyle = wb.createStyle(style.date)
    const tableHeaderStyle = wb.createStyle(style.table.header)
    const tableBodyStyle = wb.createStyle(style.table.body)

    ws.row(1).setHeight(defaultHeight)
    ws.row(4).setHeight(defaultHeight)
    ws.row(5).setHeight(defaultHeight)
    ws.row(6).setHeight(defaultHeight)
    ws.row(9).setHeight(defaultHeight)
    ws.row(11).setHeight(defaultHeight)
    
    ws.column(COL.D).setWidth(15)
    ws.column(COL.E).setWidth(15)
    ws.column(COL.L).setWidth(20)
    ws.column(COL.M).setWidth(20)
    ws.column(COL.N).setWidth(20)

    ws.cell(1, COL.A, 1, COL.K, true).string('■ RRMSE 정보').style(titleStyle)
    ws.cell(3, COL.N).string('검증기간 : ' + data.meta.startDate + ' ~ ' + data.meta.endDate).style(dateStyle)
    ws.cell(4, COL.A, 5, COL.C, true).string('고객명').style(headerStyle)
    ws.cell(4, COL.D, 5, COL.D, true).string('고객번호').style(headerStyle)
    ws.cell(4, COL.E, 5, COL.E, true).string('기준일').style(headerStyle)
    ws.cell(4, COL.F, 4, COL.H, true).string('평일 사용량 중').style(headerStyle)
    ws.cell(4, COL.I, 4, COL.K, true).string('RRMSE').style(headerStyle)
    ws.cell(4, COL.L, 5, COL.L, true).string('계약전력').style(headerStyle)
    ws.cell(4, COL.M, 5, COL.M, true).string('요금적용전력').style(headerStyle)
    ws.cell(4, COL.N, 5, COL.N, true).string('적용요금제').style(headerStyle)

    ws.cell(5, COL.F).string('MAX').style(headerStyle)
    ws.cell(5, COL.G).string('MID').style(headerStyle)
    ws.cell(5, COL.H).string('MIN').style(headerStyle)
    
    ws.cell(5, COL.I).string('MAX(4/5)').style(headerStyle)
    ws.cell(5, COL.J).string('MID(6/10)').style(headerStyle)
    ws.cell(5, COL.K).string('결 과').style(importHeaderStyle)

    ws.cell(6, COL.A, 6, COL.C, true).string(data.meta.siteNm).style(bodyStyle)
    ws.cell(6, COL.D).string(data.meta.kepcoCustomerId).style(bodyStyle)
    ws.cell(6, COL.E).string(data.meta.evaluateDate).style(bodyStyle)
    ws.cell(6, COL.F).number(data.usage.max).style(thuosandBodyStyle)
    ws.cell(6, COL.G).number(data.usage.mid).style(thuosandBodyStyle)
    ws.cell(6, COL.H).number(data.usage.min).style(thuosandBodyStyle)
    ws.cell(6, COL.I).number(data.rrmse.max).style(percentBodyStyle)
    ws.cell(6, COL.J).number(data.rrmse.mid).style(percentBodyStyle)
    ws.cell(6, COL.K).string(getPercentScope(data.rrmse.max > data.rrmse.mid ? data.rrmse.mid : data.rrmse.max)).style(importBodyStyle)
    ws.cell(6, COL.L).number(data.meta.contractElec).style(thuosandBodyStyle)
    ws.cell(6, COL.M).string(data.meta.supplyKw !== undefined ? data.meta.supplyKw : '-').style(bodyStyle)
    ws.cell(6, COL.N).string(data.meta.contractClass + ' ' + data.meta.supplyMethod).style(bodyStyle)

    ws.cell(9, COL.A, 9, COL.X, true).string('■ 사용량/CBL 차트').style(titleStyle)
    ws.cell(10, COL.A, 10, COL.X, true).string('')
    ws.row(10).setHeight(chartHeight)
  
    ws.cell(11, COL.A, 11, COL.X, true).string('■ 계절별 부하패턴 분석').style(titleStyle)
    ws.cell(12, COL.A, 12, COL.X, true).string('12~2월(겨울)').style(subHeaderStyle)
    ws.cell(13, COL.A, 13, COL.H, true).string('사용량').style(subHeaderStyle)
    ws.cell(13, COL.I, 13, COL.P, true).string('CBL(MAX)').style(subHeaderStyle)
    ws.cell(13, COL.Q, 13, COL.X, true).string('CBL(MID)').style(subHeaderStyle)
    ws.cell(14, COL.A, 14, COL.H, true).string('').style(tableBodyStyle)
    ws.cell(14, COL.I, 14, COL.P, true).string('').style(tableBodyStyle)
    ws.cell(14, COL.Q, 14, COL.X, true).string('').style(tableBodyStyle)
    ws.row(14).setHeight(chartHeight)

    ws.cell(15, COL.A, 15, COL.X, true).string('3~5월(봄)').style(subHeaderStyle)
    ws.cell(16, COL.A, 16, COL.H, true).string('사용량').style(subHeaderStyle)
    ws.cell(16, COL.I, 16, COL.P, true).string('CBL(MAX)').style(subHeaderStyle)
    ws.cell(16, COL.Q, 16, COL.X, true).string('CBL(MID)').style(subHeaderStyle)
    ws.cell(17, COL.A, 17, COL.H, true).string('').style(tableBodyStyle)
    ws.cell(17, COL.I, 17, COL.P, true).string('').style(tableBodyStyle)
    ws.cell(17, COL.Q, 17, COL.X, true).string('').style(tableBodyStyle)
    ws.row(17).setHeight(chartHeight)

    ws.cell(18, COL.A, 18, COL.X, true).string('6~8월(여름)').style(subHeaderStyle)
    ws.cell(19, COL.A, 19, COL.H, true).string('사용량').style(subHeaderStyle)
    ws.cell(19, COL.I, 19, COL.P, true).string('CBL(MAX)').style(subHeaderStyle)
    ws.cell(19, COL.Q, 19, COL.X, true).string('CBL(MID)').style(subHeaderStyle)
    ws.cell(20, COL.A, 20, COL.H, true).string('').style(tableBodyStyle)
    ws.cell(20, COL.I, 20, COL.P, true).string('').style(tableBodyStyle)
    ws.cell(20, COL.Q, 20, COL.X, true).string('').style(tableBodyStyle)
    ws.row(20).setHeight(chartHeight)

    ws.cell(21, COL.A, 21, COL.X, true).string('9~11월(가을)').style(subHeaderStyle)
    ws.cell(22, COL.A, 22, COL.H, true).string('사용량').style(subHeaderStyle)
    ws.cell(22, COL.I, 22, COL.P, true).string('CBL(MAX)').style(subHeaderStyle)
    ws.cell(22, COL.Q, 22, COL.X, true).string('CBL(MID)').style(subHeaderStyle)
    ws.cell(23, COL.A, 23, COL.H, true).string('').style(tableBodyStyle)
    ws.cell(23, COL.I, 23, COL.P, true).string('').style(tableBodyStyle)
    ws.cell(23, COL.Q, 23, COL.X, true).string('').style(tableBodyStyle)
    ws.row(23).setHeight(chartHeight)

    const seasonalAvgUseageRowStart = 25
    const seasonalAvgUseageColumnStart = COL.A
    const seasonalAvgUseage = data.data.seasonal.avg

    ws.cell(seasonalAvgUseageRowStart, seasonalAvgUseageColumnStart).string('평균 사용량').style(tableHeaderStyle)
    ws.cell(seasonalAvgUseageRowStart + 1, seasonalAvgUseageColumnStart).string('12~2월').style(tableBodyStyle)
    ws.cell(seasonalAvgUseageRowStart + 2, seasonalAvgUseageColumnStart).string('3~5월').style(tableBodyStyle)
    ws.cell(seasonalAvgUseageRowStart + 3, seasonalAvgUseageColumnStart).string('6~8월').style(tableBodyStyle)
    ws.cell(seasonalAvgUseageRowStart + 4, seasonalAvgUseageColumnStart).string('9~11월').style(tableBodyStyle)

    _.range(1, 25).map((hour, hourIndex) => {
      ws.cell(seasonalAvgUseageRowStart, seasonalAvgUseageColumnStart + hour).string((hour < 10 ? '0' + hour : '' + hour) + '시').style(tableHeaderStyle)
    })

    // winter
    _.range(1, 25).map((hour, hourIndex) => {
      ws.cell(seasonalAvgUseageRowStart + 1, seasonalAvgUseageColumnStart + hour).string(_.floor(seasonalAvgUseage.winter[hourIndex].value, 1) + '').style(tableBodyStyle)
    })

    // spring
    _.range(1, 25).map((hour, hourIndex) => {
      ws.cell(seasonalAvgUseageRowStart + 2, seasonalAvgUseageColumnStart + hour).string(_.floor(seasonalAvgUseage.spring[hourIndex].value, 1) + '').style(tableBodyStyle)
    })

    // summer
    _.range(1, 25).map((hour, hourIndex) => {
      ws.cell(seasonalAvgUseageRowStart + 3, seasonalAvgUseageColumnStart + hour).string(_.floor(seasonalAvgUseage.summer[hourIndex].value, 1) + '').style(tableBodyStyle)
    })

    // autumn
    _.range(1, 25).map((hour, hourIndex) => {
      ws.cell(seasonalAvgUseageRowStart + 4, seasonalAvgUseageColumnStart + hour).string(_.floor(seasonalAvgUseage.autumn[hourIndex].value, 1) + '').style(tableBodyStyle)
    })
  } catch (e) {
    console.log(e)
    console.log('[ Error ] ' + e.message)
  }
}

async function renderChart (wb, ws, key, data, info) {
  try {
    const dataMap = fromJS(data)
    const keys = key.split('.')
    const array = dataMap.getIn(keys).toJS()

    serialize(array).then(result => {
      const { serialized, scope } = result
      const startRow = 2
      const startColumn = 2
      
      ws.cell(1, 1).string('날짜')
      
      let colIndex = startColumn
      for (let start = scope.start; start <= scope.end; start++) {
        ws.cell(1, colIndex).number(start)
        colIndex++
      }
  
      if (serialized && serialized.length > 0) {
        serialized.map((row, rowIndex) => {
          ws.cell(rowIndex + startRow, 1).string(row.timestamp)
          row.values.map((value, valIndex) => {
            ws.cell(rowIndex + startRow, valIndex + startColumn).number(value ? value : 0)
          })
        })
      }

      info.set(key, {
        column: scope,
        row: {
          start: startRow,
          end: serialized.length
        }
      })
    })
  } catch (e) {
    console.log('[ Error ] ' + e.message)
  }
}

app.use(function(req, res, next) {
  res.header('Access-Control-Allow-Origin', '*')
  res.header('Access-Control-Allow-Methods', 'GET, PUT, POST, DELETE, OPTIONS')
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, Content-Length, X-Requested-With')

  if ('OPTIONS' === req.method) {
    res.send(200)
  } else {
    next()
  }
})

/*
app.get('/vaildation', async function(req, res) {
  const raw = sample.raw_report

  const result = await serialize(raw.data)

  res.json(result)
})

app.get('/test', async function (req, res) {
  try {
    const result = await GoogleSheets.spreadsheets.get(
      {
        spreadsheetId: '1iMgx21CMxPQ4_zwrwMDlfBVUYKaLCUyiMdwNAKBcia8'
      }
    )
  
    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot GET Excel File')
      return null
    } else {
      console.log('[Success] GET Excel File [ ' + result.data.id + ' ]')
      res.json(result)
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
})
*/

app.get('/download', function(req, res) {
  let idx = req.query.idx

  if (idx) {
    console.log('-------------------------------------------')
    console.log('1. Created a Templte Workbook')
    console.log('2. Call Download Data [idx:' + idx + ']')

    GET_OPTION.uri = API_URL + '/history/' + idx

    request(GET_OPTION, async function (error, response, body) {
      try {
        if (response.statusCode === 200) {
          console.log('3. Received Download Data')
          const json = JSON.parse(body)
          const raw = JSON.parse(json.raw_report)
          const primary = JSON.parse(json.primary_report)
          const secondary = JSON.parse(json.secondary_report)

          const data = {
            raw,
            primary,
            secondary
          }

          const workbook = new XLSX.Workbook()
          const worksheets = []
          const chartScope = new Map()

          const convert = async () => {
            SHEETS_OPTION.forEach(async sheetOpt => {
              const worksheet = workbook.addWorksheet(sheetOpt.name)
              const style = SHEET_STYLE[sheetOpt.theme]

              if (sheetOpt.theme === 'data') {
                await renderRaw(workbook, worksheet, data[sheetOpt.key], style)
              } else if (sheetOpt.theme === 'detail') {
                await renderSummary(workbook, worksheet, data[sheetOpt.key], style)
              } else {
                await renderChart(workbook, worksheet, sheetOpt.key, primary.data, chartScope)
              }

              console.log('4. Convent Sheet Data [' + sheetOpt.name + ']')
              worksheets.push(worksheet)
            })
          }

          await convert()

          console.log('5. Done')
          console.log('')

          const uploadDir = './temp/rrmse/'
          const extension = '.xlsx'
          const serial = randomstring.generate(24)
          const fileName = raw.meta.siteNm + '_RRMSE산출이력_' + moment().format('YYYYMMDD')

          workbook.write(uploadDir + serial + extension, async (error, status) => {
            if (error) {
              res.json({ status: 'error', message: 'RRMSE 분석 결과를 엑셀로 변환하지 못했습니다.' })
              console.log('[Error] RRMSE 분석 결과를 엑셀로 변환하지 못했습니다.')
            } else {
              console.log('6. Excel File Upload')
              let sheetData = await createSheetFile(uploadDir, serial, extension)
              if (sheetData !== null) {
                const isUploaded = await createSheetsChart(sheetData, chartScope)
                const isUpdated = await updateSheetProperties(sheetData)
                if (isUploaded && isUpdated) {
                  console.log('')
                  console.log('07. Export Excel')
                  const dest = fs.createWriteStream(uploadDir + fileName + extension)
                  const sheetStream = await exportSheetFile(sheetData.id)

                  if (sheetStream) {
                    sheetStream.on('error', err => {
                      res.json({ status: 'error', message: 'RRMSE 분석 결과를 엑셀로 다운로드하지 못했습니다.' })
                    }).on('end', () => {
                      console.log('')
                      console.log('08. Done')
                      setTimeout(() => {
                        res.download(uploadDir + fileName + extension)
                      }, 100)
                    }).pipe(dest)
                  } else {
                    res.json({ status: 'error', message: 'RRMSE 분석 결과를 엑셀로 다운로드하지 못했습니다.' })
                  }
                } else {
                  res.json({ status: 'error', message: 'RRMSE 분석 결과를 엑셀로 업로드하지 못했습니다.' })
                }
              } else {
                res.json({ status: 'error', message: 'RRMSE 분석 결과를 엑셀로 변환하지 못했습니다.' })
                console.log('[Error] RRMSE 분석 결과를 엑셀로 변환하지 못했습니다.')
              }
            }
          })
        } else {
          res.json({ status: 'error', message: 'RRMSE 분석 결과가 존재하지 않습니다.' })
          console.log('[Error] RRMSE 분석 결과가 존재하지 않습니다.')
        }
      } catch (e) {
        res.json({ status: 'error', message: '[Error] ' + e.message })
        console.log('[Error] ' + e.message)
      }
    })
  } else {
    res.json({ status: 'error', message: '[Error] Cannot get idx' })
    console.log('[Error] Cannot get idx')
  }
})

let cbl = new CBL(app)
let GoogleDrivers = null
let GoogleSheets = null
 
let server = app.listen(3001, function () {
   let host = server.address().address
   let port = server.address().port

   moment.tz.setDefault('Asia/Seoul')

   fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('[Error] Error loading client secret file:', err)
    authorize(JSON.parse(content), initialization)
  })

   console.log("00. Wait DRMS RRMSE/CBL Export Excel Server [%s:%s]", host, port)
})

console.log('00. Start DRMS RRMSE/CBL Export Excel Server')
