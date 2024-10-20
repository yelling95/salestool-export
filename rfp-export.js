const fs = require('fs')
const readline = require('readline')
const { google } = require('googleapis')
const express = require('express')
const app = express()
const request = require('request')
const _ = require('lodash')
const moment = require('moment')
const timezone = require('moment-timezone')
const randomstring = require('randomstring')
// const sample = require('./rfp-sample')

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
  'https://www.googleapis.com/auth/presentations',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/spreadsheets'
]

const ROOT_FOLDER = '11vxjqszdhHTh9z1h5tundFbxSWJyQx1v'

const TEMPLETES = [
  '1t4UnnWaHk0mTo_V2iotxSZLQI7mtKzxEFPSa_OBr5Sk', // ppt
  '1J0R6JZEWGEpxMsposuIo1SSkwcT-xWT5QLj5K3ovH-w', // excel-bar-monthly
  '1PGuQjyVXmfClw4_trpNaNAUrTZatRQ1UCERmYjJJqgU', // excel-line-hourly
  '13cOFIDxFbAwxPhrDBAPyAOq2AHv05JaapEZqFFlu6I0' // excel-bar-hourly
]

const CODE = {
  SUCCESS: 200
}

const SLIDE_CHART = [14, 15, 16]
const SLIDE_TABLE = [16]

const TOKEN_PATH = 'token.json'

const uploadDir = './temp/rfp/'
const extension = '.pptx'

app.use(function(req, res, next) {
  res.header('Access-Control-Allow-Origin', '*')
  res.header('Access-Control-Allow-Methods', 'GET, PUT, POST, DELETE, OPTIONS')
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, Content-Length, X-Requested-With')

  if ('OPTIONS' === req.method) {
    res.sendStatus(200)
    res.end()
  } else {
    next()
  }
})

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
  GoogleSlides = google.slides({ version: 'v1', auth })
  GoogleSheets = google.sheets({ version: 'v4', auth })
}

function errorCallback(res) {
  res.json({ status: 'error', message: 'RFP 분석 결과 변환중 오류가 발생했습니다.' })
}

async function createFolder(serial) {
  try {
    const result = await GoogleDrivers.files.create({
      resource: {
        name: serial,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [ROOT_FOLDER]
      }
    })
  
    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Created Google Drive Folder [ ' + serial + ' ]')
      return null
    } else {
      console.log('[Success] Created Google Drive Folder [ ' + result.data.id + ' ]')
      return result.data.id
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function copyTempleteFile(folderId, serial) {
  try {
    let created = {}
    let excels = []
    await Promise.all(
      TEMPLETES.map(async (tempId, tempIdx) => {
        const result = await GoogleDrivers.files.copy({
          fileId: tempId,
          resource: {
            name: serial + '_' + tempIdx,
            parents: [folderId]
          }
        })
      
        if (result.status !== CODE.SUCCESS) {
          console.log('[Error] Cannot Copy Google Drive File [ ' + serial + '_' + tempIdx + ' ]')
        } else {
          console.log('[Success] Created Google Drive File [ ' + serial + '_' + tempIdx + ' ]')
          if (result.data.mimeType.indexOf('spreadsheet') > -1) {
            excels.push(Object.assign(result.data, { tempIdx: tempIdx, tempId: tempId }))
          } else if (result.data.mimeType.indexOf('presentation') > -1) {
            created.ppt = Object.assign(result.data, { tempIdx: tempIdx, tempId: tempId })
          }
        }
      })
    )
    console.log('')
    created.excels = excels
    return created
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function getFormatData(data) {
  try {
    let rowData = []
    await Promise.all(
      data.map(d => {
        let cellData = []
        if (d.month) {
          cellData.push({
            userEnteredValue: {
              stringValue: d.month
            }
          })
        }

        if (d.hour) {
          cellData.push({
            userEnteredValue: {
              numberValue: d.hour
            }
          })
        }
        
        if (d.value && d.value.constructor === Array) {
          d.value.map(val => cellData.push({
            userEnteredValue: {
              numberValue: val
            }
          }))
        } else {
          cellData.push({
            userEnteredValue: {
              numberValue: d.value ? d.value : 0
            }
          })
        }
        rowData.push({
          values: cellData
        })
      })
    )
    return rowData
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function getFormatTableData(data) {
  try {
    let tableData = []
    let totalBasicWon = 0
    let totalIncenWon = 0
    let totalSum = 0

    await Promise.all(
      data.map(async (values, rowIndex) => {
        totalBasicWon += (values['basicWon'] ? values['basicWon'] : 0)
        totalIncenWon += (values['incentiveWon'] ? values['incentiveWon'] : 0)
        totalSum += (values['sum'] ? values['sum'] : 0)

        _.mapKeys(values, (value, key) => {
          tableData.push({
            replaceAllText: {
              containsText: {
                text: '{{a' + rowIndex + key + '}}',
                matchCase: true
              },
              replaceText: value + ''
            }
          })
        })
      })
    )

    tableData.push({
      replaceAllText: {
        containsText: {
          text: '{{totalBasicWon}}',
          matchCase: true
        },
        replaceText: totalBasicWon + ''
      }
    })

    tableData.push({
      replaceAllText: {
        containsText: {
          text: '{{totalIncenWon}}',
          matchCase: true
        },
        replaceText: totalIncenWon + ''
      }
    })

    tableData.push({
      replaceAllText: {
        containsText: {
          text: '{{totalSum}}',
          matchCase: true
        },
        replaceText: totalSum + ''
      }
    })
    return tableData
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function getFormatTextData(data) {
  try {
    let textData = []
    let rrmseMax = data['rrmseMax'] ? data['rrmseMax'] * 100 : 0
    let rrmseMid = data['rrmseMid'] ? data['rrmseMid'] * 100 : 0
    let partAmountMax = data['partAmountMax'] ? data['partAmountMax'] : 300
    let partAmountMid = data['partAmountMid'] ? data['partAmountMid'] : 500
    let rrmseStandard = rrmseMax > rrmseMid ? rrmseMid : rrmseMax
    let rfpRrmse = 0
    let rfpUsage = 0

    // 제안 RRMSE는 RRMSE 30% 미만일 경우, 참여용량(partAmount)로 비교해서 결정
    // 두 중 하나라도 30% 미만이 아닐 경우 RRMSE 값이 낮은걸로 결정
    // 제안용량은 RRMSE 30% 미만일 경우, 참여용량(partAmount)이 큰 값으로 결정
    // 두 중 하나라도 30% 미만이 아닐 경우 값이 작은 RRMSE의 참여용량으로 결정
    if (rrmseMax < 30 && rrmseMid < 30) {
      rfpRrmse = partAmountMax > partAmountMid ? rrmseMax : rrmseMid
      rfpUsage = partAmountMax > partAmountMid ? partAmountMax : partAmountMid
    } else {
      rfpRrmse = rrmseMax > rrmseMid ? rrmseMid : rrmseMax
      rfpUsage = rrmseMax > rrmseMid ? partAmountMid : partAmountMax
    }

    _.mapKeys(data, (value, key) => {
      let cVal = ''
      if (value) {
        if (value.constructor === Number) {
          if (key === 'usageYearSum') {
            cVal = _.round(value / 1000, 2) + ''
          } else if (value.toString().indexOf('.') > -1) {
            if (key.indexOf('rrmse') > -1) cVal = _.round(value * 100, 2) + ''
            else cVal = _.round(value, 2) + ''
          } else {
            cVal = convertThousandSeparator(value)
          }
        } else {
          cVal = value
        }
      }

      textData.push({
        replaceAllText: {
          containsText: {
            text: '{{' + key + '}}',
            matchCase: true
          },
          replaceText: cVal
        }
      })
    })

    textData.push({
      replaceAllText: {
        containsText: {
          text: '{{rrmseResult}}',
          matchCase: true
        },
        replaceText: getPercentScope(rrmseStandard)
      }
    })

    textData.push({
      replaceAllText: {
        containsText: {
          text: '{{rfpRrmse}}',
          matchCase: true
        },
        replaceText: _.round(rfpRrmse, 2) + ''
      }
    })

    textData.push({
      replaceAllText: {
        containsText: {
          text: '{{rfpUsage}}',
          matchCase: true
        },
        replaceText: _.round((rfpUsage / 1000), 2) + ''
      }
    })

    textData.push({
      replaceAllText: {
        containsText: {
          text: '{{YYYY-MM}}',
          matchCase: true
        },
        replaceText: moment().format('YYYY. MM')
      }
    })
    
    textData.push({
      replaceAllText: {
        containsText: {
          text: '{{YY}}',
          matchCase: true
        },
        replaceText: moment().format('YY')
      }
    })
    return textData
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

function getPercentScope (standard) {
  let scope = ''

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

function convertThousandSeparator (number) {
  return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

async function updateSheetsData(data, excels) {
  const dataKeys = ['', 'monthlyUsage', 'hourlyTotalUsage', 'hourlyMaxUsage']
  const gridIds = [0, 1654711040, 0, 1654711040]
  
  try {
    await Promise.all(
      excels.map(async excel => {
        const formatData = await getFormatData(data['data'][dataKeys[excel.tempIdx]])
        if (formatData) {
          const result = await GoogleSheets.spreadsheets.batchUpdate({
            spreadsheetId: excel.id,
            resource: {
              requests: [{
                updateCells: {
                  rows: formatData,
                  fields: '*',
                  start: {
                    sheetId: gridIds[excel.tempIdx],
                    rowIndex: 1,
                    columnIndex: 0
                  }
                }
              }]
            }
          })
          
          if (result.status !== CODE.SUCCESS) {
            console.log('[Error] Cannot Update Google Sheets [ ' + excel.name + ' ]')
          } else {
            console.log('[Success] Updated Google Sheets [ ' + excel.name + ' ]')
            const excelDetail = await GoogleSheets.spreadsheets.get({
              spreadsheetId: excel.id
            })
            if (excelDetail.status === CODE.SUCCESS) {
              excel.chart = excelDetail.data.sheets[0].charts[0]
            }
          }
        }
      })
    )
    return excels
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function getSlidesPages(ppt) {
  try {
    let slidesPages = null
    const slidesDetail = await GoogleSlides.presentations.get({
      presentationId: ppt.id
    })
    if (slidesDetail.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Get Google Slides File [ ' + ppt.id + ' ]')
    } else {
      slidesPages = slidesDetail.data.slides
    }
    return slidesPages
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function createdSheetsChart(fileObj, chartPages) {
  const chartOpts = [
    {},
    { width: 9.16, height: 11.17, x: 12.86, y: 4.74 }, 
    { width: 18.9, height: 7.6, x: 3.09, y: 9.12 }, 
    { width: 9.01, height: 9.83, x: 3.12, y: 5.75 }
  ]

  try {
    await Promise.all(
      fileObj.excels.map(async (excel, excelIdx) => {
        const result = await GoogleSlides.presentations.batchUpdate({
          presentationId: fileObj.ppt.id,
          resource: {
            requests: [{
                createSheetsChart: {
                  spreadsheetId: excel.id,
                  chartId: excel.chart.chartId,
                  linkingMode: 'LINKED',
                  elementProperties: {
                    pageObjectId: chartPages[excel.tempIdx - 1].objectId,
                    size: {
                      width: {
                        magnitude: chartOpts[excel.tempIdx].width * 360000,
                        unit: 'EMU'
                      },
                      height: {
                        magnitude: chartOpts[excel.tempIdx].height * 360000,
                        unit: 'EMU'
                      }
                    },
                    transform: {
                      scaleX: 1,
                      scaleY: 1,
                      translateX: chartOpts[excel.tempIdx].x * 360000,
                      translateY: chartOpts[excel.tempIdx].y * 360000,
                      unit: 'EMU'
                    }
                  }
                }
              }
            ]
          }
        })

        if (result.status !== CODE.SUCCESS) {
          console.log('[Error] Cannot Create Google Sheets Chart [ ' + excel.name + ' ]')
        } else {
          console.log('[Success] Create Google Sheets Chart [ ' + excel.name + ' ]')
        }
      })
    )
  } catch (e) {
    console.log('[Error] ' + e.message)
  }
}

async function updateSheetsTable(data, fileObj) {
  try {
    const formatData = await getFormatTableData(data['data']['settlementAnaly'])
    if (formatData) {
      const result = await GoogleSlides.presentations.batchUpdate({
        presentationId: fileObj.ppt.id,
        resource: {
          requests: formatData
        }
      })

      if (result.status !== CODE.SUCCESS) {
        console.log('[Error] Cannot Update Google Sheets Table [ ' + fileObj.ppt.id + ' ]')
      } else {
        console.log('[Success] Updated Google Sheets Table [ ' + fileObj.ppt.id + ' ]')
      }
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
  }
}

async function updateSheetsText(data, fileObj) {
  try {
    const formatData = await getFormatTextData(data['meta'])

    if (formatData) {
      const result = await GoogleSlides.presentations.batchUpdate({
        presentationId: fileObj.ppt.id,
        resource: {
          requests: formatData
        }
      })

      if (result.status !== CODE.SUCCESS) {
        console.log('[Error] Cannot Update Google Sheets Table [ ' + fileObj.ppt.id + ' ]')
      } else {
        console.log('[Success] Updated Google Sheets Table [ ' + fileObj.ppt.id + ' ]')
      }
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
  }
}

async function exportPdf(ppt) {
  try {
    const result = await GoogleDrivers.files.export({
      fileId: ppt.id,
      mimeType: 'application/pdf'
    }, {
      responseType: 'stream'
    })

    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Export PDF [ ' + ppt.id + ' ]')
      return null
    } else {
      console.log('[Success] Exported PDF [ ' + ppt.id + ' ]')
      return result.data
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

async function exportPPT(ppt) {
  try {
    const result = await GoogleDrivers.files.export({
      fileId: ppt.id,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    }, {
      responseType: 'stream'
    })

    if (result.status !== CODE.SUCCESS) {
      console.log('[Error] Cannot Export PDF [ ' + ppt.id + ' ]')
      return null
    } else {
      console.log('[Success] Exported PDF [ ' + ppt.id + ' ]')
      return result.data
    }
  } catch (e) {
    console.log('[Error] ' + e.message)
    return null
  }
}

function clearTempFiles(path, dirveFolderId) {
  try {
    fs.unlinkSync(path)
    GoogleDrivers.files.delete({
      fileId: dirveFolderId
    }, function(err) {
      if (err) {
        console.log('[Error] Cannot Remove Files ')
      } else {
        console.log('[Success] Removed All Files ')
      }
    })
  } catch (e) {
    console.log('[Error] ' + e.message)
  }
}

app.get('/download', function(req, res) {
  let token = req.headers.Authorization
  let key = req.query.key
  let fileName = req.query.fileName ? req.query.fileName : moment().format('YYYYMMDD') + '_DRMS 2차제안서'
  let delKey = req.query.delKey
  
  if (key) {
    try {
      if (!delKey) {
        res.download(uploadDir + key + extension, fileName + extension)
      } else {
        res.download(uploadDir + key + extension, fileName + extension, function(err) {
          if (!err) {
            clearTempFiles(uploadDir + key + extension, delKey)
          }
        })
      }
    } catch (e) {
      res.json({ status: 'error', message: '파일이 존재하지 않습니다.' })
    }
  } else {
    res.json({ status: 'error', message: '파일이 존재하지 않습니다.' })
  }
})

app.get('/convert', async function(req, res) {
  let token = req.headers.Authorization
  let idx = req.query.idx
  
  if (idx) {
    GET_OPTION.uri = API_URL + '/history/' + idx
    GET_OPTION.qs.type = 'secondary'

    request(GET_OPTION, async function (error, response, body) {
      try {
        if (response.statusCode === 200) {
          const data = JSON.parse(body)
          const serial = randomstring.generate(24)

          console.log('')
          console.log('01. Creating Google Drive Folder')
          let createdFolderId = await createFolder(serial)
          if (!createdFolderId) errorCallback(res)

          console.log('')
          console.log('02. Coping Google Drive Templete Files')
          let createdFiles = await copyTempleteFile(createdFolderId, serial)
          if (!createdFiles) errorCallback(res)
          
          console.log('')
          console.log('03. Update Google Sheets Data')
          let updatedExcels = await updateSheetsData(data, createdFiles.excels)
          if (!updatedExcels) errorCallback(res)
          
          console.log('')
          console.log('04. Get Google Slides Page Info')
          let slidesPages = await getSlidesPages(createdFiles.ppt)
          if (!slidesPages) errorCallback(res)
          let chartPages = slidesPages.filter((page, pageIdx) => _.indexOf(SLIDE_CHART, pageIdx + 1) > -1)
          let tablePages = slidesPages.filter((page, pageIdx) => _.indexOf(SLIDE_TABLE, pageIdx + 1) > -1)

          console.log('')
          console.log('05. Create Google Sheets Chart In Slides Page')
          await createdSheetsChart(createdFiles, chartPages)

          /*
          console.log('')
          console.log('06. Update Table Data In Slides Page')
          await updateSheetsTable(createdFiles)
          */

          console.log('')
          console.log('06. Update Text Data In Slides Page')
          await updateSheetsText(data, createdFiles)

          console.log('')
          console.log('07. Export PPT')
          const dest = fs.createWriteStream(uploadDir + serial + extension)
          let pptStream = await exportPPT(createdFiles.ppt)

          if (pptStream) {
            pptStream.on('error', err => {
              errorCallback(res)
            }).on('end', () => {
              console.log('')
              console.log('08. Done')
              res.send({
                status: 'success',
                key: serial,
                delKey: createdFolderId,
                siteNm: data.meta.siteNm,
                timestamp: moment().format('LLLL')
              })
            }).pipe(dest)
          } else {
            errorCallback(res)
          }

          /* PDF -> PPT
          console.log('')
          console.log('07. Export PDF')
          const dest = fs.createWriteStream(uploadDir + serial + '.pdf')
          let pdfStream = await exportPdf(createdFiles.ppt)

          if (pdfStream) {
            pdfStream.on('error', err => {
              errorCallback(res)
            }).on('end', () => {
              console.log('')
              console.log('08. Done')
              res.send({
                status: 'success',
                key: serial,
                delKey: createdFolderId,
                siteNm: data.meta.siteNm,
                timestamp: moment().format('LLLL')
              })
            }).pipe(dest)
          } else {
            errorCallback(res)
          }
          */
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

let GoogleDrivers = null
let GoogleSlides = null
let GoogleSheets = null
 
let server = app.listen(3002, function () {
  let host = server.address().address
  let port = server.address().port

  moment.tz.setDefault('Asia/Seoul')

  fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('[Error] Error loading client secret file:', err)
    authorize(JSON.parse(content), initialization)
  })

   console.log("00. Wait DRMS RFP Export Server [%s:%s]", host, port)
})

console.log('00. Start DRMS RFP Export Server')