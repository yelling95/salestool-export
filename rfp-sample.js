const sample = {
  meta: {
    startDate: '2019. 02. 02', // 검증 기간 (시작일)
    endDate: '2019. 02. 03', // 검증 기간 (종료일)
    siteNm: '테스트사업장', // 사업장명
    kepcoCustomerId: '345345', // 고객번호
    evaluateDate: '2019.01.01', // 기준일
    workdayUsageMax: 34333, // 평일 사용량 MAX
    workdayUsageAvg: 3454, // 평일 사용량 AVG
    workdayUsageMin: 23423, // 평일 사용량 MIN
    basicWon: 2300, // 기본요금
    basicUnitPrice: 45000, // 기본요금단가
    priceApplyKw: 34000, // 요금적용전력
    contractClass: '계약1', // 요금종류
    usageYearMax: 2000, // 연간전기사용량(MAX)
    usageYearSum: 2000, // 연간전기사용량(SUM)
    rrmseMax: 20.3, // RRMSE(MAX)
    rrmseMid: 234, // RRMSE(MID)
    holidayMinLoad: 3400, // 휴일기저부하
    workdayLoad: 2000, // 업무일 평균 부하
    holidayAvgLoad: 3000, // 휴일 평균 부하,
    reduction1Type: '감축방법1', // 감축방법1(업종에 따른 감축방법)
    reduction2Type: '감축방법2', // 감축방법1(업종에 따른 감축방법)
    reduction3Type: '감축방법3', // 감축방법1(업종에 따른 감축방법)
    reduction4Type: '감축방법4', // 감축방법1(업종에 따른 감축방법)
    reduction5Type: '감축방법5' // 감축방법1(업종에 따른 감축방법)
  },
  data: {
    monthlyUsage: [
      { month: '2020.01', value: 2300 },
      { month: '2020.02', value: 434 },
      { month: '2020.03', value: 244 },
      { month: '2020.04', value: 2322 },
      { month: '2020.05', value: 2311 },
      { month: '2020.06', value: 2100 },
      { month: '2020.07', value: 2000 },
      { month: '2020.08', value: 3300 },
      { month: '2020.09', value: 300 },
      { month: '2020.10', value: 4500 },
      { month: '2020.11', value: 600 },
      { month: '2020.12', value: 700 }
    ],
    hourlyTotalUsage: [
      { hour: 1, value: [ 2300, 3300, 300 ] }, // MAX, MID, MIN
      { hour: 2, value: [ 2300, 343, 700 ] },
      { hour: 3, value: [ 2300, 3300, 222 ] },
      { hour: 4, value: [ 2300, 343, 222 ] },
      { hour: 5, value: [ 4500, 343, 700 ] },
      { hour: 6, value: [ 2300, 343, 222 ] },
      { hour: 7, value: [ 2300, 3300, 700 ] },
      { hour: 8, value: [ 2300, 343, 222 ] },
      { hour: 9, value: [ 2300, 3300, 300 ] },
      { hour: 10, value: [ 4500, 343, 700 ] },
      { hour: 11, value: [ 2300, 343, 700 ] },
      { hour: 12, value: [ 2300, 343, 222 ] },
      { hour: 13, value: [ 2300, 3300, 700 ] },
      { hour: 14, value: [ 4500, 343, 222 ] },
      { hour: 15, value: [ 2300, 343, 700 ] },
      { hour: 16, value: [ 2300, 3300, 222 ] },
      { hour: 17, value: [ 2300, 343, 222 ] },
      { hour: 18, value: [ 2300, 343, 300 ] },
      { hour: 19, value: [ 2300, 343, 300 ] },
      { hour: 20, value: [ 2300, 343, 300 ] },
      { hour: 21, value: [ 4500, 343, 700 ] },
      { hour: 22, value: [ 2300, 3300, 222 ] },
      { hour: 23, value: [ 2300, 343, 700 ] },
      { hour: 24, value: [ 2300, 343, 222 ] }
    ],
    hourlyMaxUsage: [
      { hour: 1, value: 2300 },
      { hour: 2, value: 434 },
      { hour: 3, value: 244 },
      { hour: 4, value: 2322 },
      { hour: 5, value: 2311 },
      { hour: 6, value: 2100 },
      { hour: 7, value: 2000 },
      { hour: 8, value: 3300 },
      { hour: 9, value: 300 },
      { hour: 10, value: 4500 },
      { hour: 11, value: 600 },
      { hour: 12, value: 700 },
      { hour: 13, value: 2300 },
      { hour: 14, value: 434 },
      { hour: 15, value: 244 },
      { hour: 16, value: 2322 },
      { hour: 17, value: 2311 },
      { hour: 18, value: 2100 },
      { hour: 19, value: 2000 },
      { hour: 20, value: 3300 },
      { hour: 21, value: 300 },
      { hour: 22, value: 4500 },
      { hour: 23, value: 600 },
      { hour: 24, value: 700 }
    ]
  }
}

module.exports = sample