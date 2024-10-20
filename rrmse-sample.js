const raw = require('./raw-sample')
const _ = require('lodash')

const sample = {
    "raw_report": {
        "meta": {
            "startDate": "2019.10.02",
            "endDate": "2021.03.18",
            "evaluateDate": "2021.03.19",
            "siteNm": "테스트",
            "kepcoCustomerId": "1316115920",
            "basicWon": 0,
            "basicUnitPrice": 0, 
            "contractElec": 2700,
            "priceApplyKw": 845,
            "contractClass": "산업용(을)",
            "supplyMethod": "고압A"
        },
        "data": raw
    },
    "primary_report": {
        "meta": {
            "startDate": "2019.10.02",
            "endDate": "2021.03.18",
            "evaluateDate": "2021.03.19",
            "avgDemKw": 0,
            "abnormalWorkingDays": ["2020-01-01", "2020-01-02"],
            "contractElec": 2700,
            "priceApplyKw": 845,
            "contractClass": "산업용(을)",
            "supplyMethod": "고압A",
            "siteNm": "테스트",
            "kepcoCustomerId": "1316115920",
            "bp": 45340,
            "smp": 90,
            "dutyDrTime": 4,
            "voluntaryDrTime": 36,
            "paymentRatio": 0.7
        },
        "usage": {
            "max": 838.56,
            "mid": 59.54085395327416,
            "min": 0
        },
        "rrmse": {
            "max": 11,
            "mid": 12
        },
        "data": { // 60일 사용량/cbl 데이터
            "days": {
                "usage": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                "cbl": {
                    "max": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                    "mid": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) }))
                }
            },
            "seasonal": {
                "usage": { // winter: 12~2, spring: 3~5, summer: 6~8, autumn: 9~11
                    "winter": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                    "spring": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                    "summer": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                    "autumn": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) }))
                },
                "cbl": {
                    "max": {
                      "winter": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                      "spring": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                      "summer": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                      "autumn": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) }))	
                    },
                    "mid": {
                      "winter": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                      "spring": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                      "summer": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) })),
                      "autumn": raw.slice(0, 60 * 24).map(data => ({ timestamp: data.timestamp, value: _.random(1, 100) }))	
                    }
                },
                "avg": { // hour: 1~24
                    "winter": _.range(1, 25).map(hour => ({ hour, value: _.random(1, 100) })),
                    "spring": _.range(1, 25).map(hour => ({ hour, value: _.random(1, 100) })),
                    "summer": _.range(1, 25).map(hour => ({ hour, value: _.random(1, 100) })),
                    "autumn": _.range(1, 25).map(hour => ({ hour, value: _.random(1, 100) }))	
                }
            }
        }
    },
    "secondary_report": {
        "meta": {
            "startDate": "2019.10.02",
            "endDate": "2021.03.18",
            "evaluateDate": "2021.03.19",
            "siteNm": "테스트",
            "kepcoCustomerId": "1316115920",
            "basicWon": 0,
            "basicUnitPrice": 0, 
            "contractElec": 2700,
            "priceApplyKw": 845,
            "contractClass": "산업용(을)",
            "supplyMethod": "고압A"
        },
        "data": []
    }
}

module.exports = sample