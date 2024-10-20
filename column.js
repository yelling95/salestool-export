const _ = require('lodash')

const COL = (() => {
  const column = {}
  const keys = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
  _.range(1, 27).map((col, colIndex) => {
    column[keys[colIndex]] = col
  })
  return column
})()

module.exports = COL