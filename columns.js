const _COLUMS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

const columns = () => {
  const c = {}
  _COLUMS.map((a, i) => {
    c[a] = i + 1
  })
  return c
}

module.exports = columns