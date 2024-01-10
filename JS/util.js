
export {

}

// 表单校验
var InputStrategy = (function () {
  var strategy = {
    notNull: function (value, text) {
      return !value ? `${text || '请输入内容'}` : ""
    },
    number: function (value, text) {
      return /^[1-9]\d*$/.test(value) ? "" : `${text || "请输入数字"}`
    },
    valiableName: function (value, text) {
      return /^[0-9a-zA-Z]+$/.test(value) ? "" : `${text}`
    }
  }
  return {
    check: function (type, value, content) {
      value = String(value).trim();
      return strategy[type] ? strategy[type](value, content) : "没有该类型的检测方法"
    },
    addStrage: function (type, fn) {
      strategy[type] = fn
    }
  }
})()


/**
 * 计算data中的数据(模拟Excel)
 * @param {object} data 待计算的数据 
 */
function calData(data) {
  const result = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const addRow = [];
    for (let j = 0; j < row.length; j++) {
      const cellValue = row[j];
      // 是否需要计算
      if (typeof cellValue === "object") {
        const expressions = cellValue.f // 拿到计算用的表达式 例如: Q4+Q3-R2
        // 拆分表达式并转换为索引值  
        const parts = expressions.split(/[+\-]/).map(part => {
          // 提取单元格引用，例如 "R4" 或 "S2"  
          const cellRef = part.match(/[A-Z]+[0-9]+/)[0]
          // 将单元格引用转换为索引值  
          const index = this.convertCellReferenceToIndex(cellRef)
          var value = 0
          // 根据索引值获取对应的单元格值  
          if (index[0] == i && index[1] == j - 1)
            value = addRow[index[1]] || 0
          else
            value = result[index[0]][index[1]] || 0
          return value
        })
        const operators = expressions.match(/[\+\-]/g) // 操作符
        // 计算表达式的值  
        const resultValue = this.calculate(parts, operators) // 计算表达式的值  
        addRow.push(resultValue)
      }
      else {
        addRow.push(cellValue)
      }
    }
    result.push(addRow)
  }
  return result;
}

/**
 * 行列转化成索引
 * @param {String} cellRef 单元格
 */
function convertCellReferenceToIndex(cellRef) {
  var col = cellRef.match(/[A-Z]+/)[0];
  var row = parseInt(cellRef.match(/\d+/)[0]);
  var colIndex = 0;
  for (var i = 0; i < col.length; i++) {
    colIndex = colIndex * 26 + col.charCodeAt(i) - 64;
  }
  return [row - 1, colIndex - 1]
}

/**
 * 计算值  
 * 
 * @param {*} numbers 值, [1,2,3]
 * @param {*} operators 操作符 ['+','-']
 * @returns {Number} 表达式的计算结果   1+2-3
 */
function calculate(numbers, operators) {
  let expression = ''
  for (let i = 0; i < numbers.length - 1; i++) {
    expression += numbers[i] + operators[i % operators.length];
  }
  expression += numbers[numbers.length - 1]
  return eval(expression)
}


/**
  * 将字符串转换为小驼峰命名法
  * @param {string} inputString - 需要转换的字符串
  * @returns {string} - 转换后的小驼峰命名法字符串
  */
convertToCamelCase(inputString) {
  return inputString.replace(/_(\w)/g, function (match, group1) {
    return group1.toUpperCase();
  }).replace(/^[A-Z]/, function (match) {
    return match.toLowerCase();
  });
}