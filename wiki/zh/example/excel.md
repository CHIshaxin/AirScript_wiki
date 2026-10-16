---
title: "表格案例"
source: https://airsheet.wps.cn/docs/example/excel.html
updated: 1789700916000
hash: db-riDEh
---

# 表格案例

为了帮助您理解AirScript，以及快速上手脚本开发，我们将表格中**开始**-**快捷工具**中的部分功能源码展示出来， 希望您能在阅读完这些真实的案例代码后，能有效利用脚本编辑器来解决问题。

## 清除公式仅保留值

手动选中一个或多个单元格，然后执行脚本，执行完成后会批量清除其公式，仅保留其值。

```js
const API = Application

// 用户选区
const selection = API.Selection
// 获取用户激活工作簿的使用范围
const usedRange = API.ActiveSheet.UsedRange

// 取二者的最小集合，即所需遍历最小集合
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
  const row = API.ActiveSheet.Rows(i) // 确定行
  for (let j = colFrom; j <= colTo; j++) {
    const rg = row.Columns(j) // 从行对象中指定列，从而确定单元格
    // Text是单元格的显示值，即如果该单元格原先是公式，那Text即公式计算后的结果
    // 如果想获取未计算的结果，使用rg.Formula
    const text = rg.Text
    if (text !== '') {
      rg.Value2 = text // 原先的Value2是公式，通过将Value2重写成计算后的结果从而清除公式
    }
  }
}
```

## 高亮错误手机号

手动选中一个或多个单元格，然后执行脚本，脚本会自动判断单元格的内的手机号是否正确，执行完成后会高亮错误手机号单元格。

```js
// 手机号码的正则表达式,匹配第一位是1,第二位是3-9,其余位是数字的11位字符串
const PhoneReg = /^1[3-9](\d{9})$/i

// 高亮颜色为黄色
const color = RGB(255, 255, 0)

// API 的简化引用
const API = Application

// 用户选择的区域
const selection = API.Selection

// 工作簿激活的范围
const usedRange = API.ActiveSheet.UsedRange

// 取二者的最小集合，即所需遍历最小集合
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // 确定行对象
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // 再确定列对象,即确定单元格
        const text = rg.Text  // 获取单元格的文本内容
        if (text !== '' && !PhoneReg.test(text)) { // 正则表达式匹配不通过的情况下
            rg.Interior.Color = color // 用黄色高亮该单元格
        }
    }
}
```

## 取消合并填充相同内容

手动选中一个或多个单元格，然后执行脚本，脚本会自动取消合并单元格并填充相同内容。

```js
const API = Application
// 用户选择的区域
const selection = API.Selection

// 如果选区内有合并的单元格
if (selection.MergeCells) {
  // 获取所有合并的单元格区域
  const areas = selection.MergeArea.Areas
  const areaCount = areas.Count

  // 遍历每个合并的单元格区域
  for (let areaIndex = 1; areaIndex <= areaCount; areaIndex++) {
    const area = areas.Item(areaIndex)
    const text = area.Formula // 保存取消合并前该合并单元格的内容

    // 获取合并单元格区域的范围
    const rowFrom = area.Row
    const rowTo = rowFrom + area.Rows.Count - 1
    const colFrom = area.Column
    const colTo = colFrom + area.Columns.Count - 1

    // 取消合并
    area.UnMerge()
    // 将原先合并单元格的区域内容改写成合并单元格的值
    for (let i = rowFrom; i <= rowTo; i++) {
      const row = API.ActiveSheet.Rows(i) // 确定行
      for (let j = colFrom; j <= colTo; j++) {
        const cell = row.Columns(j) // 确定列
        cell.Value2 = text // 将之前保存的内容写到单元格里
      }
    }
  }
}
```

## 统计重复次数

手动选中一个或多个单元格，然后执行脚本，脚本会自动计算选区内的重复的值和重复的次数，新建工作表并将重复的值和重复的次数写入新建工作表的`A2单元格`和`B2单元格`。

```js
const API = Application
// 用户选择的区域
const selection = API.Selection
// 工作簿激活的范围
const usedRange = API.ActiveSheet.UsedRange

// 取二者的最小集合，即所需遍历最小集合
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

// countMap用于统计每个字符串出现次数
const countMap = {}
for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // 确定行
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // 再确定列对象,即确定单元格
        const text = rg.Text // 取出该单元格的值用于统计
        if (text) {
            if (countMap[text]) {
                countMap[text]++
            } else {
                countMap[text] = 1
            }
        }
    }
}

const Name = API.ActiveSheet.Name
// 在当前工作表之后新建一个工作表，使用默认名称
const newSheet = API.Sheets.Add(null, API.Sheets(Name), 1)

// 在新工作表的A1和B1写入'重复值'和'重复次数'
newSheet.Range('A1').Value2 = '重复值'
newSheet.Range('B1').Value2 = '重复次数'

// 从第二行开始,每行第一列输出字符串,第二列输出该字符串重复出现的次数
let index = 2
for (let k in countMap) {
    const row = newSheet.Rows(index)
    row.Columns(1).Value2 = "'" + k
    row.Columns(2).Value2 = countMap[k] + ''
    index++;
}
```
