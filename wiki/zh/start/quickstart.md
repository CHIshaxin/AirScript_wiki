---
title: "开始"
source: https://airsheet.wps.cn/docs/start/quickstart.html
updated: 1789700916000
hash: BLTQmnOs
---

# 开始

## 打开AirScript编辑器

1. 在金山文档首页新建一个表格并打开来体验AirScript。
2. 打开表格之后，在上方**效率**-**AirScript编辑工具**弹出编辑页面。
3. 将下方的例子，逐个运行，查看效果来快速上手AirScript。

## HelloWorld

### console.log()

```js
console.log("hello world!")
```

### console.error()

```js
console.error("hello world!")
```

## 文档操作

### 遍历文件的SheetName，切换对应的Sheet，读取单元格，修改单元格

```js
// 遍历并打印所有工作表的名称
let sheets = Application.Sheets
for (let i = 0; i < sheets.Count; i++) {
    let sheet = sheets.Item(i + 1)
    console.log(sheet.Name) // 打印每个工作表的名称
}

// 打印当前激活Sheet的名称
console.log(Application.ActiveSheet.Name) // 打印当前激活Sheet的名称

// 获取当前激活工作表的A1单元格
let A1 = Application.ActiveSheet.Range("A1")

// 打印A1单元格内容
console.log(A1.Text) // 打印A1单元格的内容

// 修改A1单元格内容
A1.Value2 = "bar"
console.log(A1.Text) // 打印修改后的A1单元格内容

// 修改A1单元格的背景颜色为黄色(255,255,0)
A1.Interior.Color = RGB(255,255,0) // 高亮A1单元格
```

### 将文档所有非空单元格增加指定后缀

```js
// 通过Application.ActiveSheet.UsedRange获得用户使用区域
const usedRange = Application.ActiveSheet.UsedRange
const startRow = usedRange.Row // 获取起始行
const startCol = usedRange.Column // 获取起始列
const endRow = startRow + usedRange.Rows.Count // 获取终止行
const endCol = startCol + usedRange.Columns.Count // 获取终止列

for (let i = startRow; i < endRow; i++) {
    const row = Application.ActiveSheet.Rows(i) // 确定行
    for (let j = startCol; j < endCol; j++) {
        const rg = row.Columns(j) // 从行对象中指定列，从而确定单元格
        const text = rg.Text
        // 如果单元格非空，则添加foo后缀
        if (text !== '') {
            rg.Value2 = text + 'foo' // 修改单元格内容
        }
    }
}

// 遍历UsedRange中的每个单元格并打印其内容
for (let i = startRow; i < endRow; i++) {
    for (let j = startCol; j < endCol; j++) {
        const cell = Application.ActiveSheet.Cells(i, j)
        console.log(cell.Text) // 打印所有单元格数据
    }
}
```

### 新建Sheet

```js
// 当前工作表名称
const currentSheetName = Application.ActiveSheet.Name

// 在当前Sheet之后新增两个默认名称的工作表
Application.Sheets.Add(null, Application.Sheets(currentSheetName), 2)

// 在当前Sheet之后新增名称为'新工作表(右)'的Sheet
const newSheetRight = Application.Sheets.Add(null, Application.Sheets(currentSheetName), 1)
newSheetRight.Name = '新工作表(右)'

// 在当前Sheet之前新增名称为'新工作表(左)'的Sheet
const newSheetLeft = Application.Sheets.Add(Application.Sheets(currentSheetName), null, 1)
newSheetLeft.Name = '新工作表(左)'
```

## 更多

更多真实例子可通过[示范案例](../example/excel.md)尝试。
