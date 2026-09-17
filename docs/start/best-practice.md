---
title: "最佳实践"
source: https://airsheet.wps.cn/docs/start/best-practice.html
---
# 最佳实践

## 体验优化

#### 通过超链接运行脚本

脚本的运行可以通过超链接触发，绑定方法如下

1. 打开**插入**-**链接**弹出插入链接配置页面。
2. 在文本输入框中填写合适的提示文本，描述该脚本的功能。
3. 点击**类型**旁边的下拉框弹出选项选择**AirScript脚本**再指定想绑定的脚本。
4. 点击**确定**。

点击插入的超链接即可运行脚本。

通过这种方式，用户所见即所得，相比于打开脚本编辑器点击运行友好很多。

#### 更符合直觉的运行体验

脚本需要选定某些单元格范围时，有多种实现方式，通过`Application.Selection`读取并操作用户选中单元格是最符合用户直觉的运行体验。

下面是通过脚本变量划定遍历范围的例子。该例子每次修改范围都需要修改4个变量，非常不方便，体检较差。

```javascript
// 错误例子
const startRow = 1
const startColumn = 1
const endRow = 3
const endColumn = 3
for(let i = startRow; i <= endRow; i++){
  let row=Application.Rows(i)
  for (let j = startColumn; j <= endColumn; j++){
    console.log(row.Columns(j).Text)
  }
}
```

下面是通过`Application.Selection`获取用户在文档界面选中范围单元格，使用方便并符合直觉。

```javascript
// 正确例子
let selection = Application.Selection
let startRow = selection.Row
let startCol = selection.Column
let endRow = startRow + selection.Rows.Count - 1
let endCol = startCol + selection.Columns.Count - 1

for (let i = startRow; i <= endRow; i++) {
    let row = Application.Rows(i)
    for (let j = startCol; j <= endCol; j++) {
        console.log(row.Columns(j).Text)
    }
}
```

## 性能优化

#### 尽可能地复用对象

对Application的每次函数调用会使用到脚本引擎去操作文件数据，重复调用会造成性能浪费。

如同样实现遍历并打印100\*100单元格的内容，复用对象能使性能提升一倍。

无复用对象:

```javascript
// 错误例子
let start = new Date()
for(let i= 1;i<=100;i++){
  for (let j =1;j<=100;j++){
    Application.Rows(i).Columns(j).Text
  }
}
console.log(new Date()-start,'ms')
```

更好的写法:

```javascript
// 正确例子
let start = new Date()
for(let i= 1;i<=100;i++){
  let row=Application.Rows(i)
  for (let j =1;j<=100;j++){
    row.Columns(j).Text
  }
}
console.log(new Date()-start)
```

#### 使用UsedRange缩小遍历范围

上面更符合直觉的运行体验的例子中， 全选整个表格则会进行上万亿次遍历,即使只选中一整行也需要进行上百万次遍历， 但实际上我们单元表中使用到范围并没有那么大。 配合使用`Application.ActiveSheet.UsedRange`可以确认工作簿的使用范围，因此大大加快脚本执行时间。

下面是该例子的改进版本：

```javascript
const selection = Application.Selection
const usedRange = Application.ActiveSheet.UsedRange

// 确定遍历的范围
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = Application.ActiveSheet.Rows(i) // 复用对象
    for (let j = colFrom; j <= colTo; j++) {
        console.log(row.Columns(j).Text)
    }
}
```
