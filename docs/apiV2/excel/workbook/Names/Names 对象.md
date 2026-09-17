---
title: "Names 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Names/Names%20%E5%AF%B9%E8%B1%A1.html
---
# Names (对象)

应用程序或工作簿中所有 **[Name](./../Name/Name 对象.html)** 对象的集合。

## 说明

每一个 **[Name](./../Name/Name 对象.html)** 对象都代表一个单元格区域的定义名称。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

*RefersTo* 参数必须以 A1 样式表示法指定，包括必要时使用的美元符 ($)。例如，如果在 Sheet1 上选定了单元格 A10，并且通过将 *RefersTo* 参数“=Sheet1!A1:B1”而定义了一个名称，那么该新名称实际上指向单元格区域 A10:B10（因为指定的是相对引用）。若要指定绝对引用，请使用“=Sheet1!$A$1:$B$1”。

## 示例

使用 **Names** 属性可返回 **[Names](./../Names/Names 对象.html)** 集合。下例创建活动工作簿中所有名称及其引用地址的列表。

```javascript
function test() {
    let nms = Application.ActiveWorkbook.Names
    let wks = Application.Worksheets.Item(1)
    for(let r = 1; r <= nms.Count; r++){
        wks.Cells.Item(r, 2).Value2 = nms.Item(r).Name
        wks.Cells.Item(r, 3).Value2 = nms.Item(r).RefersToRange.Address
    }
}
```

使用 **Add**方法可创建一个名称并将它添加到集合。下例创建一个新名称，该名称引用名为“Sheet1”的工作表上的单元格 A1:C20。

```javascript
function test() {
    Application.Names.Add ("test1", "=sheet1!$a$1:$c$20")
}
```

使用 **Names**(*index*)（其中 *index* 是名称索引号或定义名称）可返回一个 **[Name](./../Name/Name 对象.html)** 对象。下例从活动工作簿中删除名称“mySortRange”。

```javascript
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
}
```
