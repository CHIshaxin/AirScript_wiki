---
title: "Name 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/Name%20%E5%AF%B9%E8%B1%A1.html
---
# Name (对象)

代表单元格区域的定义名。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

## 说明

### 应用程序、工作簿和 Worksheet 对象

**[Name](./../Name/Name 对象.html)** 对象是 **Application**、 **[Workbook](./../Workbook/Workbook 对象.html)** 和 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象的 **[Names](./../Names/Names 对象.html)** 集合的成员。使用 **Names**(*index* )（其中 *index* 是名称索引号或定义名称）可返回一个 **[Name](./../Name/Name 对象.html)** 对象。

索引号表明名称在集合中的位置。名称按字母顺序从 a 到 z 放置，不区分大小写。

### Range 对象

虽然 **[Range](./../Range/Range 对象.html)** 对象可以有多个名称，但 **[Range](./../Range/Range 对象.html)** 对象没有 **[Names](./../Names/Names 对象.html)** 集合。对 **[Range](./../Range/Range 对象.html)** 对象调用 **Name** 属性可从名称列表（按字母顺序排序）中返回第一个名称。

## 示例

```javascript
/*本示例显示应用程序集合中第一个名称的单元格引用。*/
function test() {
    console.log(Application.Application.Names.Item(1).RefersTo)
}
```

```javascript
/*下例从活动工作簿中删除名称“mySortRange”。*/
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
}
```

```javascript
/*使用 Name 属性可返回或设置名称本身的文本。本示例更改活动工作簿中第一个 Name 对象的名称。*/
function test() {
    Application.Names.Item(1).Name = "stock_values"
}
```

```javascript
/*以下示例为第一张工作表上分配给单元格 A1:B1 的第一个名称设置 Visible 属性。*/
function test() {
    Application.Worksheets.Item(1).Range("A1:B1").Name.Visible = false
}
```
