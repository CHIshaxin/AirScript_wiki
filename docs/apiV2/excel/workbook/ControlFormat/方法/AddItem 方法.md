---
title: "AddItem 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ControlFormat/%E6%96%B9%E6%B3%95/AddItem%20%E6%96%B9%E6%B3%95.html
---
# AddItem (方法)

为列表框或组合框添加项。

## 说明

使用本方法可清除由 **ListFillRange** 属性指定的任何区域。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Text | string | 必填 | 要添加的文本。 |
| Index | any | 可选 | 新输入项的位置。如果列表中输入项数目少于指定索引，则将列表末尾的空白项目添加到指定位置。如果省略该参数，则将项目追加到现有列表。 |

## 示例

```javascript
/*本示例创建一个列表框并用 1 到 10 的整数对其进行填充。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.AddFormControl(xlListBox, 100, 10, 100, 100)
    for (let x = 1; x <= 10; x++) {
        shape.ControlFormat.AddItem(x)
    }
}
```

```javascript
/*本示例创建一个组合框并按照指定索引对其进行填充。*/
function test() {
    let shape = ActiveSheet.Shapes.AddFormControl(xlDropDown, 500, 100, 110, 110)
    shape.ControlFormat.AddItem(1, 1)
    shape.ControlFormat.AddItem(2, 1)
}
```
