---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CustomProperties/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/* 下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值*/
function test() {

    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
        console.log(cusProperties.Name + "\t" + cusProperties.Value)

}
```
