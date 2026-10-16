---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CustomProperties/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BEZQsInG
---

# Add (方法)

添加自定义属性信息，返回一个代表自定义属性信息的 **[CustomProperty](./../../CustomProperty/CustomProperty 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Name | string | 必填 | 自定义属性的名称。 |
| Value | any | 必填 | 自定义属性的值。 |

## 示例

```javascript
/* 此示例向活动工作表添加标识符信息，并将其名称和值返回给用户*/
function test() {

    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    console.log(cusProperties.Name + "\t" + cusProperties.Value)

}
```
