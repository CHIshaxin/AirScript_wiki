---
title: "CustomProperty 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CustomProperty/CustomProperty%20%E5%AF%B9%E8%B1%A1.html
---
# CustomProperty (对象)

代表标识符信息。标识符信息可用于 XML 的元数据。

## 说明

使用 **Add** 方法或 **[CustomProperties](./../CustomProperties/CustomProperties 对象.html)** 集合的 **Item** 属性可返回 **[CustomProperty](./../CustomProperty/CustomProperty 对象.html)** 对象。

返回 **[CustomProperty](./../CustomProperty/CustomProperty 对象.html)** 对象后，可在 **Add** 方法中使用 **[CustomProperties](./../CustomProperties/CustomProperties 对象.html)** 属性向工作表中添加元数据。

在本示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

```javascript
function test() {
    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    console.log(cusProperties.Name + "\t" + cusProperties.Value)
}
```
