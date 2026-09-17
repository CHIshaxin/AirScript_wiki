---
title: "CustomProperties 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CustomProperties/CustomProperties%20%E5%AF%B9%E8%B1%A1.html
---
# CustomProperties (对象)

由代表附加信息的 **[CustomProperty](./../CustomProperty/CustomProperty 对象.html)** 对象组成的集合，这些信息可用作 XML 的元数据。

## 说明

使用 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象的 **[CustomProperties](./../CustomProperties/CustomProperties 对象.html)** 属性返回 **[CustomProperties](./../CustomProperties/CustomProperties 对象.html)** 集合。

返回 **[CustomProperties](./../CustomProperties/CustomProperties 对象.html)** 集合后，可根据选择向工作表和智能标记中添加元数据。

若要向工作表添加元数据，请在 **Add** 方法中使用 **[CustomProperties](./../CustomProperties/CustomProperties 对象.html)** 属性。

下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

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
