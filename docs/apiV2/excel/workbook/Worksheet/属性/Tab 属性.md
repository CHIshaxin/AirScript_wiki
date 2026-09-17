---
title: "Tab 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Tab%20%E5%B1%9E%E6%80%A7.html
---
# Tab (属性)

返回一个工作表的 **Tab** 对象。

## 示例

```javascript
/*此示例确定第一张工作表选项卡的颜色索引是否设置为无，并通知用户。*/
function test() {
    // Determine if color index of 1st tab is set to none.
    if (Worksheets.Item(1).Tab.ColorIndex == xlColorIndexNone) {
        console.log("The color index is set to none for the 1st worksheet tab.")
    } else {
        console.log("The color index for the tab of the 1st worksheet is not set none.")
    }
}
```

```javascript
/*此示例显示活动工作表选项卡的颜色。*/
function test() {
    console.log(Application.ActiveSheet.Tab.Color)
}
```
