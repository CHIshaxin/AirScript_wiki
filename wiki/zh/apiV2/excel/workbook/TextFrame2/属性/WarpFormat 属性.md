---
title: "WarpFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/WarpFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DaEi7Fl4
---

# WarpFormat (属性)

返回或设置指定文本框的弯曲类型。可读/写 **MsoWarpFormat**。

## 示例

```javascript
/*本示例显示活动工作表中第二个形状的文本框的弯曲类型是否为msoWarpFormat11。*/
function test() {
    let textFrame = ActiveSheet.Shapes.Item(2).TextFrame2
    console.log(textFrame.WarpFormat == msoWarpFormat11)
}
```

```javascript
/*本示例将第一张工作表中第一个形状的文本框的弯曲类型设置为msoWarpFormat10。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame2.WarpFormat = msoWarpFormat10
}
```
