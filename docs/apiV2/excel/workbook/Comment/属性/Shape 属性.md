---
title: "Shape 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/%E5%B1%9E%E6%80%A7/Shape%20%E5%B1%9E%E6%80%A7.html
---
# Shape (属性)

返回一个 **[Shape](./../../Shape/Shape 对象.html)** 对象，它代表连接到指定批注的形状。

## 说明

确保批注未被隐藏。在“ 审阅 ”选项卡上，选择“ 批注>显示所有批注”。

## 示例

```javascript
/*本示例选定活动工作表上的第二条批注。*/
function test() {
    Application.ActiveSheet.Comments.Item(2).Shape.Select()
}
```

```javascript
/*本示例显示活动工作表上第三条批注的形状的宽度。*/
function test() {
    console.log(Application.ActiveSheet.Comments.Item(3).Shape.Width)
}
```
