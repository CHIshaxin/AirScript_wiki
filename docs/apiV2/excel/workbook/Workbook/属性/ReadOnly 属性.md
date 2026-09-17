---
title: "ReadOnly 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/ReadOnly%20%E5%B1%9E%E6%80%A7.html
---
# ReadOnly (属性)

如果对象以只读方式打开，则返回 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*此示例演示如果当前活动工作簿是只读的，则将此活动工作簿另存为 Newfile.xls。*/
function test() {
    if (Application.ActiveWorkbook.ReadOnly == true) {
        Application.ActiveWorkbook.SaveAs("NEWFILE.XLS")
    }
}
```

```javascript
/*此示例演示当前活动工作簿 ReadOnly 属性的数据类型。*/
function test() {
    console.log(typeof (Application.ActiveWorkbook.ReadOnly))
}
```
