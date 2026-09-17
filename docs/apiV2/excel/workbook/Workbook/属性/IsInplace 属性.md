---
title: "IsInplace 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/IsInplace%20%E5%B1%9E%E6%80%A7.html
---
# IsInplace (属性)

如果指定工作簿正在适当位置上进行编辑，则该值为 **True**。如果该工作簿已在 ET 中打开，并准备进行编辑，则该值为 **False**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例表示工作簿是在适当位置上打开并正在编辑还是仅在 ET 中打开。*/
function test() {
    if (ThisWorkbook.IsInPlace) {
        console.log("Editing in place")
    } else {
        console.log("Editing in ET")
    }
}
```

```javascript
/*本示例演示IsInplace属性的数据类型。*/
function test() {
    console.log("IsInplace属性的数据类型为：" + typeof ActiveWorkbook.IsInplace)
}
```
