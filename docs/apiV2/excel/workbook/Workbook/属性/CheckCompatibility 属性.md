---
title: "CheckCompatibility 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/CheckCompatibility%20%E5%B1%9E%E6%80%A7.html
---
# CheckCompatibility (属性)

控制在保存工作簿时是否自动运行兼容性检查器。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例演示在保存工作簿时自动运行兼容性检查器。*/
function test() {
    Workbooks.Item(1).CheckCompatibility = true
}
```

```javascript
/*本示例演示 CheckCompatibility 属性的数据类型。*/
function test() {
    console.log(typeof ActiveWorkbook.CheckCompatibility)
}
```
