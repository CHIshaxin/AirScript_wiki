---
title: "AlertStyle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/AlertStyle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dsp3pLyY
---

# AlertStyle (属性)

返回有效性检验警告样式。 **[XlDVAlertStyle](./../../枚举/XlDVAlertStyle 枚举.html)** 类型，只读。

## 说明

使用 **Add** 方法设置某一区域的警告样式。如果该区域已有数据有效性检验，则可用 **Modify** 方法更改警告样式。

## 示例

```javascript
/*本示例显示单元格 E5 的数据有效性检验的警告样式是否为xlValidAlertInformation。*/
function test() {
    console.log(Range("E5").Validation.AlertStyle == xlValidAlertInformation)
}
```

```javascript
/*本示例显示活动工作表上 A1 单元格的数据有效性检验的警告样式。*/
function test() {
    let validation = ActiveSheet.Range("A1").Validation
    switch (validation.AlertStyle) {
        case xlValidAlertInformation:
            console.log("警告样式为信息图标")
            break
        case xlValidAlertStop:
            console.log("警告样式为停止图标")
            break
        case xlValidAlertWarning:
            console.log("警告样式为警告图标")
            break
    }
}
```
