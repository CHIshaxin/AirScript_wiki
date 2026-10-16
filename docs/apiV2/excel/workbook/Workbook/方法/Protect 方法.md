---
title: "Protect 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E6%96%B9%E6%B3%95/Protect%20%E6%96%B9%E6%B3%95.html
---
# Protect (方法)

保护工作簿使其不被修改。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Password | any | 可选 | 一个字符串，该字符串为工作表或工作簿指定区分大小写的密码。如果省略此参数，不用密码就可以取消对工作表或工作簿的保护。否则，必须指定密码才能取消对工作表或工作簿的保护。如果忘记了密码，就无法取消对工作表或工作簿的保护。 |
| Structure | any | 可选 | 如果为 **True**，则保护工作簿结构（工作表的相对位置）。默认值是 **False**。 |
| Windows | any | 可选 | 如果为 **True**，则保护工作簿窗口。如果省略此参数，则窗口不受保护。 |

## 示例

```javascript
/*本示例演示对活动工作簿开启保护工作簿，并指定密码。*/
function test() {
    Application.ActiveWorkbook.Protect("Y6dh!et5")
}
```

```javascript
/*本示例演示保护第一个工作簿结构和窗口。*/
function test() {
    Workbooks.Item(1).Protect(undefined, true, true)
}
```
