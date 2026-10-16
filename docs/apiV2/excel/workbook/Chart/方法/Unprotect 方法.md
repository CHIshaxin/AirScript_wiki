---
title: "Unprotect 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/Unprotect%20%E6%96%B9%E6%B3%95.html
---
# Unprotect (方法)

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

## 说明

如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Password | any | 可选 | 指定用于解除图表保护的密码，此密码是区分大小写的。如果图表不设密码保护，则忽略该参数。 |

## 示例

```javascript
/*本示例判断如果图表工作表 Chart1 的内容受保护，则取消对该工作表的保护。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    if (chart.ProtectContents) {
        chart.Unprotect()
        console.log("已取消对工作表的保护")
    }
}
```

```javascript
/*本示例取消对已设置密码保护的图表工作表 Chart2 的保护。*/
function test() {
    let chart = Application.Charts.Item("Chart2")
    chart.Unprotect("Y6dh!et5")
}
```
