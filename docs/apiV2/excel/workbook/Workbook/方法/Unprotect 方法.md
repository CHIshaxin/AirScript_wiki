---
title: "Unprotect 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E6%96%B9%E6%B3%95/Unprotect%20%E6%96%B9%E6%B3%95.html
---
# Unprotect (方法)

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

## 说明

如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Password | any | 可选 | 一个字符串，指定用于解除工作表或工作簿保护的密码，此密码是区分大小写的。如果工作表或工作簿不设密码保护，则省略此参数。如果对工作表省略此参数，而该工作表又设有密码保护，ET 将提示您输入密码。如果对工作簿省略此参数，而该工作簿又设有密码保护，则该方法将失效。 |

## 示例

```javascript
/*此示例取消当前活动工作簿的保护。*/
function test() {
    ActiveWorkbook.Unprotect()
}
```

```javascript
/*此示例演示用密码开启保护工作簿并提示，再用密码取消对活动工作簿的保护。*/
function test() {
    ActiveWorkbook.Protect("Y6dh!et5")
    console.log("当前已开启保护工作簿，点击确定后会取消保护工作簿")
    ActiveWorkbook.Unprotect("Y6dh!et5")
}
```
