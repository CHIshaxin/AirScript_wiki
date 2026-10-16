---
title: "IMEMode 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/IMEMode%20%E5%B1%9E%E6%80%A7.html
---
# IMEMode (属性)

返回或设置日文输入规则的说明。可以是下表列出的 **[XlIMEMode](./../../枚举/XlIMEMode 枚举.html)** 常量之一。**Long** 类型，可读写。

## 说明

| 常量 | 说明 |
| --- | --- |
| **xlIMEModeAlpha** | 半角字母数字 |
| **xlIMEModeAlphaFull** | 全角字母数字 |
| **xlIMEModeDisable** | 禁用 |
| **xlIMEModeHiragana** | 平假名 |
| **xlIMEModeKatakana** | 片假名 |
| **xlIMEModeKatakanaHalf** | 片假名（半角） |
| **xlIMEModeNoControl** | 无控制 |
| **xlIMEModeOff** | 关闭（英文模式） |
| **xlIMEModeOn** | 开启 |

只有安装和选择了日文语言支持时，才可设置该属性。

## 示例

```javascript
/*本示例设置单元格 E5 的数据有效性的输入规则。*/
function test() {
    let validation = Application.Range("E5").Validation
    validation.Add(xlValidateWholeNumber, xlValidAlertStop, xlBetween, "5", "10")
    validation.InputTitle = " seisuuchi seisuuchi seisuuchi"
    validation.ErrorTitle = " seisuuchi seisuuchi seisuuchi"
    validation.InputMessage = "5  kara kara 10  否 seisuu seisuu o nyuuryoku nyuuryoku shite shite kudasai kudasai kudasai kudasai ."
    validation.ErrorMessage = " nyuuryoku nyuuryoku dekirunowa dekirunowa dekirunowa dekirunowa dekirunowa 5  kara kara 10  madeno madeno madeno ne desu desu ."
    validation.IMEMode = xlIMEModeAlpha
}
```

```javascript
/*本示例显示单元格 E1 的数据有效性的输入规则是否为xlIMEModeNoControl。*/
function test() {
    console.log(Range("E1").Validation.IMEMode == xlIMEModeNoControl)
}
```
