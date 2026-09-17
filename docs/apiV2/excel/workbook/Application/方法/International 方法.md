---
title: "International 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/International%20%E6%96%B9%E6%B3%95.html
---
# International (方法)

返回当前的国家/地区和国际设置信息。只读 **Variant** 类型。

## 说明

### 括号和大括号

| 索引 | 类型 | 含义 |
| --- | --- | --- |
| **xlLeftBrace** | String | 在字面数组中左大括号 ({) 的替代字符。 |
| **xlLeftBracket** | String | 在 R1C1-样式相对引用中左方括号 ([) 的替代字符。 |
| **xlLowerCaseColumnLetter** | String | 小写列字母。 |
| **xlLowerCaseRowLetter** | String | 小写行字母。 |
| **xlRightBrace** | String | 在字面数组中右大括号 (}) 的替代字符。 |
| **xlRightBracket** | String | 在 R1C1-样式引用中右方括号 (]) 的替代字符。 |
| **xlUpperCaseColumnLetter** | String | 大写列字母。 |
| **xlUpperCaseRowLetter** | String | 大写行字母（对于 R1C1-样式引用）。 |

### 国家/地区设置

| 索引 | 类型 | 含义 |
| --- | --- | --- |
| **xlCountryCode** | Long | ET 的国家/地区版本。 |
| **xlCountrySetting** | Long | Windows“控制面板”中的当前国家/地区设置。 |
| **xlGeneralFormatName** | String | “常规”数字格式名称。 |

### 货币

| 索引 | 类型 | 含义 |
| --- | --- | --- |
| **xlCurrencyBefore** | Boolean | 如果货币符号在货币值之前，则返回 **True**；如果货币符号在货币值之后，则返回 **False**。 |
| **xlCurrencyCode** | String | 货币符号。 |
| **xlCurrencyDigits** | Long | 货币格式中的小数位数。 |
| **xlCurrencyLeadingZeros** | Boolean | 如果显示零货币值的前导零，则返回 **True**。 |
| **xlCurrencyMinusSign** | Boolean | 如果对负数使用负号，则返回 **True**；如果使用括号，则返回 **False**。 |
| **xlCurrencyNegative** | Long | 负数货币值的货币格式： 0 = (symbolx) 或 (xsymbol) 1 = -symbolx 或 -xsymbol 2 = symbol-x 或 x-symbol 3 = symbolx- 或 xsymbol- 其中 symbol 为国家/地区的货币符号。请注意货币符号的位置由 **xlCurrencyBefore** 决定。 |
| **xlCurrencySpaceBefore** | Boolean | 如果在货币符号前面添加空格，则返回 **True**。 |
| **xlCurrencyTrailingZeros** | Boolean | 如果显示零货币值的尾部零，则返回 **True**。 |
| **xlNoncurrencyDigits** | Long | 非货币格式中所使用的十进制数字的个数。 |

### 日期和时间

| 索引 | 类型 | 含义 |
| --- | --- | --- |
| **xl24HourClock** | **Boolean** | 如果使用 24 小时制时间，则返回 **True**；如果使用 12 小时制时间，则返回 **False**。 |
| **xl4DigitYears** | **Boolean** | 如果使用四位年，则返回 **True**；如果使用两位年，则返回 **False**。 |
| **xlDateOrder** | **Long** | 日期元素的次序： 0 = 月-日-年 1 = 日-月-年 2 = 年-月-日 |
| **xlDateSeparator** | **String** | 日期分隔符 (**/**)。 |
| **xlDayCode** | **String** | 日符号 (d)。 |
| **xlDayLeadingZero** | **Boolean** | 如果在日期中显示前导零，则返回 **True**。 |
| **xlHourCode** | **String** | 小时符号 (h)。 |
| **xlMDY** | **Boolean** | 如果长日期显示中日期次序为月-日-年，则返回 **True**；如果次序为日-月-年，则返回 **False**。 |
| **xlMinuteCode** | **String** | 分钟符号 (m)。 |
| **xlMonthCode** | **String** | 月符号 (m)。 |
| **xlMonthLeadingZero** | **Boolean** | 如果以数字显示月份时显示月份中的前导零，则返回 **True**。 |
| **xlMonthNameChars** | **Long** | 为了向后兼容总是返回三个字符。月份名称的缩写从 Windows 中读取并且可以为任意长度。 |
| **xlSecondCode** | **String** | 秒符号 (s)。 |
| **xlTimeSeparator** | **String** | 时间分隔符 (😃。 |
| **xlTimeLeadingZero** | **Boolean** | 如果在时间表示中显示前导零，则该值为 **True**。 |
| **xlWeekdayNameChars** | **Long** | 为了向后兼容总是返回三个字符。星期名称的缩写从 Windows 中读取并且可以为任意长度。 |
| **xlYearCode** | **String** | 数字格式中的年符号 (y)。 |

### 度量系统

| 索引 | 类型 | 含义 |
| --- | --- | --- |
| **xlMetric** | **Boolean** | 如果使用米制度量系统，则返回 **True**；如果使用英制度量系统，则返回 **False**。 |
| **xlNonEnglishFunctions** | **Boolean** | 如果不以英文显示函数，则返回 **True**。 |

### 分隔符

| 索引 | 类型 | 含义 |
| --- | --- | --- |
| **xlAlternateArraySeparator** | **String** | 如果当前数组分隔符与小数分隔符相同时替代的数组项分隔符。 |
| **xlColumnSeparator** | **String** | 数组文字类型的列分隔符。 |
| **xlDecimalSeparator** | **String** | 小数分隔符。 |
| **xlListSeparator** | **String** | 列表分隔符。 |
| **xlRowSeparator** | **String** | 数组文字类型的行分隔符。 |
| **xlThousandsSeparator** | **String** | 零或千位分隔符。 |

上表显示的符号、分隔符和货币格式可能因所用语言和地理位置而异，而且可能不可用，这取决于选择或安装的语言支持（例如，美国英语）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 要返回的设置。可以是表中列出的 **[XlApplicationInternational](./../../枚举/XlApplicationInternational 枚举.html)** 常量之一。 |

## 示例

```javascript
/*本示例显示区域设置中的十进制分隔符。*/
function test() {
    console.log(`The decimal separator is ${Application.International(xlDecimalSeparator)}`)
}
```

```javascript
/*本示例显示区域设置中是否使用的是四位数年份。*/
function test() {
    console.log(Application.International(xl4DigitYears))
}
```
