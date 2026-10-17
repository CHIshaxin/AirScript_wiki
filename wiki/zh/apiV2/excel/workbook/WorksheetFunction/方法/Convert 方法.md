---
title: "Convert (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Convert%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C9CamGM-
---

# Convert (方法)

将数字从一个度量系统转换为另一个度量系统。例如，Convert 可以将一个以“英里”为单位的距离表转换成一个以“公里”为单位的距离表。

## 说明

| 重量和质量 | From_unit 或 to_unit |
| --- | --- |
| 克 | "g" |
| 斯勒格 | "sg" |
| 磅（常衡制） | "lbm" |
| U（原子质量单位） | "u" |
| 盎司（常衡制） | "ozm" |

| 距离 | From_unit 或 to_unit |
| --- | --- |
| 米 | "m" |
| 法定英里 | "mi" |
| 海里 | "Nmi" |
| 英寸 | "in" |
| 英尺 | "ft" |
| 码 | "yd" |
| 埃 | "ang" |
| 皮卡（1/72 英寸） | "Pica" |

| 时间 | From_unit 或 to_unit |
| --- | --- |
| 年 | "yr" |
| 日 | "day" |
| 时 | "hr" |
| 分 | "mn" |
| 秒 | "sec" |

| 压强 | From_unit 或 to_unit |
| --- | --- |
| 帕斯卡 | "Pa"（或 "p"） |
| 大气压 | "atm"（或 "at"） |
| 毫米汞柱 | "mmHg" |

| 力 | From_unit 或 to_unit |
| --- | --- |
| 牛顿 | "N" |
| 达因 | "dyn"（或 "dy"） |
| 磅力 | "lbf" |

| 能量 | From_unit 或 to_unit |
| --- | --- |
| 焦耳 | "J" |
| 尔格 | "e" |
| 热力学卡 | "c" |
| IT 卡 | "cal" |
| 电子伏 | "eV"（或 "ev"） |
| 马力-小时 | "HPh"（或 "hh"） |
| 瓦特-小时 | "Wh"（或 "wh"） |
| 英尺磅 | "flb" |
| BTU | "BTU"（或 "btu"） |

| 功率 | From_unit 或 to_unit |
| --- | --- |
| 马力 | "HP"（或 "h"） |
| 瓦特 | "W"（或 "w"） |

| 磁 | From_unit 或 to_unit |
| --- | --- |
| 特斯拉 | "T" |
| 高斯 | "ga" |

| 温度 | From_unit 或 to_unit |
| --- | --- |
| 摄氏度 | "C"（或 "cel"） |
| 华氏度 | "F"（或 "fah"） |
| 开尔文 | "K"（或 "kel"） |

| 液体度量 | From_unit 或 to_unit |
| --- | --- |
| 茶匙 | "tsp" |
| 汤匙 | "tbs" |
| 液量盎司 | "oz" |
| 杯 | "cup" |
| U.S. 品脱 | "pt"（或 "us_pt"） |
| U.K. 品脱 | "uk_pt" |
| 夸脱 | "qt" |
| 加仑 | "gal" |
| 升 | "l"（或 "lt"） |

下列缩写的单位前缀可以加在任何公制单位 from_unit 或 to_unit 之前。

| 前缀 | 乘子 | 缩写 |
| --- | --- | --- |
| exa | 1E+18 | "E" |
| peta | 1E+15 | "P" |
| tera | 1E+12 | "T" |
| giga | 1E+09 | "G" |
| mega | 1E+06 | "M" |
| kilo | 1E+03 | "k" |
| hecto | 1E+02 | "h" |
| dekao | 1E+01 | "e" |
| deci | 1E-01 | "d" |
| centi | 1E-02 | "c" |
| milli | 1E-03 | "m" |
| micro | 1E-06 | "u" |
| nano | 1E-09 | "n" |
| pico | 1E-12 | "p" |
| femto | 1E-15 | "f" |
| atto | 1E-18 | "a" |

- 如果输入数据类型不正确，则 Convert 将生成一个错误。
- 如果单位不存在，则 Convert 将生成一个错误。
- 如果单位不支持缩写单位前缀，则 Convert 将生成一个错误。
- 如果单位属于不同的组，则 Convert 将生成一个错误。
- 单位名称和前缀区分大小写。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 要转换的以 from_units 为单位的值。 |
| Arg2 | any | 必填 | 数字的单位。 |
| Arg3 | any | 必填 | 结果的单位。Convert 接受将以下文本值（引号中）用作 from_unit 和 to_unit，请参阅下面的“说明”部分列出的内容。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用Convert方法将数字从一种度量系统转换为另一种度量系统，并将结果赋值于D3等单元格。*/
function test() {
    Range("D3").Value2 = Application.WorksheetFunction.Convert(10, "C", "F")
    Range("D4").Value2 = Application.WorksheetFunction.Convert(6, "gal", "l")
    Range("D5").Value2 = Application.WorksheetFunction.Convert(15, "mi", "km")
}
```

```javascript
/*本示例为A2等单元格分别赋值，并使用Convert方法将数字从一种度量系统转换为另一种度量系统。*/
function test() {
    Range("A2").Value2 = 20
    Range("A3").Value2 = 30
    Range("A4").Value2 = 40
    let convert1 = Application.WorksheetFunction.Convert(Range("A2").Value2, "km", "mi")
    let convert2 = Application.WorksheetFunction.Convert(Range("A3").Value2, "in", "ft")
    let convert3 = Application.WorksheetFunction.Convert(Range("A4").Value2, "cm", "in")
    console.log(convert1)
    console.log(convert2)
    console.log(convert3)
}
```
