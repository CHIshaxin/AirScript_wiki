---
title: "Replace 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Replace%20%E6%96%B9%E6%B3%95.html
---
# Replace (方法)

返回 **Boolean**，它表示指定区域内单元格中的字符。使用此方法并不会更改选定区域或活动单元格。

## 说明

每次使用此方法后，参数 *LookAt*、*SearchOrder*、*MatchCase* 和 *MatchByte* 的设置都将被保存。如果下次调用此方法时不指定这些参数的值，就使用保存的值。设置这些参数将更改 **“查找”** 对话框中的设置，如果省略这些参数，更改 **“查找”** 对话框中的设置将更改使用的保存值。要避免出现这一问题，每次使用此方法时请明确设置这些参数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| What | any | 必填 | ET 要搜索的字符串。 |
| Replacement | any | 必填 | 替换字符串。 |
| LookAt | any | 可选 | 可为以下 **[XlLookAt](./../../枚举/XlLookAt 枚举.html)** 常量之一：**xlWhole** 或 **xlPart**。 |
| SearchOrder | any | 可选 | 可为以下 **[XlSearchOrder](./../../枚举/XlSearchOrder 枚举.html)** 常量之一：**xlByRows** 或 **xlByColumns**。 |
| MatchCase | any | 可选 | 如果为 **True**，则搜索区分大小写。 |
| MatchByte | any | 可选 | 只有在 ET 中选择或安装了双字节语言时，才能使用此参数。如果为 **True**，则双字节字符只与双字节字符匹配。如果为 **False**，则双字节字符可与其对等的单字节字符匹配。 |
| SearchFormat | any | 可选 | 该方法的搜索格式。 |
| ReplaceFormat | any | 可选 | 该方法的替换格式。 |

## 返回值

Boolean

## 示例

```javascript
/*本示例将每个“SIN”三角函数都替换为“COS”函数。替换所进行的区域为工作表 Sheet1 中的 A 列。*/
function test() {
    Worksheets.Item("Sheet1").Columns.Item("A").Replace("SIN", "COS", undefined, xlByColumns, true)
}
```

```javascript
/*本示例将 C3 单元格的值设置为 testC3，然后将 C3 单元格中的 test 替换为 range*/
function test() {
    Range("C3").Value2 = "testC3"
    Range("C3").Replace("test", "range")
}
```
