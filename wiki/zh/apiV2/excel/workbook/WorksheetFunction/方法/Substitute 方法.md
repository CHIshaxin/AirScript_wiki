---
title: "Substitute (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Substitute%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DJCdJKaf
---

# Substitute (方法)

在文本字符串中用 new_text 替代 old_text。如果需要在某一文本字符串中替换指定的文本，请使用函数 SUBSTITUTE；如果需要在某一文本字符串中替换指定位置处的任意文本，请使用函数 REPLACE。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | Text - 要替换其中字符的文本，或对含有要替换其中字符的文本的单元格的引用。 |
| Arg2 | string | 必填 | old_text - 要替换的文本。 |
| Arg3 | string | 必填 | New_text - 要用于替换 old_text 的文本。 |
| Arg4 | any | 可选 | Instance_num - 指定要用 new_text 替换 old_text 的第几个匹配项。如果指定了 instance_num，则只有 old_text 的该实例被替换；否则将用 new_text 替换 text 中出现的所有 old_text 的匹配项。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Substitute 方法在文本字符串中用新文本替代旧文本，并分别赋值给C4和C5单元格。*/
function test() {
    Range("C4").Value2 = WorksheetFunction.Substitute("你好，欢迎你的加入", "你的加入", "你")
    Range("C5").Value2 = WorksheetFunction.Substitute("你好，欢迎你的加入", "你", "大家", 2)
}
```

```javascript
/*本示例为A3单元格赋值，使用 Substitute 方法在文本字符串中用 “金山办公” 替代 “wps”。*/
function test() {
    Range("A3").Value2 = "这是wps的研发部门"
    console.log(WorksheetFunction.Substitute(Range("A3").Value2, "wps", "金山办公"))
}
```
