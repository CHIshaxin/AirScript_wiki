---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Names/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: wqeJxsMf
---

# Add (方法)

为单元格区域定义新名称。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Name | any | 可选 | 如果未指定 NameLocal 参数，则指定要用作名称的英文文本。名称不能包括空格，并且不能设置为单元格引用的格式。 |
| RefersTo | any | 可选 | 如果未指定 RefersToLocal、RefersToR1C1 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 A1 格式表示法以英文表示）。 注释 如果引用不存在，则返回 null。 |
| Visible | any | 可选 | True 指定将名称定义为可见。False 指定将名称定义为隐藏。已隐藏的名称不会在“定义名称”、“粘贴名称”或“转到”对话框中显示。默认值为 True。 |
| MacroType | any | 可选 | 由以下值之一确定的宏类型： 1 - 用户定义函数（Function 过程） 2 - 宏（Sub 过程） 3 或省略 - 无（该名称不引用用户定义函数或宏） |
| ShortcutKey | any | 可选 | 指定宏的快捷键。必须是单个字母，例如“z”或“Z”。仅适用于命令宏。 |
| Category | any | 可选 | 如果 MacroType 参数等于 1 或 2，则此参数为宏或函数的分类。该分类在“函数向导”中使用。可以用数字（从 1 开始）或名称（以英文指定）引用现有的分类。如果指定的分类不存在，ET 将创建新分类。 |
| NameLocal | any | 可选 | 如果未指定 Name 参数，则指定要用作名称的本地化的文本。名称不能包括空格，并且不能设置为单元格引用的格式。 |
| RefersToLocal | any | 可选 | 如果未指定 RefersTo、RefersToR1C1 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 A1 格式表示法以本地化的文本表示）。 |
| CategoryLocal | any | 可选 | 如果未指定 Category 参数，则指定标识自定义函数分类的本地化的文本。 |
| RefersToR1C1 | any | 可选 | 如果未指定 RefersTo、RefersToLocal 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 R1C1 格式表示法以英文表示）。 |
| RefersToR1C1Local | any | 可选 | 如果未指定 RefersTo、RefersToLocal 和 RefersToR1C1 参数，则说明名称引用的内容（使用 R1C1 格式表示法以本地化的文本表示）。 |

## 示例

```javascript
/*此示例为活动工作簿中 Sheet1 上的区域 A1:D3 定义一个新名称，如果工作表 Sheet1 不存在，则返回 null。*/
function test() {
    Application.ActiveWorkbook.Names.Add("tempRange", "=Sheet1!$A$1:$D$3")
}
```

```javascript
/*此示例在单元格C1中定义一个新名称并隐藏。*/
function test() {
    Application.Names.Add("test3", "=Sheet1!$C$1", false)
}
```
