---
title: "Text (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/%E6%96%B9%E6%B3%95/Text%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CQkD8hSc
---

# Text (方法)

设置批注文本。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Text | any | 可选 | 要添加的文本。 |
| Start | any | 可选 | 所添加文本的起始位置（字符数）。如果省略此参数，则删除批注中的所有现有文字。 |
| Overwrite | any | 可选 | 如果为 True，则覆盖现有文本。默认值是 False（插入文本）。 |

## 返回值

String

## 示例

```javascript
/*本示例将 A2 单元格的批注文本设置为"这是一个批注"。*/
function test() {
    Application.Range("A2").Comment.Text("这是一个批注")
}
```

```javascript
/*本示例在 A4 单元格的批注内容中第三个字符开始插入“hello”，并显示结果。*/
function test() {
    console.log(Application.Range("A4").Comment.Text("hello", 3))
}
```
