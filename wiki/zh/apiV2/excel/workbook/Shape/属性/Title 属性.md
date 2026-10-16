---
title: "Title (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Title%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DuEsxbML
---

# Title (属性)

返回或设置与指定形状关联的可选文本的标题。 读/写。

## 说明

若要设置形状的可选文字的描述性文本字符串，请使用 **AlternativeText** 属性。

## 返回值

String

## 示例

```javascript
/*本示例显示活动工作表中第二个形状的标题。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.Title)
}
```

```javascript
/*本示例将第一张工作表中第一个形状的标题设置为"圆角矩形"。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(1)
    shape.Title = "圆角矩形"
}
```
