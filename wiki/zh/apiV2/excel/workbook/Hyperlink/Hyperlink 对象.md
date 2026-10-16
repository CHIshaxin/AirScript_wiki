---
title: "Hyperlink (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/Hyperlink%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BugXaTTj
---

# Hyperlink (对象)

代表一个超链接。

## 说明

**[Hyperlink](./../Hyperlink/Hyperlink 对象.html)** 对象是 **[Hyperlinks](./../Hyperlinks/Hyperlinks 对象.html)** 集合的成员。

使用 **Hyperlink** 属性可返回某个形状的超链接（一个形状只能有一个超链接）。

## 示例

```javascript
/*本示例加载附于第一个形状（第一张工作表中）的超链接*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.Follow(true)
}
```

区域或工作表可以有多个超链接。使用 **Hyperlinks**(*index*) 可返回单个 **[Hyperlink](./../Hyperlink/Hyperlink 对象.html)** 对象，其中 *index* 是超链接编号。

```javascript
/*本示例打开活动工作表中 A1:B2 区域中的第二个超链接*/
function test() {
    Application.Worksheets.Item(1).Range("A1:B2").Hyperlinks.Item(2).Follow()
}
```
