---
title: "Hyperlinks (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Hyperlinks%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BXeMNmdn
---

# Hyperlinks (属性)

返回 **[Hyperlinks](./../../Hyperlinks/Hyperlinks 对象.html)** 集合，它代表区域的超链接。

## 示例

```javascript
/*此示例检查第一张工作表上的超链接是否含有“Kingsoft”字样。*/
function test() {
    let hyperlinks = Worksheets.Item(1).Hyperlinks
    for (let i = 1; i <= hyperlinks.Count; i++) {
        let h = hyperlinks.Item(i)
        if (h.Name.indexOf("Kingsoft") >= 0) {
            h.Follow()
        }
    }
}
```

```javascript
/*本示例为活动工作表 A3 单元格添加超链接。*/
function test() {
    ActiveSheet.Hyperlinks.Add(Range("A3"), "https://www.wps.cn/")
}
```
