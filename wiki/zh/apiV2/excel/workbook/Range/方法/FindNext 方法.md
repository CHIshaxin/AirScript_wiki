---
title: "FindNext (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/FindNext%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: iI18CizK
---

# FindNext (方法)

继续由 **Find** 方法开始的搜索。查找匹配相同条件的下一个单元格，并返回表示该单元格的 **[Range](./../../Range/Range 对象.html)** 对象。该操作不影响选定内容和活动单元格。

## 说明

当查找到指定查找区域的末尾时，本方法将环绕至区域的开始继续搜索。发生环绕后，为停止查找，可保存第一次找到的单元格地址，然后测试下一个查找到的单元格地址是否与其相同。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| After | any | 可选 | 指定一个单元格，查找将从该单元格之后开始。此单元格对应于从用户界面搜索时的活动单元格位置。注意，After 必须是查找区域中的单个单元格。注意，搜索是从该单元格之后开始的；直到本方法环绕到此单元格时，才检测其内容。如果未指定本参数，查找将从区域的左上角单元格之后开始。 |

## 示例

```javascript
/*本示例在单元格区域 A1:A500 中查找值为 2 的单元格，并将这些单元格的值变为 5。*/
function test() {
    let add = Application.Worksheets.Item(1).Range("A1:A500")
    let c = add.Find(2, undefined, xlValues)
    if (c != null) {
        let firstAddress = c.Address()
        do {
            c.Value2 = 5
            c = add.FindNext(c)
        } while (c != null && c.Address() != firstAddress)
    }
}
```

```javascript
/*本示例搜索 A1:E15 区域中第二个含有 12 的单元格，并将它的值设置为 200*/
function test() {
    let c = Range("A1:E15").Find(12, undefined, xlValues)
    c = Range("A1:E15").FindNext(c)
    c.Value2 = 200
}
```
