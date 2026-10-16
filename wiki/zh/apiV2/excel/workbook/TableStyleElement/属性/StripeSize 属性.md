---
title: "StripeSize (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElement/%E5%B1%9E%E6%80%A7/StripeSize%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bt8BxMEg
---

# StripeSize (属性)

返回或设置条带的大小。可读/写 **Long** 类型。

## 说明

此属性不会应用于所有 **[TableStyleElement](./../../TableStyleElement/TableStyleElement 对象.html)** 对象。它只应用于 **xlColumnStripe1**、**xlColumnStripe2**、**xlRowStripe1** 和 **xlRowStripe2** 类型。

## 示例

```javascript
/*本示例显示活动工作表上第一张列表的列条纹1样式元素条带的大小。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlColumnStripe1).StripeSize)
}
```

```javascript
/*本示例将第一张工作表上第一张列表的行条纹2样式元素条带的大小设置为2。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    listObj.TableStyle.TableStyleElements.Item(xlRowStripe2).StripeSize = 2
}
```
