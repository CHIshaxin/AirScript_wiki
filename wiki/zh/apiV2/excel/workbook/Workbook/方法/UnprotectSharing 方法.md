---
title: "UnprotectSharing (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E6%96%B9%E6%B3%95/UnprotectSharing%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: hxz12Dgb
---

# UnprotectSharing (方法)

关闭共享保护功能并保存工作簿。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| SharingPassword | any | 可选 | 工作簿密码。 |

## 示例

```javascript
/*本示例关闭共享保护，并保存活动工作簿。*/
function test() {
    ActiveWorkbook.UnprotectSharing()
}
```

```javascript
/*此示例演示保存第一张工作簿并设置共享保护，再用密码取消对第一张工作簿的共享保护，过程中会提示保存工作簿。*/
function test() {
    Workbooks.Item(1).ProtectSharing(undefined, undefined, undefined, undefined, undefined, "Y6dh!et5")
    Workbooks.Item(1).UnprotectSharing("Y6dh!et5")
}
```
