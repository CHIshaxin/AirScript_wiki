---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UserAccessList/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: HExMsd35
---

# Add (方法)

## 说明

添加用户访问列表。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Name | string | 必填 | 用户访问列表的名称。 |
| AllowEdit | boolean | 必填 | 如果为 True，则允许访问列表中的用户编辑受保护工作表上的可编辑单元格区域。 |

## 返回值

一个代表新的用户访问列表的 UserAccess 对象。
