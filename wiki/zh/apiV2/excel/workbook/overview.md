---
title: "工作表 API"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/overview.html
updated: 1789700916000
hash: CtI91TM_
---

# 工作表 API

工作表 API 提供了对 Excel 工作表的完整控制能力，包括单元格操作、格式设置、图表创建、数据筛选等功能。

## 核心对象

- **Application** - 应用程序对象，提供全局方法和属性
- **Workbook** - 工作簿对象，管理工作表集合
- **Worksheet** - 工作表对象，操作单个工作表
- **Range** - 区域对象，操作单元格区域
- **Chart** - 图表对象，创建和操作图表

## 快速开始

```javascript
// 获取当前活动工作表
const sheet = ActiveSheet;

// 读取单元格值
const value = sheet.Cells(1, 1).Value2;

// 设置单元格值
sheet.Cells(1, 1).Value2 = "Hello World";

// 获取区域
const range = sheet.Range("A1:B10");
```

请从左侧导航选择具体的 API 对象查看详细文档。
