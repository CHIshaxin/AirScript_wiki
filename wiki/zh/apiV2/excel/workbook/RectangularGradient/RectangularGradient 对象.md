---
title: "RectangularGradient (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/RectangularGradient/RectangularGradient%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B14p-G8z
---

# RectangularGradient (对象)

**[RectangularGradient](./../RectangularGradient/RectangularGradient 对象.html)** 对象沿特定角度以线性方式在一系列颜色间转换。

## 说明

| 使用 RectangularGradient 对象时，请考虑以下几点： |
| --- |

- 试图访问不具有现有渐变填充的 **[Interior](./../Interior/Interior 对象.html)** 对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意 `Interior.Pattern` 属性。
- 如果将 Interior.Pattern从渐变类型更改为非渐变类型，Gradient 对象将采用默认值进行填充。
