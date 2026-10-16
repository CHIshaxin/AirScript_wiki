---
title: "什么是AirScript 2.0"
source: https://airsheet.wps.cn/docs/apiV2/overview.html
updated: 1789700916000
hash: C0pyKoy6
---

# 什么是AirScript 2.0

AirScript 1.0脚本能够满足常规的脚本操作文档需求。但在API的丰富度、执行性能、JavaScript高级语法支持上仍然存在一些限制。为了突破这些局限，我们推出了AirScript 2.0版本。

AirScript 2.0脚本采用全新的架构设计，在功能和性能上带来大幅升级，API完全兼容WPS JS宏，执行速度更快，并全面支持现代JavaScript语法。

尽管AirScript 1.0大部分API可以在AirScript 2.0兼容运行，但仍然存在一些API定义存在冲突的问题导致无法实现完全兼容，因此需要通过`1.0` `2.0`来标识脚本的运行时环境版本，AirScript 2.0是AirScript的一个大版本升级，部分API用法与AirScript 1.0无法兼容。

# AirScript 2.0有哪些亮点

### ✔️完全兼容WPS JS宏

AirScript 2.0完全兼容WPS JS宏，提供了2000+API（比1.0多4倍），极大地扩展了您的开发能力。无论您是进行简单的自动化任务，还是复杂的脚本开发，AirScript 2.0都能满足您的需求。无需重新学习新的API定义，即可在WPS端和在线表格使用同样的API开发脚本。

### 🚀更快的执行速度

我们对AirScript 2.0的逻辑执行和API调用速度进行了优化，使您的脚本运行更加高效，响应更加迅速。对于需要频繁调用API的大型脚本项目，这种性能提升尤为显著。如：循环对1000个单元格文本去除首尾空格并写回，2.0比1.0速度快1倍。

> 可以分别新建1.0和2.0脚本，执行以下代码体验2.0的性能提升

```javascript
// 对当前激活sheet第一列前1k个格子去除首尾空格，并打印脚本执行耗时
const start = Date.now();

const sheet = ActiveSheet;
for (let i = 1; i <= 1000; i++) {
  const cell = sheet.Cells(i, 1);
  cell.Value2 = cell.Text.trim();
}

console.log(Date.now() - start);
```

### 🌐全面支持现代 JavaScript 语法

AirScript 2.0全面支持现代 JavaScript 语法，如：await/async、Promise、class等语法（1.0不支持），让您的开发体验更加现代和高效。特别对于复杂的脚本项目，代码编写维护更方便。

# AirScript 2.0支持哪些场景

AirScript 2.0设计上能够满足所有AirScript 1.0具备的能力，目前一期开放支持的是工作表场景，数据表的API待开放。若您的脚本需要处理数据表场景，您可以先继续使用AirScript 1.0脚本，后续2.0我们将逐步开放支持数据表的API。

> 当前AirScript 2.0处于beta版，部分API能力将在后续迭代中持续补齐。

|  | AirScript 1.0 | AirScript 2.0 |
|---|---|---|
| 工作表 | ✅API（500+） | ✅API（2000+） |
| 数据表 | ✅API | 🏗️API（待开放） |
| 高级服务 | ✅网络API<br>✅云文档API<br>✅邮件API<br>✅数据库API | ✅网络API<br>✅云文档API<br>🏗️邮件API（待开放）<br>🏗️数据库API（待开放） |

# 如何使用AirScript 2.0

> 多维表格由于涉及数据表支持，待开放

### 如何新建2.0脚本

1.打开AirScript脚本编辑器

`效率`->`高级开发`->`AirScript脚本编辑器`
