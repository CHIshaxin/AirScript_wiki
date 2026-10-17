---
title: "LogEst (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/LogEst%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: k_cZ2Ro-
---

# LogEst (方法)

在回归分析中，计算最符合数据的指数回归拟合曲线，并返回描述该曲线的数值数组。因为此函数返回数值数组，所以必须以数组公式的形式输入。

## 说明

此曲线的方程为：

y = b*m^x 或

y = (b*(m1^x1)*(m2^x2)*_)（如果有多个 x 值）

其中因变量 y 是自变量 x 的函数值。m 值是各指数 x 的底，而 b 值是常量值。注意：公式中的 y、x 和 m 均可以是向量，函数 LOGEST 所返回的数组为 {mn,mn-i,...,m1,b}。

- 如果数组 known_y's 在单独一列中，则 known_x's 的每一列被解释为一个独立的变量。
- 如果数组 known_y's 在单独一行中，则 known_x's 的每一行被解释为一个独立的变量。
- known_x's 数组可以包括一组或多组自变量。如果仅使用一个变量，那么只要 known_x's 和 known_y's 具有相同的维数，则它们可以是任何形状的区域。如果使用多个变量，则 known_y's 必须是向量（即具有一列高度或一行宽度的单元格区域）。
- 如果省略 known_x's，则假定该数组为 {1,2,3,...}，其大小与 known_y's 相同。
- 如果 const 为 TRUE 或被省略，b 将按正常方式计算。
- 如果 const 为 FALSE，则常量 b 将设为 1，而 m 的值满足公式 y = m^x。
- 如果 stats 为 TRUE，函数 LOGEST 将返回附加的回归统计值，因此返回的数组为 {mn,mn-1,...,m1,b;sen,sen-1,...,se1,seb;r 2,sey; F,df;ssreg,ssresid}。
- 如果 stats 为 FALSE 或省略，则函数 LOGSET 只返回系数 m 和常量 b。

有关附加回归统计值的详细信息，请参阅函数 LINEST。

- 通过数据绘出的图线越近似于指数曲线，则计算出来的曲线就越符合原来给定的数据。正如函数 LINEST 一样，函数 LOGEST 也返回一组描述数值相互关系的数值数组，但函数 LINEST 是用直线来拟合数据的，而函数 LOGEST 则以指数曲线来拟合数据。有关详细信息，请参阅函数 LINEST。
- 当仅有一个自变量 x 时，可直接用下面的公式计算出 y 轴截距 (b) 的值： Y 轴截距 (b)： INDEX(LOGEST(known_y's,known_x's),2) 可用 y=b*m＾x 公式来预测 y 的值，但是 ET 另外提供了可以预测因变量 y 值的 GROWTH 函数。有关详细信息，请参阅 GROWTH 函数。
- 对于返回结果为数组的公式，必须以数组公式的形式输入。
- 当输入一个数组常量（如 known_x's）作为参数时，用逗号来分隔同一行中的值，用分号来分隔不同的行。分隔符可能因**“控制面板”**中的**“区域和语言选项”**中区域设置的不同而有所不同。
- 应注意的一点是：如果由回归公式所预测的 y 值超出用来计算回归公式的 y 的取值区间，则该值可能无效。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Known_y's - 关系表达式 y = b*m^x 中一组已知的 y 值。 |
| Arg2 | any | 可选 | Known_x's - 关系表达式 y = b*m^x 中一组已知的 x 值，为可选参数。 |
| Arg3 | any | 可选 | Const - 一个逻辑值，用于指定是否强制常量 b 等于 1。 |
| Arg4 | any | 可选 | Stats - 一个逻辑值，指定是否返回附加回归统计值。 |

## 返回值

Variant

## 示例

```javascript
/*本示例为C1等单元格分别赋值，使用 LogEst 方法分别获取适合区域C1:C2和区域D1:D2数据的指数曲线，分别计算描述曲线的值数组，并分别赋值给A1等单元格。*/
function test() {
    Range("C1").Value2 = 100
    Range("C2").Value2 = 80
    Range("D1").Value2 = 5
    Range("D2").Value2 = 10
    Range("A1:B1").Value2 = WorksheetFunction.LogEst(Range("C1:C2"), Range("D1:D2"), true, true)
    Range("A2:B2").Value2 = WorksheetFunction.LogEst(Range("C1:C2"), Range("D1:D2"), false, false)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 LogEst 方法获取适合区域C1:C2数据的指数曲线，并计算描述曲线的值数组。*/
function test() {
    Range("C1").Value2 = 70.5
    Range("C2").Value2 = 100
    Range("D1").Value2 = 5
    Range("D2").Value2 = 6
    let arr = WorksheetFunction.LogEst(Range("C1:C2"), Range("D1:D2"))
    console.log(arr.toString())
}
```
