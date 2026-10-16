---
title: "应用场景"
source: https://airsheet.wps.cn/docs/apitoken/scenario.html
updated: 1789700916000
hash: D5r6Czz-
---

# 应用场景

在脚本令牌的加持下，智能表格的强大能力得到完美释放，开发者能以一种更为高效和准确的方式执行任务。无论是网页爬取、数据分析，还是自动化流程，我都可以借助脚本令牌来完成。

如下为我们根据实际场景写的一些示例和说明，希望能给开发者一定的启发。

## 1. 私密信息查询

现代社会越来越重视个人的隐私。这种趋势在很多方面都有所体现。

首先，在教育领域，许多学校开始加强对学生在校期间的信息保护，禁止将学生的个人信息出售或分享给他人。这包括学生的**家庭信息、教育记录、考试成绩**等。

其次，在医疗保健领域，病人的隐私保护成为了一个重要的问题。医生和其他医疗工作者需要遵守严格的隐私规定，确保**病人的个人信息**不会被泄露。

此外，在社交媒体领域，许多平台也开始加强对用户信息的保护。他们采取了更严格的数据安全措施，以确保用户的数据不会被泄露或滥用。

如下，我们将展示一个学生成绩查询的在线示例，你可以在线体验利用脚本令牌实现的私密信息查询的功能。现有学生成绩表如下图所示：

上述在线示例的底层实现逻辑就是利用了**脚本令牌执行 AirScript 脚本**，在表格内查询到数据后返回给界面显示，脚本示例代码如下：

```js
// 获取学生的学号
const student_id = '1'
// 获取学生的姓名
const student_name = '金小妹'
// 获取已使用区域的行数
const rowCount = Application.ActiveSheet.UsedRange.Rows.Count

// 获取匹配的学生数据
const data = []
for (let i = 1; i <= rowCount; i++) {
  const studentId = Application.ActiveSheet.Range(`A${i}`).Value2
  const studentName = Application.ActiveSheet.Range(`B${i}`).Value2

  if (studentId === student_id && studentName === student_name) {
    const sex = Application.ActiveSheet.Range(`C${i}`).Value2
    const className = Application.ActiveSheet.Range(`D${i}`).Value2
    const language = Application.ActiveSheet.Range(`E${i}`).Value2
    const math = Application.ActiveSheet.Range(`F${i}`).Value2
    const english = Application.ActiveSheet.Range(`G${i}`).Value2
    const literature = Application.ActiveSheet.Range(`H${i}`).Value2
    const total = Application.ActiveSheet.Range(`I${i}`).Value2
    const rate = Application.ActiveSheet.Range(`J${i}`).Value2

    data.push({
      id: studentId,
      name: studentName,
      sex: sex,
      className: className,
      language: language,
      math: math,
      english: english,
      literature: literature,
      total: total,
      rate: rate
    })
    break
  }
}

// 返回匹配的学生数据
return data
```

## 2. 电商数据同步

电商数据同步是一个重要的环节，确保电商平台在销售和运营方面能够高效运作。在数据来源上，它可能是多个不同的系统，包括数据库、ERP系统、CRM系统等。而开发者首先要做的是手动导出或者 webhook 的形式获取到商品信息、订单信息、物流信息等数据，然后在您的个人服务器内得到这些数据，进行数据清洗和转换，最后再通过脚本令牌写入到金山文档智能表格内，完成数据的同步。
