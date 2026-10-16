---
title: "Use Cases"
source: https://airsheet.wps.cn/docs/en-US/apitoken/scenario.html
updated: 1789700916000
hash: BsY-YBob
---

# Use Cases

With the support of Script Token, the powerful capabilities of Smart Spreadsheet are fully unleashed, enabling developers to execute tasks in a more efficient and accurate manner. Whether it's web scraping, data analysis, or automation processes, I can use Script Token to accomplish them.

Below are some examples and explanations we wrote based on real scenarios, hoping to provide some inspiration for developers.

## 1. Private Information Query

Modern society is increasingly valuing personal privacy. This trend is reflected in many aspects.

First, in the education field, many schools are strengthening the protection of students' information during their school years, prohibiting the sale or sharing of students' personal information with others. This includes students' **family information, educational records, exam scores**, etc.

Second, in the healthcare field, patient privacy protection has become an important issue. Doctors and other healthcare workers need to comply with strict privacy regulations to ensure that **patients' personal information** is not leaked.

In addition, in the social media field, many platforms have also begun to strengthen the protection of user information. They have adopted stricter data security measures to ensure that users' data is not leaked or misused.

Below, we will show an online example of student score query. You can experience the private information query functionality implemented using Script Token online. The existing student score sheet is shown below:

The underlying implementation logic of the above online example uses **Script Token to execute AirScript scripts**, querying data in the spreadsheet and returning it to the interface for display. The script example code is as follows:

```js
// Get student ID
const student_id = '1'
// Get student name
const student_name = 'Jin Xiaomei'
// Get the number of rows in the used range
const rowCount = Application.ActiveSheet.UsedRange.Rows.Count

// Get matching student data
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

// Return matching student data
return data
```

## 2. E-commerce Data Sync

E-commerce data synchronization is an important aspect that ensures e-commerce platforms can operate efficiently in sales and operations. In terms of data sources, it may come from multiple different systems, including databases, ERP systems, CRM systems, etc. What developers need to do first is manually export or obtain product information, order information, logistics information, and other data through webhooks. Then get this data on your personal server, perform data cleaning and transformation, and finally write it to WPS Docs Smart Spreadsheet through Script Token to complete the data synchronization.
