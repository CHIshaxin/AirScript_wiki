---
title: "What Is AirScript 2.0"
source: https://airsheet.wps.cn/docs/en-US/apiV2/overview.html
updated: 1789700916000
hash: aBXUKRx4
---

# What Is AirScript 2.0

AirScript 1.0 scripts cover typical document automation needs. However, there are still limitations in API coverage, execution performance, and support for advanced JavaScript syntax. To overcome these limits, we introduced AirScript 2.0.

AirScript 2.0 uses a new architecture and delivers major upgrades in capability and performance. The APIs are fully aligned with WPS JS macros, execution is faster, and modern JavaScript syntax is fully supported.

Although most AirScript 1.0 APIs can run in AirScript 2.0, some API definitions conflict between versions, so full compatibility is not possible. You must use `1.0` and `2.0` to identify the script runtime version. AirScript 2.0 is a major upgrade, and some APIs are not compatible with AirScript 1.0.

# Highlights of AirScript 2.0

### ✔️ Fully compatible with WPS JS macros

AirScript 2.0 is fully compatible with WPS JS macros and provides 2,000+ APIs (about four times as many as 1.0), greatly expanding what you can build. Whether you are automating simple tasks or developing complex scripts, AirScript 2.0 has you covered. You do not need to learn a new API model; you can use the same APIs to develop scripts in WPS and in online spreadsheets.

### 🚀 Faster execution

We optimized logical execution and API calls in AirScript 2.0 so scripts run more efficiently and respond more quickly. The gain is especially noticeable in large scripts that call APIs frequently. For example, looping over 1,000 cells to trim leading and trailing spaces and write the values back runs about twice as fast in 2.0 as in 1.0.

> You can create separate 1.0 and 2.0 scripts and run the following code to feel the performance improvement in 2.0.

```javascript
// Trim leading/trailing spaces for the first 1k cells in column 1 of the active sheet, and log execution time
const start = Date.now();

const sheet = ActiveSheet;
for (let i = 1; i <= 1000; i++) {
  const cell = sheet.Cells(i, 1);
  cell.Value2 = cell.Text.trim();
}

console.log(Date.now() - start);
```

### 🌐 Full support for modern JavaScript syntax

AirScript 2.0 fully supports modern JavaScript syntax, including `await`/`async`, `Promise`, `class`, and more (not supported in 1.0), for a more modern and productive development experience. This is especially helpful for complex scripts that are easier to write and maintain.

# Supported scenarios

AirScript 2.0 is designed to cover everything AirScript 1.0 can do. In the first phase, worksheet scenarios are supported; Databook APIs are not yet available. If your script targets Databook scenarios, keep using AirScript 1.0 for now; we will gradually add Databook API support in 2.0.

> AirScript 2.0 is currently in beta. Some API capabilities will be filled in over upcoming releases.

|  | AirScript 1.0 | AirScript 2.0 |
| --- | --- | --- |
| Worksheet | APIs (500+) | APIs (2,000+) |
| Databook | APIs | APIs (coming soon) |
| Advanced services | Network API<br>Cloud document API<br>Email API<br>Database API | Network API<br>Cloud document API<br>Email API (coming soon)<br>Database API (coming soon) |

# How to use AirScript 2.0

> Multidimensional tables depend on Databook support, which is not yet available.

### How to create a 2.0 script

1. Open the AirScript script editor

**How to create a new AirScript 2.0 script**

First, open the dedicated script editor from the main menu:

- Go to the top navigation bar and click the **Efficiency** tab
- Find the **Advanced Development** dropdown menu in the secondary toolbar
- Select **AirScript Script Editor** from the dropdown options

This will launch the full-featured script editor where you can create, edit, and run both 1.0 and 2.0 version scripts.

2. Create a script (shared document scripts / My scripts)

Click `+` to create a 2.0 script directly.

You can also choose the script version from the dropdown menu. When creating a new script, you can explicitly choose which version to use by clicking the small dropdown arrow next to the plus sign (+) in the left sidebar (next to "Document Shared Scripts").

Two options are available:

- **AirScript 2.0 Beta (Recommended)**: Supports over 2000 API calls, has full compatibility with client-side script syntax, and offers the latest features and performance improvements
- **AirScript 1.0**: Has limited API support, only recommended for maintaining old legacy scripts

Note that the exact same version selection process applies to scripts created under the "My Scripts" section.

> The steps are the same for "My scripts."

### How to find the 2.0 API documentation

API documentation is available in 1.0 and 2.0 versions. The APIs differ between versions. See [API documentation (2.0)](overview.md).
