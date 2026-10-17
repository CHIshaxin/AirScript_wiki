---
title: "View Configuration"
source: https://airsheet.wps.cn/docs/en-US/start/view-setting.html
updated: 1789700916000
hash: ClTiqHXn
---

# View Configuration

To facilitate users in delivering their scripts to third parties without development experience, we have introduced parameter visualization configuration. Users can open the parameter arrangement tool by clicking `View Configuration` in the script toolbar. By configuring `form` forms combined with scripts, quick operation tools can be generated.

## Relationship Between View and Script

- The form parameters configured in the view will be mounted under `Context.argv` in the script editor, and can be directly referenced during programming. `Context` is a global variable, similar to `Application`.
- The parameters configured in the view will generate syntax hints in real-time in the code editor. Enter Context.argv to see the configured parameters.
- Preview view, click the `Run Script` button in the view to run the script with form input parameters, and the log will be printed using the code editor's log box.

## Simple Use Case

#### 1. Configure Parameters

Open view configuration and add the following parameters:

| Parameter Name | Variable Name | Variable Type | Input Hint |
| --- | --- | --- | --- |
| Select Sheet | sheet | string | Enter sheet name: "sheet1" |
| Select Cell | range | string | Enter cell: "A1" |
| Input Value | value | string | Enter value to set |

#### 2. Input Code

```js
Application.Sheets.SelectByName(Context.argv.sheet).Activate()
Application.Range(Context.argv.range).Value=Context.argv.value
```

#### 3. Editing Result

After completing steps 1 and 2, the editing result is shown below:

After you finish writing the first two lines of your parameterized script, the AirScript editor will automatically detect the `Context.argv` parameters you defined in the code. A smart code completion popup will appear as you type, showing all available parameters (`sheet`, `range`, `value`) with their type hints.

At the same time, a visual configuration panel will be generated on the right side of the editor. This panel lets users run your script without writing any code—they just need to fill in the required values: select the target worksheet, specify the cell range, and enter the value to set. You can also add more parameters or clear all existing ones with one click, then hit "Run Script" to test.

#### 4. Execution Result

Click `Preview`, enter test values, run the script, and the result is shown below:

Here's what happens when you test the script:

1. Click the "Preview" button in the right configuration panel
2. Enter test values: select `Sheet1` as the target worksheet, type `A1:B2` as the cell range, and enter `Test` as the value to set
3. Click "Run Script"

The script will execute immediately, and you'll see real-time logs at the bottom showing the initialization, start, and completion of the execution. The result will appear directly in your spreadsheet: all cells from A1 to B2 will be filled with the word "Test" exactly as you specified.
