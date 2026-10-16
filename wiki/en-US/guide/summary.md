---
title: "AirScript Overview"
source: https://airsheet.wps.cn/docs/en-US/guide/summary.html
updated: 1789700916000
hash: mcLtySsz
---

# AirScript Overview

WPS Docs AirScript is a simple, fast, and lightweight scripting application development platform. Built on cloud technology, it enables you to quickly and easily create business applications that interact with WPS Docs Office files.

We provide a code editor within WPS Docs components, and your code will run securely on our servers. You can write logic code using modern JavaScript, and call AirScript's powerful built-in component (spreadsheet) APIs to build solutions for specific scenarios.

AirScript is a pioneering platform that provides this service.

## What Can AirScript Do

AirScript currently serves as a secondary development platform for online spreadsheets and multidimensional tables, providing capabilities for CRUD operations on spreadsheet data, cell format modification, and property settings through programming.

#### Tool Advantages

- No need to set up a local environment, develop scripts directly in the cloud within the document.
- Built-in customized global Application object, editor intelligent prompts, development, debugging, and execution all in one place.
- Synchronous property retrieval, synchronous method execution, reducing the mental burden of traditional asynchronous calls.
- Thanks to the integrated development environment, whether creating scheduled tasks, batch processing data, or automatically generating documents, developers can fully unleash their imagination here.

## How to Use AirScript

Open the online spreadsheet KSheet, switch to the "Efficiency" tab, find "AirScript Editor Tool" in the secondary toolbar below, and click to open the AirScript editor.

**Getting started with AirScript: Run your first Hello World script**

Here's how to launch the editor and run your first test script:

- Open the KSheet online spreadsheet
- Switch to the **Efficiency** tab in the top navigation bar
- Find and click **AirScript Editor** in the secondary toolbar below
- The editor will open with a default "HelloWorld" script that includes a welcome message, links to the official documentation, and an important note: scripts run on the first worksheet by default unless you specify otherwise in code
- Click the **Run** button in the top toolbar to execute the script
- You'll see real-time execution logs at the bottom of the editor, showing the script environment initializing, execution starting, and finally the "Hello, World!" message printed to the log console

This confirms that your AirScript environment is set up correctly and ready for you to write more complex automation scripts.

Tip

Only collaborators with file editing permissions can open the development tool

### Editor Features

The script editor is divided into two parts: the left part is the script file management area; the right part is the coding and execution area. It provides capabilities for script creation, deletion, modification, execution, and runtime log generation.

For document operation APIs, the editor has built-in Application object syntax hints, allowing you to write the required OpenAPIs directly based on documentation and syntax prompts.

Feature details:

- Script Management
  - The left side of the editor and the top operation area on the right provide common operations such as creating, saving, deleting, and running code files;
  - Script files are stored in the cloud and divided into "My Scripts" and "Document Shared Scripts", where My Scripts follow the account, and Document Shared Scripts follow the current document;
- Code Editing
  - The code editor is based on the open-source `monaco editor`, and user operations are basically similar to `VS Code`;
  - Code Saving
    - Spreadsheet: Code is automatically saved locally but not automatically synced to the cloud. Users need to manually click "Save" or "Ctrl+S" to sync to the cloud. Unsynced code files will display a "small green dot" before the file name;
    - Multidimensional Table: Switching scripts and closing the development tool will automatically save the code and sync to the cloud;
- Execution Log
  - Displays code execution information, allowing users to view script execution status and print results based on the returned information;
  - If the code has errors, error information and location will be printed. Clicking the location will navigate the editor to the error code block.
