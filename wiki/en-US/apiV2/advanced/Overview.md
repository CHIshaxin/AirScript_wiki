---
title: "Overview"
source: https://airsheet.wps.cn/docs/en-US/apiV2/advanced/Overview.html
updated: 1789700916000
hash: BLGU7-om
---

# Overview

With AirScript's advanced services, developers can connect to certain public WPS Docs APIs with minimal setup. Their usage is very similar to the built-in functions of AirScript scripts.

AirScript automatically handles the [authorization process](#authorization-process) at runtime. However, developers must [enable advanced services](#enable-advanced-services) before using them in scripts. Skipping this step will throw an **undefined** error because the service cannot be found.

## Enable Advanced Services

To use advanced services, follow these instructions:

1. Open **Efficiency** - **AirScript Editor Tool** to open the editor page.
2. Click **Services** at the top of the AirScript Editor Tool.
3. Click **Add Service**.
4. Select a service and click **Confirm**.

After enabling an advanced service, it will appear in auto-completion.

## Authorization Process

AirScript requires user authorization to access private data in advanced services.

### Granting Execution Permission

AirScript determines the authorization scope based on the advanced service configuration when developers write scripts (e.g., accessing specified files or network). If the script requires authorization, an authorization dialog will pop up when the user runs the script, describing the authorization scope involved in this script.

Normal code changes will not clear user authorization for the script. However, if developers change the advanced service configuration (add, modify, or delete), user authorization for the script will be cleared, and the authorization process will be triggered again when running the script.

**Note**: Scripts in **My Scripts** are completely owned by the user, and running scripts in this category does not require triggering the authorization process.

### Revoke Authorization

Users can manually revoke authorization for authorized scripts by following these instructions:

1. Open **Efficiency** - **AirScript Editor Tool** to open the editor page.
2. Find the script you want to revoke authorization for under Document Shared Scripts, click **...** to show more options.
3. Click **Revoke Service Authorization**

## Usage Restrictions

To prevent providing malicious scripts to users, there are some restrictions on using advanced services for security reasons.

- Using advanced services too frequently. When this happens, the script will throw an obvious error notifying the user of abnormal calls.
- When using the [HTTP](/docs/en-US/apiV2/advanced/HTTP.html) service, using IP addresses for requests is prohibited, and using ports for requests is prohibited.
- When using the [HTTP](/docs/en-US/apiV2/advanced/HTTP.html) service, the maximum message body content received is 2M. Exceeding 2M will throw an error.
- If the [File](/docs/en-US/apiV2/advanced/KSDrive.html#file) object obtained from [KSDrive.openFile](/docs/en-US/apiV2/advanced/KSDrive.html#openfile-openinfo) is not closed, using [KSDrive.openFile](/docs/en-US/apiV2/advanced/KSDrive.html#openfile-openinfo) again will throw an error.
