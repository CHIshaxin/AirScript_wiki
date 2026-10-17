---
title: "Overview"
source: https://airsheet.wps.cn/docs/en-US/api/advanced/Overview.html
updated: 1789700916000
hash: BOIGqWGK
---

# Overview

With AirScript advanced services, developers only need minimal setup to connect to certain public WPS Docs APIs. Their usage is very similar to AirScript's built-in functions.

AirScript handles the [authorization flow](#authorization-flow) automatically at runtime. However, developers must [enable advanced services](#enable-advanced-services) to use a service in a script; skipping this step throws an **undefined** error because the service cannot be found.

## Enable advanced services

To use an advanced service, follow these steps:

1. Open **Efficiency** - **AirScript editor** to open the editing page.
2. Click **Services** in the AirScript editor toolbar.
3. Click **Add service**.
4. Select a service, then click **Confirm**.

After an advanced service is enabled, it appears in autocomplete.

## Authorization flow

AirScript requires user authorization to access private data in advanced services.

### Grant run permission

AirScript determines the authorization scope from the advanced service configuration when the script was written (for example, access to a specific file or network access). If the script needs authorization, an authorization dialog appears when the user runs the script, describing the scope of permissions involved.

Ordinary code changes do not clear the user's authorization for the script. If the developer changes the advanced service configuration (add, modify, or remove), the user's authorization for the script is cleared, and the authorization flow runs again on the next run.

**Note**: Scripts under **My Scripts** are fully owned by the user; running scripts in that category does not trigger the authorization flow.

### Revoke authorization

Users can manually revoke authorization for a script that was authorized. Follow these steps:

1. Open **Efficiency** - **AirScript editor** to open the editing page.
2. Under file-shared scripts, find the script to revoke, click **...** for more actions.
3. Click **Revoke service authorization**

## Usage limits

To prevent malicious scripts, advanced services have security-related limits.

- Excessive use of advanced services causes the script run to fail with a clear error indicating abnormal usage.
- When using the [HTTP](HTTP.md) service, requests using IP addresses are not allowed, and requests using ports are not allowed.
- When using the [HTTP](HTTP.md) service, the response body is limited to 2 MB; exceeding 2 MB throws an error.
- If the [File](KSDrive.md#file) object from [KSDrive.openFile](KSDrive.md#openfile-openinfo) is used again with [KSDrive.openFile](KSDrive.md#openfile-openinfo) without calling close, an error is thrown.
