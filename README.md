# Microsoft Teams Message Sender with Adaptive Card (PowerShell)

This PowerShell script demonstrates how to send messages to Microsoft Teams users using the Microsoft Graph SDK and Adaptive Cards. It uses authentication via Microsoft Graph, creates a one-on-one chat, and sends a customized Adaptive Card message to users listed in an Excel file. The script also includes basic error handling and logging features.

## Prerequisites

Before running the script, make sure you have the following installed:

- **PowerShell** (Version 7.x or later recommended)
- **Microsoft Graph SDK** (Version 27.0.0 or higher)
  - `Microsoft.Graph.Authentication` module
  - `Microsoft.Graph.Teams` module
- **ImportExcel module** (for reading Excel files)

### Required PowerShell Modules

Install the modules using:

```powershell
Install-Module Microsoft.Graph.Authentication -Force -AllowClobber
Install-Module Microsoft.Graph.Teams -Force -AllowClobber
Install-Module ImportExcel -Force -AllowClobber
```

## Service Account Requirements

- It is **highly recommended** to use a **dedicated service account** to run this script.
- The **service account must have an active Microsoft Teams license** to be able to initiate and send messages in 1:1 chats.
- Optionally, you can **add a disclaimer at the end of the Adaptive Card** content such as:
  > "*Note: This is a system-generated message. Please do not reply to this chat.*"

## Script Overview

The script authenticates with Microsoft Graph, reads a list of users from an Excel file (`Users.xlsx`), and sends an Adaptive Card message to each user in a new 1:1 chat. It handles simple errors and logs the outcome (success or failure) to a CSV file. It uses the Microsoft Graph SDK with the Microsoft.Graph.Authentication and Microsoft.Graph.Teams modules (version 27.0.0)

## Features

- ✅ **Microsoft Graph Authentication** via browser popup
- 📊 **Reads user list** from an Excel file (column: `UPN`)
- 🧾 **Sends Adaptive Card messages** in 1:1 Teams chats
- 🪵 **Logging** of all messages sent (with timestamp and status)
- 🛠️ **Easy customization** of Adaptive Card content
- ⚠️ **Basic error handling**

## Excel File

Place your Excel file in the desired location and ensure it contains a column titled `UPN` with the user principal names.

## Logs

Logs are written to `CopilotMessageLog.csv` in CSV format:

```
Timestamp,UPN,Status,Message
```

Each row logs the result of sending a message to a user.

## Important Notes

- Test the script thoroughly in a development or test environment before deploying in production.
- Always validate that all modules are up-to-date and compatible with your PowerShell version.
- Customize the Adaptive Card content freely — including additional prompts, text, links, or images.
- It is **advisable to include a note or disclaimer** in the Adaptive Card if the service account is not monitored for replies.

## Disclaimer

This is a sample script provided for demonstration purposes only. It is your responsibility to test and validate the script thoroughly before using it in production. Modify it to meet your organizational and compliance requirements.
