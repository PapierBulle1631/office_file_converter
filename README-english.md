# Old Office File Manager Manual

_For: Mayr-Melnhof Group_  
_By: Heriaud Thomas_  
_Date: 22/01/2025_

---

## Table of Contents

- [Before Opening the Application](#before-opening-the-application)
- [Two Ways to Open the Application](#two-ways-to-open-the-application)
  - [User Mode](#user-mode)
  - [Admin Mode](#admin-mode)
- [Functionality](#functionality)
  - [Folder Selection](#folder-selection)
  - [Logs](#logs)
- [The Report (.csv)](#the-report-csv)

---

## Before Opening the Application

A few points to check:

- Be on **Windows** and have **PowerShell 5.0 or higher** (pre-installed with Windows 10).
- For file conversion, ensure you have the latest versions of **Word**, **PowerPoint**, and **Excel**.

---

## Two Ways to Open the Application

### User Mode
This mode only allows **file reading** with the same permissions as the user and **report generation**.

### Admin Mode
This mode enables **report generation** AND **file conversion**, depending on the installed Office applications.

---

## Functionality

### Folder Selection

- Use the **"Browse"** button to select the folder to scan. All subfolders will also be scanned for `.doc`, `.xls`, and `.ppt` files.
- Use the **"Browse"** button to select the folder where a CSV report will be generated.  
  **Important:** Write permissions in the folder are required since there is no default folder.

---

### Logs

This scrollable area displays the following:
- Folder selection errors.
- Complete paths of selected folders.
- Files with the targeted extensions.
- Conversion progress.
- Names and paths of the report.

#### Additional Options:
- A checkbox, if selected, will trigger file conversion (only works if the application is launched in admin mode).
- A button starts the file search process, performs conversion if enabled, and displays steps in the log area.

---

## The Report (.csv)

The generated report is a CSV document. If the user's delimiter is incorrect, additional steps are required to make the data readable.

### Correction Procedure:

1. Go to the **Data** tab, then click **From Table/Range** in the _Get & Transform Data_ section.
2. Check **"My table has headers"** to allow sorting of large datasets, then confirm. Do not change the range: it is automatically correct.
3. In the **Home** tab, under the _Transform_ section, select **Split Column** and then the **By Delimiter** option.
4. Ensure **"Comma"** is selected, then confirm with **OK**.
5. Confirm with **Close & Load** to make it appear in the original document.
6. Save the document.

---

_For: Mayr-Melnhof Group_  
_By: Heriaud Thomas_  
_Date: 22/01/2025_
