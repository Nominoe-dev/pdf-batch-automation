# PDF Batch Automation

A PowerShell automation tool designed to simplify repetitive document-processing workflows by automatically identifying, validating, merging, and renaming documents in batch.

The application provides a graphical interface that allows users to select a root folder, review eligible report folders, choose which folders to process, and automatically generate the final PDF documents.

---

## Overview

Processing multiple reports manually can involve several repetitive operations:

1. Navigating through multiple folders.
2. Identifying the correct documents.
3. Checking that all required files are present.
4. Converting and combining documents into a single PDF.
5. Renaming the resulting PDF.
6. Repeating the same process for each report.

This project automates this workflow using **PowerShell** and a **Windows Forms graphical interface**.

Instead of processing each folder individually, the application scans the selected directory recursively, identifies folders matching the expected report structure, validates their contents, and allows the user to process multiple folders in a single operation.

---

## Features

- 📁 **Root folder selection**
  - Allows the user to select the directory from which the search should start.

- 🔎 **Recursive folder search**
  - Automatically searches through subdirectories for folders matching the expected report naming pattern.
  - Supports variations of the report folder name, including singular/plural and accented versions.

- ✅ **Automatic document validation**
  - Each detected report folder is checked for the required documents:
    - Excel files (`.xls` / `.xlsx`)
    - PDF files containing `ANEXO` in the filename
    - PowerPoint files (`.ppt` / `.pptx`)

- 🖥️ **Graphical User Interface**
  - Uses Windows Forms to provide a simple graphical interface.
  - Displays detected report folders in a checklist.
  - Valid folders are automatically selected.

- ☑️ **Batch selection**
  - Select or deselect individual folders.
  - Includes "Select All" and "Deselect All" functionality.

- ⚙️ **Automated PDF processing**
  - Launches Wondershare PDFelement to combine the required documents.
  - Processes multiple folders sequentially.

- ⏳ **Process monitoring**
  - Waits for PDFelement to become available.
  - Brings the application window to the foreground when necessary.
  - Waits for the generated combined PDF before continuing.

- 📝 **Automatic file naming**
  - The resulting PDF is renamed using the filename of the corresponding Excel document.

- 📊 **Processing summary**
  - Displays the number of folders successfully processed when the operation is complete.

---

## Workflow

The application follows four main stages.

### 1. Select the root folder

When the application starts, the user is prompted to select a root directory.

```text
Root Folder
    │
    ├── Folder A
    ├── Folder B
    └── ...
