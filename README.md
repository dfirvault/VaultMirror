# VaultMirror
Simple bi-directional or one-way sync tool, suitable for archiving cases across multiple drives. VaultMirror is a robust, lightweight file synchronization tool designed for DFIR (Digital Forensics and Incident Response) professionals. It ensures that your evidence folders, toolsets, or case files stay synchronized across different drives or network shares using Windows Task Scheduler for background automation.

## 🚀 Quick Start (Pre-compiled Release)

We have released **v0.1**, which includes a standalone `.exe`. This is ideal for Windows systems where you cannot or do not want to install Python.

1. Download `VaultMirror.exe` from the [Releases](https://github.com/dfirvault/VaultMirror/releases) page.
2. Right-click `VaultMirror.exe` and select **Run as Administrator** (required to interact with Task Scheduler).
3. Select **Create New Sync Task**.
4. Follow the prompts to select your folders via the GUI and set your sync interval.

## ✨ Features

* **GUI Folder Selection:** No more manual path typing; use native Windows folder pickers.
* **Intelligent Bi-directional Sync:** Uses last-modified timestamps and a state manifest to ensure the newest version of a file is preserved.
* **Deletion Propagation:** Solves the "split-brain" issue. If you delete a file in Folder A, VaultMirror recognizes the change in the state manifest and removes it from Folder B.
* **Standalone Architecture:** The compiled version handles its own background tasks. **No Python installation is required** on the target system.
* **Task Management:** Easily create, run manually, or delete synchronization tasks directly from the console interface.
* **Stealthy Background Operation:** Leverages Windows Task Scheduler to run sync jobs at your preferred interval (Minute, Hourly, Daily, Weekly).

## 🛠 How it Works

VaultMirror uses a **state-based manifest system** to track changes:

1.  **Memory:** It stores a JSON state file in `%APPDATA%\VaultMirror\sync-states\`.
2.  **Comparison:** Every time a sync triggers, it compares the current folder contents against the last known state.
3.  **Action:** * **New File:** If a file exists in A but not in B or the Manifest, it is copied to B.
    * **Updated File:** If a file is newer in A than in B, it overwrites B.
    * **Deletion:** If a file is in the Manifest but missing from A, it is automatically deleted from B to maintain a true mirror.



## 🔨 Installation from Source

If you prefer to run from source or build the binary yourself:

### Prerequisites
* Python 3.10+
* `pip install pywin32`

### Build Instructions
To create your own standalone executable:
```bash
pip install pyinstaller
pyinstaller --onefile --uac-admin --name="VaultMirror" --icon=icon.ico --hidden-import="win32timezone" --hidden-import="win32com.client" VaultMirror.py
