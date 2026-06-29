# ImapDB (Immune and panel DataBase)
This repository contains the source code for ImapDB, an immune research GUI written in PyQt, pandas, and openpyxl. ImapDB aims to serve as a guide for immune cell research, providing curated immune markers and antibody panels for FACS and CyTOF.

> **Note:** ImapDB is a local desktop application, not an online or browser-based tool.
> The repository must be **downloaded to your own computer** before use, and the program
> is then run locally as described in the *Installation and Running* section below.

### Components are tested on:
-	Windows, macOS & Linux
-	Python 3.9 and above
-	PyQt5 5.11 and above
-	Pandas 2.2.2
-	Openpyxl 3.1.5

 <br/>

## ScreenShot
 The program's interface consists of three buttons on the main window: Immune Cell Marker, CyTOF Antibody Panel, and FACS Antibody Panel. When each button is clicked, a dialog box appears that provides information related to the button's name. Within the dialog box, users can select from the listed marker and antibody panel information in the combo box of the dialog and perform a search.
 <br/>

![image](ImapDB_mainwindow.png)
 <br/>

## Installation and Running

No programming experience is required. Just follow the steps for your operating
system below. The program runs the same way on **Windows, macOS, and Linux**.

### Step 1 — Install Python (one time)

If Python 3.9 or newer is already installed, skip this step.

- **Windows:** download the installer from https://www.python.org/downloads/,
  run it, and on the first screen **check "Add Python to PATH"** before clicking
  *Install Now*.
- **macOS:** download the installer from https://www.python.org/downloads/, or
  install with Homebrew using `brew install python`.
- **Linux (Ubuntu/Debian):** run `sudo apt update && sudo apt install python3 python3-pip`.

### Step 2 — Download ImapDB to your computer

ImapDB must be downloaded and run locally; there is no online version.

1. On this GitHub page, click the green **`< > Code`** button, then **`Download ZIP`**.
2. Unzip it. Keep **all files together in the same folder** — the program needs
   the `.ui`, `.xlsx`, and image files that are included. (No manual path editing
   is required; the program automatically finds the data files that sit next to it.)

### Step 3 — Install the required libraries (one time)

Open a terminal and run the commands below.
(On **Windows**, press the Windows key, type `cmd`, and press Enter to open the
Command Prompt. On **macOS**, open Terminal from Applications → Utilities. On
**Linux**, open your terminal application.)

#### PyQt5

    pip install pyqt5

#### Pandas

    pip install pandas

#### openpyxl

    pip install openpyxl

### Step 4 — Run the program

First move into the unzipped folder, then start the program.

- **Windows** (folder in Downloads, as an example):

      cd %USERPROFILE%\Downloads\Cell_Marker_GUI
      python ImapDB_GUI.py

- **macOS / Linux** (folder in Downloads, as an example):

      cd ~/Downloads/Cell_Marker_GUI
      python3 ImapDB_GUI.py

The main window will open automatically.

### Troubleshooting

- **"python is not recognized" (Windows) or "command not found" (macOS/Linux):**
  Python is not on your PATH. Reinstall Python (on Windows, check
  *"Add Python to PATH"*), open a new terminal, and try again.
- **"No module named PyQt5" (or pandas / openpyxl):** re-run the matching
  `pip install ...` command from Step 3.
- **Blank tables or a "file not found" message:** make sure you ran the program
  from **inside** the unzipped folder (the `cd` step in Step 4), so the data
  files sit next to the script.

 <br/>

## Source of Input files
The references for immune markers of surface and intracellular can be found in the supplementary data of the manuscript slated for publication. (Not yet published.)
<br/>

## License
This project is available under the MIT license. See the included LICENSE file.
