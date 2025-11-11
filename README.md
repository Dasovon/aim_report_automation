<p align="center">
  <img src="repo_banner.png" alt="AIM Report Automation Banner" width="800">
</p>

# 📘 AIM Report Automation

**AIM Report Automation** is a cross-platform workflow developed for Texas A&M University’s **Facilities & Energy Services (FES)**.  
It streamlines **AIM work order formatting, inspection tracking, and dashboard generation** using a hybrid system of **Excel VBA** and **Python automation**.

---

## 🧭 Overview

This project automates the process of preparing **AIM work order exports** for building inspections.  
It takes the raw `browse.csv` export from **AggieWorks / AiM** and produces a fully formatted Excel workbook that includes:

- Extracted **Floor** and **Room** data from descriptions  
- Calculated **Age (Days)** using business days  
- **Inspection Status** dropdowns with row coloring (Pending, Complete, Incomplete, Needs Review)  
- A **dashboard sheet** with counts and averages  
- Cross-platform compatibility (Windows + macOS)  
- Two-way sync between sheets via VBA events  

---

## 📁 Repository Tree

```
aim_report_automation/
├── .gitignore
├── README.md
├── VBA/
│   ├── clsAIM_AppEvents.cls          # Event-driven sync between sheets
│   ├── mod_AIM_Formatter.bas         # Main VBA formatter logic
│   └── mod_AIM_WatcherCore.bas       # Event initializer for sync
│
└── Python/
    ├── aim_formatter.py              # Python formatter script (cross-platform)
    ├── template.xlsm                 # Macro-enabled Excel template with VBA logic
    ├── output/                       # Optional folder for generated files
    └── venv/                         # Local virtual environment (ignored in Git)
```

---

## 🧩 Components

| Folder | Contents |
|---------|-----------|
| **VBA/** | Core macros used in the Excel template, including:<br>• `mod_AIM_Formatter.bas` – Main formatter logic<br>• `mod_AIM_WatcherCore.bas` – VBA event initializer<br>• `clsAIM_AppEvents.cls` – Two-way inspection sync handler |
| **Python/** | Automation scripts and supporting files:<br>• `aim_formatter.py` – Python version of the formatter<br>• `template.xlsm` – Macro-enabled Excel template (includes VBA above)<br>• `venv/` – Virtual environment (ignored by Git)<br>• `output/` – Optional folder for generated workbooks |
| **.gitignore** | Keeps repo clean (ignores temp, OneDrive, venv, and output files). |

---

## ⚙️ Setup

### **Windows**
1. Install [Python 3.10+](https://www.python.org/downloads/).  
2. Clone this repository:
   ```bash
   git clone https://github.com/Dasovon/aim_report_automation.git
   cd aim_report_automation/Python
   ```
3. Create and activate the virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```
4. Install dependencies:
   ```bash
   pip install pandas openpyxl numpy
   ```
5. Run the formatter:
   ```bash
   python aim_formatter.py
   ```

### **macOS**
1. Ensure Homebrew Python 3.10+ is installed:
   ```bash
   brew install python
   ```
2. Navigate to your Python folder:
   ```bash
   cd ~/Desktop/excelProgram/Python
   python3 -m venv venv
   source venv/bin/activate
   pip install pandas openpyxl numpy
   ```
3. Run:
   ```bash
   python3 aim_formatter.py
   ```

---

## 🧾 Workflow Summary

1. Export `browse.csv` from **AggieWorks → Reports → Browse Work Orders**.  
2. Run the Python script.  
3. Choose your CSV when prompted, then choose a save name (defaults to `YYYYMMDD_WOs.xlsm`).  
4. The formatted workbook will open automatically with:  
   - Sorted and color-coded work orders  
   - Inspection dropdowns  
   - Dashboard tab  
   - VBA logic active (two-way sync, row coloring)

---

## 💡 Tips

- **Template location:**  
  The formatter expects `template.xlsm` in the same folder as `aim_formatter.py`.  

- **Virtual environment:**  
  The `venv/` directory is ignored in Git; recreate it as needed.  

- **Output destination:**  
  Files save automatically to `~/Downloads` by default on macOS, or the Python directory on Windows.  

---

## 🧰 Future Improvements

- Auto-launch dashboard charts  
- Email summary reports  
- Integration with TAMU shared drives  
- Live AiM API pulls (planned)

---

## 🏛️ Credits

Developed by **Ryan Novosad**  
Facilities Coordinator — Texas A&M University  
Department of Facilities & Energy Services  

---

## 🔒 License
This project is for internal and educational use within Texas A&M University.  
Do not redistribute or modify for external use without permission.

---

## 🧾 Git Quick Reference

| Command | Purpose |
|----------|----------|
| `git status` | Check what’s changed locally |
| `git add .` | Stage all new or modified files |
| `git commit -m "message"` | Save a new version with a message |
| `git push` | Upload your commits to GitHub |
| `git pull` | Download the latest version from GitHub |

**Workflow summary:**  
After each change:  
```bash
git status
git add .
git commit -m "Describe your change"
git push
```
When working from another machine:  
```bash
git pull
```

---

## 👥 For Collaborators

If you are another **Facilities Coordinator, SSC lead, or TAMU staff member**, you can use this automation without editing the VBA code.

**Steps:**
1. Clone or download the repository:
   ```bash
   git clone https://github.com/Dasovon/aim_report_automation.git
   ```
2. Open the `/Python` folder and run `aim_formatter.py`.  
3. Select your daily or weekly **AIM export (browse.csv)** when prompted.  
4. The script automatically:
   - Generates a formatted Excel workbook  
   - Applies all color and dropdown logic  
   - Opens the file ready for inspection review  
5. The VBA logic inside the workbook manages **real-time row color changes** and **status syncing** across all sheets.

**No editing required** — just run and review.
