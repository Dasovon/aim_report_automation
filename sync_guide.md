# 🧭 Sync Guide: AIM Report Automation Repository

Keep your **work** and **personal computers** perfectly synchronized with your main GitHub repository.

---

## 🚀 1. Initial Setup

### 🖥️ On Your Work Computer
Your repository is already configured. Verify it’s up to date:
```bash
git status
```
Expected output:
```
On branch main
Your branch is up to date with 'origin/main'.
```

### 💻 On Your Personal Computer
If a copy already exists, back it up:
```bash
cd ~/Desktop/excelProgram
mv aim_report_automation aim_report_automation_old
```
Clone the clean version:
```bash
git clone https://github.com/Dasovon/aim_report_automation.git
```
This ensures both computers share the same clean, tracked structure.

---

## 🔁 2. Sync Workflow

**Pull latest updates (on either machine):**
```bash
git pull origin main
```

**Push your new changes:**
```bash
git add .
git commit -m "Describe your change here"
git push origin main
```

GitHub is your single source of truth — both computers will stay in perfect sync.

---

## 🧹 3. Clean Folder Rules

The `.gitignore` file ensures local-only files never sync:
- `venv/` (Python virtual environments)
- `Python/output/` (auto-generated Excel files)
- `Python/Python_backup/` (local backups)
- Excel and OS temp files (`.tmp`, `.DS_Store`, etc.)

**Test your ignore list:**
```bash
git check-ignore -v Python/output/test.xlsx Python/venv/bin/activate Python/Python_backup/aim_formatter.py
```
If files appear in the output, they’re correctly ignored.

---

## ⚙️ 4. Common Issues & Fixes

**If push is rejected:**
```bash
git pull --rebase origin main
```
Then:
```bash
git push origin main
```

**If a file isn’t ignored:**
Add its path manually to `.gitignore` and commit the change.

**If your repo gets messy:**
```bash
rm -rf aim_report_automation
```
Then re-clone it cleanly.

---

## 🔍 5. Verify Remote Connection

Check your GitHub link:
```bash
git remote -v
```
Expected result:
```
origin  https://github.com/Dasovon/aim_report_automation.git (fetch)
origin  https://github.com/Dasovon/aim_report_automation.git (push)
```

---

## 🧱 6. Repository Structure Overview

```
aim_report_automation/
├── VBA/                  → Excel automation macros
├── Python/               → Python automation scripts
│   ├── aim_formatter.py
│   ├── aim_report_automation.py
│   ├── aim_report_tk.py
│   ├── template.xlsm
│   └── venv/ (ignored)
├── .gitignore            → Repo hygiene rules
├── README.md             → Overview and instructions
├── SYNC_GUIDE.md         → This file
└── repo_banner.png       → GitHub visual banner
```

---

## 🏁 7. Quick Reference Commands

| Action | Command |
|--------|----------|
| Clone repo | `git clone https://github.com/Dasovon/aim_report_automation.git` |
| Pull latest | `git pull origin main` |
| Stage all changes | `git add .` |
| Commit with message | `git commit -m "message"` |
| Push to GitHub | `git push origin main` |
| Verify remote | `git remote -v` |
| Test ignores | `git check-ignore -v <file>` |

---

### 💡 Pro Tip
If you regularly switch between systems, make small, frequent commits. It keeps merges clean and reduces risk of conflicts.

---

**Maintained by:** Ryan Novosad  
**Repository:** [Dasovon/aim_report_automation](https://github.com/Dasovon/aim_report_automation)  
**Last Updated:** November 2025

