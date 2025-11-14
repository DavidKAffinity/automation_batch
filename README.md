Automation_Art Installation and Usage Guide

**Requires Python 3.11 or newer.**

---

## Running from Server (PowerShell)

### Open VS Code
The initial steps to running must be done in a PowerShell Terminal for now

Press CTRL + ` (Backtick, same button as tilde)
You should see a line that says PS "path/you're/in"
If it does not say PS check the right side of the Terminal, there will be a plus sign and a trash can, click the arrow next to the plus sign and select PowerShell

### Navigate to the automation_art Directory

```powershell
cd "S:/Workstation - DavidK/Projects/automation_batch"
```

### Activate the Virtual Environment

```powershell
.venv\Scripts\Activate.ps1
```

### Open main.py in VS Code and Run it
Double click on main.py inside of "S:/Workstation - DavidK/Projects/automation_batch/srs/automation_batch"
This will cause the code to load into VS Code where you can edit it
There should also be a Play button in the upper right corner
If there is no play button, right-click on the text in the code and go to Run Python -> Run Python File in Terminal