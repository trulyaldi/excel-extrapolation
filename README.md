# Extraplus for Excel

This repository installs the Extraplus Python UDFs into Microsoft Excel on Windows.

After installation, functions such as the following are available in Excel:

```excel
=EXTRAP1(A2:A20,B2:B20)
=UNCERTAINTY1(A2:A20,B2:B20)
=DECAY_RATE1(A2:A20,B2:B20)
```

## Requirements

* Windows
* Microsoft Excel Desktop
* Python 3.11 or newer
* Internet connection during the initial installation

## Repository Files

Keep these files together in the same folder:

```text
extraplus-excel/
├── extrap.py
├── extrapolation.py
├── Extraplus.xlsm
├── requirements.txt
├── setup.bat
└── README.md
```

---

# Installation

## 1. Install Python 3.11 or Newer

Install **Python 3.11 or newer**.

During installation, enable:

```text
Add Python to PATH
```

To verify the installation, open Command Prompt and run:

```bat
py -3 --version
```

You should see Python 3.11 or newer, for example:

```text
Python 3.11.x
```

or:

```text
Python 3.12.x
```

or:

```text
Python 3.13.x
```

---

## 2. Download Extraplus

Either clone this repository:

```bat
git clone <repository-url>
```

or download it from GitHub:

```text
Code
→ Download ZIP
```

If you download the ZIP, extract it before continuing.

Do not run Extraplus directly from inside the ZIP file.

---

## 3. Close Excel

Close all open Excel windows before continuing.

---

## 4. Run `setup.bat`

Double-click:

```text
setup.bat
```

The script will automatically:

1. Check for Python 3.11 or newer.
2. Create a local `.venv` Python environment.
3. Install the required Python packages.
4. Install the xlwings Excel add-in.

Wait until the script reports that the installation is complete.

---

# Configure Excel

## 5. Open the Workbook

Open:

```text
Extraplus.xlsm
```

You should see an **xlwings** tab in the Excel ribbon.

If the xlwings tab does not appear:

1. Close Excel.
2. Open Command Prompt in the Extraplus folder.
3. Run:

```bat
.venv\Scripts\xlwings.exe addin install
```

4. Reopen Excel.

---

## 6. Set the Python Interpreter

Open the **xlwings** tab.

Set **Interpreter** to:

```text
<Extraplus folder>\.venv\Scripts\python.exe
```

For example:

```text
C:\Users\YourName\Documents\extraplus-excel\.venv\Scripts\python.exe
```

---

## 7. Set the UDF Module

In the xlwings ribbon, set:

```text
UDF Modules
```

to:

```text
extrapolation
```

Do not enter:

```text
extrapolation.py
```

Use only:

```text
extrapolation
```

---

## 8. Enable Workbook PYTHONPATH

In the xlwings ribbon, enable:

```text
Add Workbook to PYTHONPATH
```

This allows Python to find `extrapolation.py` and `extrap.py` in the workbook folder.

---

## 9. Enable VBA Project Access

In Excel go to:

```text
File
→ Options
→ Trust Center
→ Trust Center Settings
→ Macro Settings
```

Enable:

```text
Trust access to the VBA project object model
```

Click **OK**.

---

## 10. Enable the xlwings VBA Reference

Press:

```text
Alt + F11
```

to open the VBA editor.

Then go to:

```text
Tools
→ References
```

Find:

```text
xlwings
```

and tick the checkbox:

```text
☑ xlwings
```

Click **OK**.

This step is required. If the xlwings reference is not enabled, Excel may display:

```text
Object required
```

when using the UDFs.

---

## 11. Import the Python UDFs

Return to Excel.

Open the **xlwings** tab and click:

```text
Import Python UDFs
```

Wait for the import to finish.

Then click:

```text
Restart UDF Server
```

---

# Test the Installation

Suppose:

```text
A2:A20 = x values
B2:B20 = y values
```

Test:

```excel
=EXTRAP1(A2:A20,B2:B20)
```

You can also test:

```excel
=UNCERTAINTY1(A2:A20,B2:B20)
```

and:

```excel
=DECAY_RATE1(A2:A20,B2:B20)
```

The available model families are:

```excel
=EXTRAP1(...)
=EXTRAP2(...)
=EXTRAP3(...)

=UNCERTAINTY1(...)
=UNCERTAINTY2(...)
=UNCERTAINTY3(...)

=DECAY_RATE1(...)
=DECAY_RATE2(...)
=DECAY_RATE3(...)
```

---

# Using Extraplus in Another Workbook

The simplest option is to save the workbook as:

```text
Excel Macro-Enabled Workbook (*.xlsm)
```

Then configure the workbook in the xlwings ribbon:

```text
Interpreter   = <Extraplus folder>\.venv\Scripts\python.exe
UDF Modules   = extrapolation
```

Enable:

```text
Add Workbook to PYTHONPATH
```

Make sure the workbook is in the same folder as:

```text
extrap.py
extrapolation.py
```

Then:

1. Press `Alt + F11`.
2. Go to `Tools → References`.
3. Tick `xlwings`.
4. Return to Excel.
5. Click `Import Python UDFs`.
6. Click `Restart UDF Server`.

---

# Troubleshooting

## `Object required`

Press:

```text
Alt + F11
```

Then:

```text
Tools
→ References
```

Make sure:

```text
☑ xlwings
```

is checked.

After enabling it, return to Excel and run:

```text
Import Python UDFs
```

followed by:

```text
Restart UDF Server
```

---

## xlwings Is Missing from References

Close Excel completely.

Open Command Prompt in the Extraplus folder and run:

```bat
.venv\Scripts\xlwings.exe addin install
```

Reopen Excel.

Then:

```text
Alt + F11
→ Tools
→ References
→ ☑ xlwings
```

---

## `#NAME?`

If Excel displays:

```text
#NAME?
```

check that:

```text
UDF Modules = extrapolation
```

Then click:

```text
Import Python UDFs
```

again.

---

## `ModuleNotFoundError`

Make sure these files are still in the same folder:

```text
Extraplus.xlsm
extrap.py
extrapolation.py
```

Also make sure:

```text
Add Workbook to PYTHONPATH
```

is enabled.

---

## xlwings Tab Is Missing

Close Excel and run:

```bat
.venv\Scripts\xlwings.exe addin install
```

Then reopen Excel.

---

## Macros Are Blocked

If Windows blocks the downloaded workbook:

1. Close Excel.
2. Right-click `Extraplus.xlsm`.
3. Select **Properties**.
4. If **Unblock** appears, tick it.
5. Click **Apply**.
6. Reopen the workbook.

---

# Updating Extraplus

After downloading or pulling a newer version:

```bat
git pull
```

If only the Python fitting code changed, use:

```text
xlwings
→ Restart UDF Server
```

If the Excel UDF names or arguments changed, use:

```text
xlwings
→ Import Python UDFs
```

again.

If `requirements.txt` changed, rerun:

```text
setup.bat
```
