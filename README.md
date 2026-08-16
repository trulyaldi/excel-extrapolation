# Extraplus for Excel

Extraplus provides Excel functions for extrapolating finite numerical data toward an asymptotic or infinite-limit value.

The numerical fitting algorithm is implemented in Python, while [xlwings](https://www.xlwings.org/) makes the functions directly callable from Excel.

Once installed, Extraplus can be used like a normal Excel function:

```excel
=EXTRAP1(A2:A20,B2:B20)
```

Additional functions provide the estimated uncertainty and fitted decay parameter:

```excel
=UNCERTY1(A2:A20,B2:B20)
=PARAMETER_B1(A2:A20,B2:B20)
```

---

# Quick Start

For a new Windows computer:

1. Install **Python 3.11**.
2. Download or clone this repository.
3. Close Excel.
4. Run `setup.bat`.
5. Open `Extraplus.xlsm`.
6. Open the **xlwings** tab in Excel.
7. Set the Python interpreter to the `.venv` created by `setup.bat`.
8. Set **UDF Modules** to:

```text
extrapolation
```

9. Enable **Add Workbook to PYTHONPATH**.
10. Click **Import Python UDFs**.
11. Start using functions such as:

```excel
=EXTRAP1(A2:A20,B2:B20)
```

Detailed installation instructions are provided below.

---

# What Extraplus Does

Extraplus estimates the limiting value

[
C = \lim_{x\rightarrow\infty} y(x)
]

from a finite sequence of calculations.

Three convergence models are currently available.

### Model 1 — Exponential

[
y(x)=C+A\exp\left(-B\frac{x}{x_{\max}}\right)
]

Excel function:

```excel
=EXTRAP1(x_range,y_range)
```

### Model 2 — Square-Root Exponential

[
y(x)=C+A\exp\left(-B\sqrt{\frac{x}{x_{\max}}}\right)
]

Excel function:

```excel
=EXTRAP2(x_range,y_range)
```

### Model 3 — Power Law

[
y(x)=C+A\left(\frac{x}{x_{\max}}\right)^{-B}
]

Excel function:

```excel
=EXTRAP3(x_range,y_range)
```

Here:

* `C` is the extrapolated infinite-limit value.
* `A` controls the magnitude and direction of convergence.
* `B` controls the convergence rate.
* `x_max` is the largest supplied x value.

The x variable is normalized internally by `x_max`.

The y data are also normalized internally for numerical stability. Extrapolated values and uncertainties returned to Excel are automatically transformed back to the original units.

---

# Repository Structure

The project intentionally uses a simple layout:

```text
extraplus-excel/
│
├── extrap.py
├── extrapolation.py
├── Extraplus.xlsm
├── requirements.txt
├── setup.bat
├── README.md
└── .gitignore
```

### `extrap.py`

Contains the actual extrapolation algorithm.

This includes:

* data normalization
* convergence-direction detection
* logarithmic linearization
* asymptote scanning
* linear regression
* parameter estimation
* uncertainty estimation

This file is the numerical core of Extraplus.

### `extrapolation.py`

Contains the Excel interface.

It does **not** duplicate the fitting algorithm.

Its responsibilities are limited to:

* receiving Excel ranges
* cleaning Excel inputs
* converting ranges to Python data
* calling the `extraplus` class from `extrap.py`
* converting results back to original units
* exposing the functions through xlwings
* caching repeated calculations

The dependency therefore looks like:

```text
Excel
   │
   ▼
extrapolation.py
   │
   ▼
extrap.py
   │
   ▼
extraplus
```

Changes to the numerical algorithm should normally be made only in `extrap.py`.

---

# System Requirements

The current Excel integration is intended for:

```text
Windows
Microsoft Excel Desktop
Python 3.11
xlwings
```

An internet connection is required during the initial installation so that Python dependencies can be downloaded.

The Python packages currently required are:

```text
numpy
scipy
pandas
matplotlib
xlwings
```

They are installed automatically by `setup.bat` from `requirements.txt`.

---

# Installation

## 1. Install Python 3.11

Install Python 3.11 on the computer.

During installation, enabling:

```text
Add Python to PATH
```

is recommended.

After installation, open Command Prompt and verify:

```bat
py -3.11 --version
```

You should see something similar to:

```text
Python 3.11.x
```

---

## 2. Download Extraplus

Either clone the repository with Git:

```bat
git clone <repository-url>
```

or use:

```text
GitHub
→ Code
→ Download ZIP
```

If using a ZIP, extract it before continuing.

For example:

```text
C:\Users\YourName\Documents\extraplus-excel\
```

The important files should remain together:

```text
extrap.py
extrapolation.py
requirements.txt
setup.bat
Extraplus.xlsm
```

Do not move `extrap.py` away from `extrapolation.py`.

---

## 3. Close Excel

Before running the installation, close all open Excel windows.

This is important because the setup process installs the xlwings Excel add-in.

---

## 4. Run `setup.bat`

Double-click:

```text
setup.bat
```

The script automatically:

```text
checks Python 3.11
        ↓
creates .venv
        ↓
installs Python dependencies
        ↓
installs xlwings
        ↓
installs the xlwings Excel add-in
```

A private Python environment is created inside:

```text
.venv\
```

You do not need to activate this environment manually.

The `.venv` directory should not be committed to Git.

---

# Excel Configuration

After `setup.bat` finishes, open Excel.

You should see an **xlwings** tab in the Excel ribbon.

If you do not see it, close Excel completely, run:

```bat
.venv\Scripts\xlwings.exe addin install
```

and reopen Excel.

---

## Python Interpreter

In the xlwings ribbon, set **Interpreter** to:

```text
<Extraplus folder>\.venv\Scripts\python.exe
```

For example:

```text
C:\Users\John\Documents\extraplus-excel\.venv\Scripts\python.exe
```

The exact path is also displayed by `setup.bat` after installation.

---

## UDF Modules

Set:

```text
UDF Modules
```

to:

```text
extrapolation
```

Do not include `.py`.

Correct:

```text
extrapolation
```

Incorrect:

```text
extrapolation.py
```

---

## Python Path

Enable:

```text
Add Workbook to PYTHONPATH
```

when the workbook is stored in the same directory as:

```text
extrap.py
extrapolation.py
```

The recommended arrangement is therefore:

```text
extraplus-excel/
├── extrap.py
├── extrapolation.py
└── Extraplus.xlsm
```

---

# Enable UDF Importing in Excel

xlwings needs permission to create the VBA wrappers used by the Python UDFs.

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

Then return to the xlwings ribbon and click:

```text
Import Python UDFs
```

This creates the VBA bridge between Excel and the Python functions in `extrapolation.py`.

---

# Workbook Format

Use a macro-enabled workbook:

```text
.xlsm
```

For example:

```text
Extraplus.xlsm
```

A normal `.xlsx` workbook cannot store the VBA wrappers required by the current xlwings UDF setup.

If you have an existing `.xlsx` workbook:

```text
File
→ Save As
→ Excel Macro-Enabled Workbook (*.xlsm)
```

Then import the Python UDFs into the new `.xlsm` file.

For legacy `.xls` files, save the workbook as `.xlsm` before importing the Python UDFs.

---

# Basic Usage

Suppose your data look like this:

| A — Basis Size | B — Energy |
| -------------: | ---------: |
|              2 |    -262.04 |
|              3 |    -262.30 |
|              4 |    -263.57 |
|              5 |    -263.65 |
|              6 |    -263.67 |
|            ... |        ... |

To extrapolate using the exponential model:

```excel
=EXTRAP1(A2:A20,B2:B20)
```

For the square-root exponential model:

```excel
=EXTRAP2(A2:A20,B2:B20)
```

For the power-law model:

```excel
=EXTRAP3(A2:A20,B2:B20)
```

The returned number is the estimated asymptotic value `C` in the original units of the supplied y data.

---

# Available Excel Functions

## Extrapolated Value

```excel
=EXTRAP1(x_range,y_range)
=EXTRAP2(x_range,y_range)
=EXTRAP3(x_range,y_range)
```

The model mapping is:

| Number | Model                   |
| ------ | ----------------------- |
| `1`    | Exponential             |
| `2`    | Square-root exponential |
| `3`    | Power law               |

---

## Uncertainty

```excel
=UNCERTY1(x_range,y_range)
=UNCERTY2(x_range,y_range)
=UNCERTY3(x_range,y_range)
```

These return the estimated uncertainty in `C`, expressed in the same units as the original y data.

For example:

```excel
=EXTRAP1(A2:A20,B2:B20)
```

might return:

```text
-263.69412
```

while:

```excel
=UNCERTY1(A2:A20,B2:B20)
```

might return:

```text
0.00031
```

which can be interpreted as an extrapolation result of approximately:

```text
-263.69412 ± 0.00031
```

according to the uncertainty model implemented by Extraplus.

---

## Decay Parameter B

The fitted convergence parameter can be returned with:

```excel
=PARAMETER_B1(x_range,y_range)
=PARAMETER_B2(x_range,y_range)
=PARAMETER_B3(x_range,y_range)
```

For example:

```excel
=PARAMETER_B1(A2:A20,B2:B20)
```

Because x is normalized internally, the returned `B` corresponds to the normalized model used by Extraplus.

---

# Fixed B

Normally Extraplus determines `B` from the supplied data.

You can instead provide a fixed positive `B` as the third argument.

For example:

```excel
=EXTRAP1(A2:A20,B2:B20,4.25)
```

or reference another Excel cell:

```excel
=EXTRAP1(A2:A20,B2:B20,$H$2)
```

The same applies to uncertainty and B functions:

```excel
=UNCERTY1(A2:A20,B2:B20,$H$2)
```

```excel
=PARAMETER_B1(A2:A20,B2:B20,$H$2)
```

When a fixed `B` is supplied, `PARAMETER_B1` returns that effective fixed value.

---

# Energy-B Propagation

One intended workflow is to determine `B` from an Energy calculation and reuse it when extrapolating another observable.

For example, suppose:

```text
Column B = Energy
Column C = another observable
```

First determine the Energy decay parameter:

```excel
=PARAMETER_B1(A2:A20,B2:B20)
```

Suppose that formula is stored in:

```text
H2
```

You can then extrapolate the other observable with the same `B`:

```excel
=EXTRAP1(A2:A20,C2:C20,$H$2)
```

and its uncertainty:

```excel
=UNCERTY1(A2:A20,C2:C20,$H$2)
```

If your physical model requires half of the Energy decay parameter, Excel can pass that explicitly:

```excel
=EXTRAP1(A2:A20,C2:C20,$H$2/2)
```

The Excel wrapper deliberately makes this propagation explicit rather than automatically deciding which observable should use `B` or `B/2`.

---

# Restricting the Number of Fitted Points

The fourth argument is `n_fit`.

For example:

```excel
=EXTRAP1(A2:A20,B2:B20,,10)
```

uses the first 10 valid observations.

When specifying `n_fit` while allowing `B` to remain free, leave the third argument blank:

```excel
=EXTRAP1(A2:A20,B2:B20,,10)
```

With both a fixed `B` and `n_fit`:

```excel
=EXTRAP1(A2:A20,B2:B20,$H$2,10)
```

At least three observations are required.

---

# Combined Output Functions

Extraplus can return:

```text
C
uncertainty
B
```

from a single calculation.

This avoids performing three separate UDF calls.

## Vertical Output

```excel
=EXTRAP1_V(A2:A20,B2:B20)
```

returns:

```text
C
uncertainty
B
```

vertically.

For example:

```text
-263.69412
0.00031
4.2578
```

Equivalent functions exist for the other models:

```excel
=EXTRAP2_V(A2:A20,B2:B20)
=EXTRAP3_V(A2:A20,B2:B20)
```

---

## Horizontal Output

```excel
=EXTRAP1_H(A2:A20,B2:B20)
```

returns:

```text
C    uncertainty    B
```

horizontally.

Equivalent functions are:

```excel
=EXTRAP2_H(A2:A20,B2:B20)
=EXTRAP3_H(A2:A20,B2:B20)
```

---

# Advanced Uncertainty Arguments

The uncertainty and combined-output functions provide additional optional parameters.

The full form is:

```excel
=UNCERTY1(
    x_range,
    y_range,
    b_fixed,
    n_fit,
    k_cov,
    q,
    use_pi,
    robust_scale
)
```

The defaults are:

```text
k_cov        = 1.0
q            = 0.84
use_pi       = TRUE
robust_scale = TRUE
```

For normal usage, these values should usually be left at their defaults.

For example:

```excel
=UNCERTY1(A2:A20,B2:B20)
```

is normally sufficient.

An advanced call could look like:

```excel
=UNCERTY1(A2:A20,B2:B20,,10,1,0.84,TRUE,TRUE)
```

The same optional uncertainty parameters are available to the `_V` and `_H` combined-output functions.

---

# Input Requirements

Extraplus automatically ignores paired rows containing blank cells, text, or non-finite numeric values.

The numerical data should nevertheless satisfy the following requirements:

* x and y ranges must contain the same number of cells.
* At least three paired numeric observations are required.
* x must contain at least two distinct values.
* the largest x value cannot be zero.
* y cannot be constant.
* `b_fixed`, when provided, must be positive.
* `n_fit`, when provided, must be at least 3.
* positive basis-size values should be used, particularly for the power-law model.

For best results, provide the data in their physical convergence order.

For example:

```text
smallest basis size
        ↓
larger basis size
        ↓
largest basis size
```

The end of the supplied sequence is treated as the chronological convergence tail.

---

# How the Fit Works

Extraplus uses a hybrid linear/nonlinear fitting strategy.

Instead of fitting `A`, `B`, and `C` simultaneously with a fully nonlinear optimizer, the method isolates the asymptote `C`.

For a candidate value of `C`, the remaining decay model can be transformed into a straight-line relationship.

For the exponential model, for example:

[
y=C+A e^{-Bt}
]

with

[
t=\frac{x}{x_{\max}}
]

becomes, after subtracting the candidate asymptote and taking a logarithm,

[
\ln|y-C|=\ln|A|-Bt.
]

Therefore, once `C` is temporarily fixed, `A` and `B` can be estimated using ordinary linear least squares.

Extraplus then scans possible values of `C`, evaluates the quality of the corresponding linearized fit, and refines the strongest candidate.

Conceptually:

```text
candidate C
    ↓
compute distance from C
    ↓
log transformation
    ↓
linear regression
    ↓
evaluate linearized fit
    ↓
search for best C
```

This isolates the nonlinear part of the problem into a one-dimensional optimization.

---

# Data Scaling

Extraplus internally uses min-max normalization for y:

[
y'=\frac{y-y_{\min}}{y_{\max}-y_{\min}}
]

The fitting algorithm operates on this normalized representation.

After fitting, the asymptote is mapped back to the original units:

[
C_{\text{original}}
===================

y_{\min}
+
(y_{\max}-y_{\min})C_{\text{scaled}}.
]

The uncertainty is similarly rescaled.

This means Excel users always receive values in the same physical units as their original data.

---

# Updating Extraplus

If you update the repository with:

```bat
git pull
```

the action required afterward depends on what changed.

### Only the fitting algorithm changed

If only `extrap.py` changed and the Excel functions themselves were not renamed or modified:

```text
xlwings
→ Restart UDF Server
```

should normally be sufficient.

You can then recalculate the workbook.

### `extrapolation.py` internals changed

If the function names and argument signatures remain unchanged, restarting the UDF server is usually sufficient.

### Excel function names or arguments changed

If a UDF was added, removed, renamed, or its function signature changed:

```text
xlwings
→ Import Python UDFs
```

again.

This updates the VBA wrappers stored inside the workbook.

### Python dependencies changed

Run:

```text
setup.bat
```

again.

The existing `.venv` will be reused and the dependencies in `requirements.txt` will be updated.

---

# Troubleshooting

## "Object required"

If Excel displays:

```text
Object required
```

the problem is usually the Excel/xlwings bridge rather than the extrapolation algorithm.

Press:

```text
Alt + F11
```

then:

```text
Tools
→ References
```

Make sure:

```text
xlwings
```

is checked.

If you see:

```text
MISSING: xlwings
```

remove the broken reference, close Excel, and run:

```bat
.venv\Scripts\xlwings.exe addin install
```

Reopen Excel, enable the xlwings reference, and click:

```text
Import Python UDFs
```

again.

Then use:

```text
Restart UDF Server
```

from the xlwings ribbon.

---

## `#NAME?`

If Excel shows:

```text
#NAME?
```

Excel does not recognize the UDF.

Check that:

```text
UDF Modules = extrapolation
```

and click:

```text
Import Python UDFs
```

again.

Also verify that the workbook is `.xlsm`.

---

## Python module cannot be found

If you receive an error involving:

```text
ModuleNotFoundError
```

verify that these two files remain in the same directory:

```text
extrap.py
extrapolation.py
```

Also verify that:

```text
Add Workbook to PYTHONPATH
```

is enabled when the workbook is stored in that directory.

---

## xlwings tab is missing

Close Excel and run:

```bat
.venv\Scripts\xlwings.exe addin install
```

Then reopen Excel.

---

## Macros are blocked

Windows may block macro-enabled workbooks downloaded from the internet.

Close Excel, right-click:

```text
Extraplus.xlsm
```

and select:

```text
Properties
```

If an **Unblock** option is shown, enable it and apply the change.

Then reopen the workbook.

Only enable macros for copies of Extraplus obtained from a source you trust.

---

## Python 3.11 cannot be found

If `setup.bat` reports:

```text
Python 3.11 was not found
```

verify in Command Prompt:

```bat
py -3.11 --version
```

If that command fails, install Python 3.11 and run `setup.bat` again.

---

## Dependency installation fails

Make sure the computer has an active internet connection.

You can manually retry installation with:

```bat
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

---

# Development

The project deliberately separates the numerical algorithm from the Excel interface.

When developing the extrapolation method:

```text
edit extrap.py
```

When changing how Excel communicates with Python:

```text
edit extrapolation.py
```

Avoid placing fitting logic directly inside `extrapolation.py`.

The intended architecture is:

```text
                     ┌─────────────────┐
                     │      Excel      │
                     └────────┬────────┘
                              │
                         xlwings UDF
                              │
                     ┌────────▼────────┐
                     │ extrapolation.py│
                     │  Excel wrapper  │
                     └────────┬────────┘
                              │
                     ┌────────▼────────┐
                     │    extrap.py    │
                     │ numerical core  │
                     └────────┬────────┘
                              │
                     ┌────────▼────────┐
                     │    extraplus    │
                     └─────────────────┘
```

This ensures there is only one implementation of the fitting algorithm.

---

# Git Guidelines

Do not commit the local Python environment.

The repository's `.gitignore` should contain at least:

```gitignore
.venv/
venv/
__pycache__/
*.pyc
~$*
```

The following files **should** normally be committed:

```text
extrap.py
extrapolation.py
Extraplus.xlsm
requirements.txt
setup.bat
README.md
.gitignore
```

This allows another user to download the repository, run `setup.bat`, and reproduce the required Python environment locally.

---

# Function Reference

| Function       | Output                             | Model                   |
| -------------- | ---------------------------------- | ----------------------- |
| `EXTRAP1`      | Asymptote `C`                      | Exponential             |
| `EXTRAP2`      | Asymptote `C`                      | Square-root exponential |
| `EXTRAP3`      | Asymptote `C`                      | Power law               |
| `UNCERTY1`     | Uncertainty in `C`                 | Exponential             |
| `UNCERTY2`     | Uncertainty in `C`                 | Square-root exponential |
| `UNCERTY3`     | Uncertainty in `C`                 | Power law               |
| `PARAMETER_B1` | Parameter `B`                      | Exponential             |
| `PARAMETER_B2` | Parameter `B`                      | Square-root exponential |
| `PARAMETER_B3` | Parameter `B`                      | Power law               |
| `EXTRAP1_V`    | `C`, uncertainty, `B` vertically   | Exponential             |
| `EXTRAP2_V`    | `C`, uncertainty, `B` vertically   | Square-root exponential |
| `EXTRAP3_V`    | `C`, uncertainty, `B` vertically   | Power law               |
| `EXTRAP1_H`    | `C`, uncertainty, `B` horizontally | Exponential             |
| `EXTRAP2_H`    | `C`, uncertainty, `B` horizontally | Square-root exponential |
| `EXTRAP3_H`    | `C`, uncertainty, `B` horizontally | Power law               |

---

# Example Workflow

For an Energy calculation stored in:

```text
A2:A20 = basis size
B2:B20 = Energy
```

calculate all three exponential-model outputs with:

```excel
=EXTRAP1_V(A2:A20,B2:B20)
```

or calculate them individually:

```excel
=EXTRAP1(A2:A20,B2:B20)
=UNCERTY1(A2:A20,B2:B20)
=PARAMETER_B1(A2:A20,B2:B20)
```

For another property in column `C`, reuse the Energy convergence parameter by storing:

```excel
=PARAMETER_B1(A2:A20,B2:B20)
```

in `H2`, then use:

```excel
=EXTRAP1(A2:A20,C2:C20,$H$2)
```

and:

```excel
=UNCERTY1(A2:A20,C2:C20,$H$2)
```

This keeps the numerical fitting algorithm centralized in Python while allowing the entire workflow to remain accessible from ordinary Excel formulas.

---

# Summary

After the initial installation, normal Extraplus usage is simply:

```excel
=EXTRAP1(...)
=EXTRAP2(...)
=EXTRAP3(...)
```

with optional functions for uncertainty and the fitted convergence parameter.

The Excel layer remains deliberately thin:

```text
Excel → extrapolation.py → extrap.py
```

so improvements to the extrapolation algorithm can be made in one place without rewriting the Excel integration.
