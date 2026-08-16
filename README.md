Extraplus for Excel

Extraplus provides Excel user-defined functions (UDFs) for extrapolating finite numerical data toward an asymptotic or infinite-limit value.

The numerical fitting algorithm lives in extrap.py. The Excel-facing wrapper lives in extrapolation.py and exposes the fitting routines through xlwings.

Once installed, Extraplus behaves like a normal Excel function:

=EXTRAP1(A2:A20,B2:B20)

You can also request the extrapolation uncertainty and fitted decay rate:

=UNCERTAINTY1(A2:A20,B2:B20)
=DECAY_RATE1(A2:A20,B2:B20)

Quick Start

For a new Windows computer:

Install Python 3.11.

Download or clone this repository.

Close all Excel windows.

Run setup.bat.

Open Extraplus.xlsm.

Open the xlwings tab in Excel.

Set Interpreter to the .venv Python created by setup.bat.

Set UDF Modules to extrapolation.

Enable Add Workbook to PYTHONPATH.

Enable Trust access to the VBA project object model in Excel's Trust Center.

Press Alt + F11 to open the VBA editor.

Go to Tools -> References and tick xlwings.

Return to Excel and click Import Python UDFs in the xlwings ribbon.

If needed, click Restart UDF Server.

Test the installation with:

=EXTRAP1(A2:A20,B2:B20)