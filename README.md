# f2p – FlowJo workspace to Excel

Turns a FlowJo workspace (.wsp) into an Excel file:
**Plates** (one 96/384-well grid per gate, one band per group),
**List** (wells down, one column per group) and **Info** (experiment, instrument, panel).

## Install (Mac, once)
1. Unzip this folder anywhere.
2. Right-click `install.command` → **Open** → confirm. (A plain double-click is blocked
   the first time because the file came from someone else.)
3. If macOS offers to install the "Command Line Tools", accept, wait, then run
   `install.command` again.

## Use
Right-click a `.wsp` file (or the folder containing it) in Finder → **Quick Actions → f2p**.
Tick the variables, groups and gates you want → **Make Excel**. The file, named after the experiment, appears next to
the .wsp. In Terminal: `f2p path/to/file.wsp`.

## Settings
Click **Settings…** in the popup (or run `f2p --settings`) to open your personal
settings file: default gates, viability colour bands, gradient colours, default
variables and skipped groups. Save it, then run f2p again. Reinstalling keeps your
settings. Delete the file to go back to the defaults.

## Update
Install the new version the same way. Your settings are kept.

## Notes
- MFI is only available if the workspace contains Median/Mean statistics (add them in FlowJo).
- Wells are read from the `WELL ID` keyword (BD FACSDiva plate exports).
