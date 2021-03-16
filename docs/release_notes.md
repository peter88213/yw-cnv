* Download the `.oxt` file.

* Install it using the LibreOffice extension manager.

* After installation (and Office restart) you find a new "yWriter Import/Export" submenu in the "Files" menu.

* If no additional "yWriter Import/Export" submenu shows up in the "Files" menu, please look at the "Tools" > "Extensions" menu.

## Requirements

* LibreOffice 5.4 or more recent.  __Please note:__  _This extension can not be installed on OpenOffice. An OpenOffice variant can be found here: https://github.com/peter88213/pywoo_
* Java Runtime Environment (LibreOffice needs it for macro execution).
* yWriter 7. 

## Changelog

### v0.37.4 Support ods spreadsheets

Change scene/plot list import (advanced features) from csv to ods file format.

Based on PyWriter v2.12.3

### v0.37.0 Import to ods spreadsheets

Change character/location/item list import from csv to ods file format.

Based on PyWriter v2.11.0

### v0.36.0 Import/export csv lists

Import/export character/location/item lists to/from Calc spreadsheets. Rows may be added, deleted and re-ordered.

Based on PyWriter v2.10.0

### v0.35.0 Underline and strikethrough no longer supported

That is because a real support would require considering nesting and multi paragraph formatting. That would make everything too complicated, considering that such formatting is not common in a fictional prose text.

Based on PyWriter v2.9.0

### v0.34.0

Delete the temporary file unconditionally after execution.
Process all yWriter formatting tags.
* Convert underline and strikethrough.
* Discard alignment.
* Discard highlighting.

Based on PyWriter v2.8.0

### v0.33.1

* Refactor and update docstrings.
* Work around a yWriter 7.1.1.2 bug.

Based on PyWriter v2.7.2

### v0.33.0

* Update UI application context.
* Set a blank line as scene divider template.

Based on PyWriter v2.7.0

### v0.32.3

Work around a bug found in yWriter 7.1.1.2 assigning invalid viewpoint
characters to scenes created by splitting.

### v0.32.2

* Refactor the code for better maintainability.

Based on PyWriter v2.6.1

### v0.31.1

* Add strict project structure check.
* Improve screen output.
* Do not indent a chapter's first scene even if marked "append to previous".
* Can now write complete yw5 Files.
* Convert work in progress that contains empty chapter titles.
* Fix location descriptions export.
* Refactor the code for better maintainability.

Based on PyWriter v2.5.1

### v0.30.0

Adapt to modified yw7 file format (yWriter 7.0.7.2+):
* "Info" chapters are replaced by "Notes" chapters (always unused).
* New "Todo" chapter type (always unused).
* Distinguish between "Notes scene", "Todo scene" and "Unused scene".
* Chapter/scene tag colors in "proofread" export correspond to those of
the yWriter chapter list.

Bugfix:
* Suppress chapter title if required.

Based on PyWriter v2.2.0

### v0.29.6

ODT export: Begin appended scenes with first-line-indent style.
Based on PyWriter v2.1.4

### v0.29.5

Adapt to yWriter 7.0.4.9 Beta: Don't replace dashes any longer by "safe" double hyphens when writing yw7 (PyWriter library v2.1.3).

### v0.29.4

Rewrite large parts of the code (PyWriter library v2.1.0).
Support author's comments: Text /* commented out */ in yWriter scenes is
exported as comment and vice versa.

### v0.28.2

Fix a bug making the output filename lowercase.