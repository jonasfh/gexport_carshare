# gexport_carshare
Export functions for files for Bildeleringen

This software is an google sheets script for exporting invoice-data from
Autopass in an excel-format, to a JSON - format importable by the letsgo
car sharing software.

Install instructions
--------------------

* Create a new google docs spreadsheet.
* Click on Tools -> Script editor to launch script editor
* In Code.gs: copy all code from export.gs
* Make new file sidebar.html and copy content from sidebar.html to this
* Setup Should be in a sheet called SETUP
* SETUP - sheet is created if it dont exist
** Keys in column A
** Values in column B
** Starting on row 5
* Keys are:
** json - id for folder where json files end up
** autopass - id for folder where autopass-files are placed
* The folders will be created if they dont exist
* Reports go in a sheet called "STATS". This sheet will be created if not exists
