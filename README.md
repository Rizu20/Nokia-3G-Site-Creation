# Nokia-3G-Site-Creation
This script creates the NetAct RAML plans from CSV/Excel WO files for the purpose of required MO creations in Nokia 3G site creation.

The file can take either CSV file or Excel file as WO input.
Each CSV input file can obviously contain one class of MO and output will have several XML files for each MO classes
Excel WO input can contain several MO classes (eg. IPNB, WBTS, WCEL, ADJL etc.) and output will be a single XML file with all MO classes.
The sample templates for each of these input files (csv/excel) can be found in Templates directory.

Additionally, there is a Directory input mode, where you can input the directory path containing the csv files of each MO classes and output will be a single XML file with all MO classes.

While running, the script will prompt for input file type (CSV/Excel/Directory). After the appropriate input filetype is given, the script will prompt for WO file or directory path. Then the output RAML/XML file will be stored in the output directory with appropriate MO name and date/time as filename, which can be loaded in NetAct CM Operation Manager as CM plan.

https://imgur.com/T45hg4H




