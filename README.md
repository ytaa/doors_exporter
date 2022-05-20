# doors_exporter
A simple python command line tool for exporting of IBM Rational DOORS modules. Currently, the only supported export formats are CSV and XLSX (Microsoft Excel Workbook), but the tool can be easily extended with support of other export formats.

Exporting of DOORS modules is done with the use of a DXL script. The script input is generated based on a configuration file in JSON format. All exported documents are automatically saved and closed - there is no need for manually clicking through all of the exported documents and saving them.

## Requirements
To run this tool you need to satisfy following requirements:
 - Python 3.x - [download](https://www.python.org/downloads/)
 - [XLWings](https://docs.xlwings.org/en/stable/) python package - `pip install xlwings`

## Usage
First you need to prepare a JSON configuration file that provides some basic information about the environment (like path to the DOORS executable) and specifies the list of modules to be exported.

Below you can find a basic configuration file `basic_config.json`. Lines starting with *//* are additional comments - they need to be deleted because comments are not part of standard JSON format. For valid example configuration file see [basic_config.json](examples/basic_config.json).

```json
{
	// Path to IBM Rational DOORS application, optional - default value: 'C:/Program Files/IBM/Rational/DOORS/9.7/bin/doors.exe'
	"doors_path":"C:/Program Files/IBM/Rational/DOORS/9.7/bin/doors.exe",
	// Command line arguments passed to the DOORS application, optional - default value: ''
	"doors_options":"-osUser -o r -O r -d 36677@server",
	// Format of exported documents, optional - default value 'xlsx'
	"format": "csv",
	// Execution time limit for DXL script, optional - default value '0' indicating no limit
	"time_limit": "0",
	// DXL script encoding, optional - default value 'utf-8'
	"encoding": "utf-8",
	// List of DOORS modules to be exported, required, at least one module has to be specified
	"modules":[
		// Each object represents a single DOORS module
		{
			// Path to the first DOORS module, required
			"path":"/path/to/example/folder one/example module one",
			// DOORS module view to be used for export, optional - default value: ''
			// In case of the default value, the default view will be used for the export
			"view":"example view"
		},
		{
			// Path to the second DOORS module, required
			"path":"/path/to/example/folder two/example module two"
			// No view specified - defulat view will be used
		}
	]
 }
```
Now, to export the specified DOORS modules, run the `doors_exporter.py` script and provide path to the configuration file as argument:
```
> python doors_exporter.py basic_config.json
```
After the script execution is finished, exported DOORS modules will be available in the `out/` directory.  
