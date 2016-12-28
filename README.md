Reads a folder of Speccy XML system reports and creates a summary table as HTML file. Useful to create a simple inventory for a bunch of Windows machines. 

```
usage: specreport.py [-h] [-o <outfile>] [-x] [-c] [-t] [-j] xmlfiles

Create a summary report from multiple Speccy XML files

positional arguments:
  xmlfiles              Speccy XML file or folder of files

optional arguments:
  -h, --help            show this help message and exit
  -o <outfile>, --outfile <outfile>
                        Output file base name (default: report.*)
  -x, --excel           Output summary as XLSX file.
  -c, --csv             Output summary as CSV file.
  -t, --html            Output summary as HTML file.
  -j, --json            Output summary as JSON file.
```

Work in progress! ;-)
