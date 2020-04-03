# Properties to Excel

Command line application for converting between multiple properties files and Excel.

Useful for management of localization files.

**Pull requests are always welcome.**

## Usage

```
java -jar properties2excel-cli.jar --help

usage: java -jar properties-to-excel.jar -e <arg> [-f] [-h] -p <arg> [-t]
       [-x <arg>]
 -e,--excel-file <arg>          Path to Excel file
 -f,--from-excel                Convert from Excel file to properties
                                files
 -h,--help                      Display usage
 -p,--properties-folder <arg>   Path to folder with properties files
 -t,--to-excel                  Convert from properties files to Excel
                                file
 -x,--extension <arg>           Property file extension (default is
                                'properties')
```

### Sample usage

Convert properties files in folder "properties" to Excel file "excel.xlsx":

```
java -jar properties2excel-cli.jar --to-excel --excel-file excel.xlsx --properties-folder properties
```

Convert Excel file "excel.xlsx" to properties files in folder "properties":

```
java -jar properties2excel-cli.jar --from-excel --excel-file excel.xlsx --properties-folder properties
```

## Project structure

* properties2excel-core - Core library, containing the convert logic
* properties2excel-cli - Command line interface for the core library
