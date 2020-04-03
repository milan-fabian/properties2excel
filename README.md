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

### Prequisites

1. Install Java JRE 8 or newer
2. Download newest properties2excel-cli.jar file from [https://github.com/milan-fabian/properties2excel/releases](https://github.com/milan-fabian/properties2excel/releases)

### Sample usage

Convert properties files in folder "properties" to Excel file "excel.xlsx":

```
java -jar properties2excel-cli.jar --to-excel --excel-file excel.xlsx --properties-folder properties
```

Convert Excel file "excel.xlsx" to properties files in folder "properties":

```
java -jar properties2excel-cli.jar --from-excel --excel-file excel.xlsx --properties-folder properties
```

## Format of files

### Properties file

Properties files are expected to be in Java properties format with UTF-8 encoding (be careful with BOM). See following links for more details about the format:
* [https://en.wikipedia.org/wiki/.properties](https://en.wikipedia.org/wiki/.properties)
* [https://docs.oracle.com/cd/E23095_01/Platform.93/ATGProgGuide/html/s0204propertiesfileformat01.html](https://docs.oracle.com/cd/E23095_01/Platform.93/ATGProgGuide/html/s0204propertiesfileformat01.html)
* [https://www.baeldung.com/java-properties](https://www.baeldung.com/java-properties)

All properties files are expected to have exactly the same keys, order of the keys is not important.

### Excel file

* First column contains keys
* First fow contains names of the corresponding properties file (without file extension)

## Information for developers

This application is written in Java and can be compiled using Maven.

### Project structure

* properties2excel-core - Core library, containing the convert logic
* properties2excel-cli - Command line interface for the core library
