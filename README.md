# ConvJudgesToCSV

## _Converting Excel file with data on judges to Csv format_

[<img alt="Goland" src="https://img.shields.io/badge/Go-00ADD8?style=float&logo=go&logoColor=white" />](https://golang.org/)
![GitHub release (latest by date)](https://img.shields.io/github/v/release/DFofanov/ConvJudgesToCSV)
[![GitHub license](https://img.shields.io/github/license/DFofanov/ConvJudgesToCSV)](https://github.com/DFofanov/ConvJudgesToCSV/blob/main/LICENSE)

## Description
The program is written to convert Excel file to CSV format.

### Excel file consists of three sheets.
* Sheet1 - Brief information about the judge
* Sheet2 - Rank
* Sheet3 - Group, standard number, breed name

### The resulting CSV file is a text file in which the separator is a semicolon.
Fields of csv file:
* Brief information about the judge
* Rank
* Group, standard number, breed name, competitions


## Usage
An example of using the program:

`ConvJudgesToSCV.exe XlsxFile CsvFile` 


Commands:
* -h (help)
* -v (program version)

![image](https://github.com/DFofanov/ConvJudgesToCSV/blob/main/images/docs.gif?raw=true)

## License
Licensed under the GPL-3.0 License.
