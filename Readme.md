# VBA installer

## Why?
Excel files are in binary format, so it is hard to track their differences. 

## How?
I propose to distribute a single Excel file called "Installer" that loads sheets, modules, classes from accompanying XML files.

<img src="./img/main.svg">

## Examples
* "Demo 1.xlsm" installs sources and one sheet with implementation of matrix multiplication.

## TODO:
* Add handling of references
* Add handling of cell formatting
* Add handling of user forms

Any contributions (proposals, discussions, pull requests) are welcome. 