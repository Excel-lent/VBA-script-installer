# VBA installer

## Why?
Excel files are in binary format, so it is hard to track their differences. 

## How?
I propose to distribute a single Excel file called "Installer" that loads sheets, modules, classes from accompanying XML files.

<img src="./img/main.svg">

## TODO:
* Create a "Publish" functionality that updates the XML for submission.
* Add handling of references
* Add handling of cell formatting
* Add handling of user forms

Any contributions (proposals, discussions, pull requests) are welcome. 