# VBA installer

## Why?
Excel files are in binary format, so it is hard to track their differences. 

## How?
I propose to distribute a single Excel file called "Installer" that loads sheets, modules, classes from accompanying XML files.

<div hidden>
@startuml
cloud GitHub
GitHub --> Installer : clone
GitHub <-- ChangedTable #black;line.dotted : commit / pull request
package Installer {
  usecase [Table] as Table1
  usecase [Classes, modules, user frames] as Classes1
  usecase [XMLs describing all sheets] as XMLs1
}
package ChangedTable {
  usecase [Classes, modules, user frames] as UC2
  usecase [XMLs describing all sheets] as UC3
}
@enduml
</div>
<img src="./img/main.svg">

## TODO:
* Create a "Publish" functionality that updates the XML for submission.
* Add handling of references
* Add handling of cell formatting
* Add handling of user forms

Any contributions (proposals, discussions, pull requests) are welcome. 