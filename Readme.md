# VBA installer

## Why?
Excel files are in binary format, so it is hard to track their differences. 

## How?
I propose to distribute a single Excel file called "Installer" that loads sheets, modules, classes from accompanying XML files.

<img src="./img/main.svg">

## Structure
For execution the installer needs XML files and modules stored in folder (see constant "backupDirectory"). The main XML file (constant "MainXmlFile", "main.xml") inside the directory contains the names of XML files corresponding to each created sheet.

An example of "main.xml" that contains one sheet ("Sheet1.xml"):
```xml
<WorkBook>
    <WorkSheets>
        <WorkSheet Path="Sheet1.xml" />
    </WorkSheets>
</WorkBook>
```

Each sheet' XML file contains following XML nodes:
* Cell / Range
    * Type (not used yet)
    * Range (string in case of range, see an example below) or Row / Column (longs in case of cell)
    * Value (cell / range can also contain formula)
    * Border lines (see functions "String2BordersIndex", "String2LineStyle", "String2BorderWeight" and example below)
    * Font color (as long) and bold (true / false)
    * HorizontalAlignment (see functions "String2HorizontalAlignment" and example below)
* Shape
    * Type (only "Button" can be used)
    * Left
    * Top
    * Width
    * Height
    * Text
    * Macro (macro name that will be executed on button press)
* Run
    * Function - a function that should be called.

An example of "Sheet1.xml":
```xml
<WorkSheet Name="Matrix Multiplication">
    <Shape Type="Button" Left = "250" Top = "150" Width = "80" Height = "35" Text="Multiply!" Macro = "MatrixMultiplication.MatrixMultiplication" />
    <Cell Type="string" Row = "1" Column = "5" HorizontalAlignment = "xlRight" Value = "Multiplication of matrices:">
        <Font Color = "-16776961" Bold = "True" />
    </Cell>
    <Range Type="int" Range = "G1" Value = "1" />
    <Range Type="formula" Range = "G7" Value = "=G1*G3+H1*G4" />
    <Run Function="DeleteInstallerSheet" />
    <Range Range="F1:F8">
        <xlEdgeLeft LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
        <xlEdgeTop LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
        <xlEdgeBottom LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
        <xlEdgeRight LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
    </Range>
</WorkSheet>
```

## Examples
* "Demo 1.xlsm" installs sources and one sheet with implementation of matrix multiplication.

## TODO:
* Enhance handling of cell formatting
* Add handling of user forms

Any contributions (proposals, discussions, pull requests) are welcome. 