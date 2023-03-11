Attribute VB_Name = "Installer"
Option Explicit

Const MainXmlFile As String = "main.xml"
Const BackupDirectory As String = "\Demo 1 files\"
Const InstallerName As String = "Installer"

Sub Install()
    Dim vbResult As VbMsgBoxResult, result As Boolean, listOfSheets$, queriesSheet As IXMLDOMElement
    Dim mainXml As New MSXML2.DOMDocument60, nd As IXMLDOMElement, newWorkSheet As MSXML2.DOMDocument60
    Dim nd2 As IXMLDOMElement, nd3 As IXMLDOMElement    ' iterators, no need to describe them
    Dim ws As Worksheet, shp As Button, importedFiles$(), listOfFiles$, i&
    Dim sheets As New Collection, sheetToInstall As Variant
    
    vbResult = MsgBox("Would you like to install sheets and modules?", vbYesNo, InstallerName)
    
    If vbResult = vbNo Then
        Exit Sub
    End If
    
    If Not VBATrusted() Then
        vbResult = MsgBox("VBA is not in trusted mode. You have two possibilities:" & vbCrLf & _
                          "1: Switch trusted mode ""On""" & vbCrLf & _
                          "(File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust Access...)" & vbCrLf & _
                          "and run the script again." & vbCrLf & _
                          "2: Install sheets and afterwards add modules manually." & vbCrLf & _
                          "Would you like to proceed with option 2?", vbYesNo, InstallerName)
        If vbResult = vbNo Then
            Set sheets = Nothing
            Exit Sub
        End If
    End If
    
    result = mainXml.Load(ThisWorkbook.Path & BackupDirectory & MainXmlFile)
    If Not result Then
        Call MsgBox("Error at loading of " & ThisWorkbook.Path & BackupDirectory & MainXmlFile, vbOKOnly, InstallerName)
        Set sheets = Nothing
        Exit Sub
    End If
    
    ' Add modules
    If VBATrusted() Then
        importedFiles = ImportModules
        listOfFiles = ""
        For i = 1 To UBound(importedFiles)
            listOfFiles = listOfFiles & importedFiles(i) & vbCrLf
        Next i
        Call MsgBox("Following files were imported:" & vbCrLf & listOfFiles)
    End If
    
    listOfSheets = ""
    For Each nd In mainXml.DocumentElement.SelectNodes("/WorkBook/WorkSheets")
        For Each nd2 In nd.ChildNodes
            Select Case LCase(nd2.BaseName)
                Case "worksheet"
                    sheets.Add nd2.getAttribute("Path")
                    listOfSheets = listOfSheets & nd2.getAttribute("Path") & vbCrLf
                Case "if"
                    vbResult = Condition(nd2)
                    For Each nd3 In nd2.ChildNodes
                        If nd3.nodeName = "True" And vbResult = vbYes Then
                            sheets.Add nd3.getAttribute("Path")
                            listOfSheets = listOfSheets & nd3.getAttribute("Path") & vbCrLf
                        End If
                        If nd3.nodeName = "False" And vbResult = vbNo Then
                            sheets.Add nd3.getAttribute("Path")
                            listOfSheets = listOfSheets & nd3.getAttribute("Path") & vbCrLf
                        End If
                    Next nd3
            End Select
        Next nd2
    Next nd
    
    MsgBox ("Following sheets will be created:" & vbCrLf & listOfSheets)
    
    For Each sheetToInstall In sheets
        Set newWorkSheet = New MSXML2.DOMDocument60
        result = newWorkSheet.Load(ThisWorkbook.Path & BackupDirectory & sheetToInstall)
        If Not result Then
            Call MsgBox("Error at loading of " & ThisWorkbook.Path & BackupDirectory & sheetToInstall, vbOKOnly, InstallerName)
            Set sheets = Nothing
            Exit Sub
        End If
        Set queriesSheet = newWorkSheet.DocumentElement.SelectNodes("/WorkSheet").Item(0)
        
        If Not SheetExists(queriesSheet.getAttribute("Name")) Then
            With ThisWorkbook
                Set ws = .sheets.Add(After:=.sheets(.sheets.count))
                ws.Name = queriesSheet.getAttribute("Name")
            End With
        Else
            Set ws = ThisWorkbook.Worksheets(queriesSheet.getAttribute("Name"))
        End If
        
        For Each nd2 In queriesSheet.ChildNodes
            Select Case LCase(nd2.BaseName)
                Case "cell"     ' Create cells and elements
                    Call SetCell(ws, nd2)
                Case "range"    ' Range
                    Call SetRange(ws, nd2, ws.Range(nd2.getAttribute("Range")))
                Case "shape"    ' Create buttons
                    Set shp = ws.Buttons.Add(CDbl(nd2.getAttribute("Left")), CDbl(nd2.getAttribute("Top")), CDbl(nd2.getAttribute("Width")), CDbl(nd2.getAttribute("Height")))
                    With shp
                      .OnAction = nd2.getAttribute("Macro")
                      .caption = nd2.getAttribute("Text")
                    End With
                Case "run"
                    Call Run(nd2.getAttribute("Function"))
                Case "if"
                    vbResult = Condition(nd2)
                    For Each nd3 In nd2.ChildNodes
                        If nd3.nodeName = "True" And vbResult = vbYes Then
                            Run (nd3.getAttribute("Function"))
                        End If
                        If nd3.nodeName = "False" And vbResult = vbNo Then
                            Call Run(nd3.getAttribute("Function"))
                        End If
                    Next nd3
            End Select
        Next nd2
    Next sheetToInstall
    Set sheets = Nothing
End Sub
Function SheetExists(sheetToFind As String) As Boolean
    Dim ws As Worksheet
    
    SheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If sheetToFind = ws.Name Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function
Function Condition(nd As IXMLDOMElement) As VbMsgBoxResult
    Dim nd2 As IXMLDOMElement, messageTxt$, captionTxt$
    
    If Not IsNull(nd.getAttribute("Message")) Then
        messageTxt = nd.getAttribute("Message")
    Else
        messageTxt = ""
    End If
    If Not IsNull(nd.getAttribute("Message")) Then
        captionTxt = nd.getAttribute("Caption")
    Else
        captionTxt = ""
    End If
    
    Condition = MsgBox(messageTxt, vbYesNo, captionTxt)
End Function
Sub SetCell(ws As Worksheet, nd As IXMLDOMElement)
    Dim wsRange As Range
    
    Set wsRange = ws.Range(ws.Cells(CInt(nd.getAttribute("Row")), CInt(nd.getAttribute("Column"))).address)
    Call SetRange(ws, nd, wsRange)
End Sub
Sub SetRange(ws As Worksheet, nd As IXMLDOMElement, wsRange As Range)
    Dim nd2 As IXMLDOMElement
    
    If Not IsNull(nd.getAttribute("Value")) Then
        wsRange = nd.getAttribute("Value")
    End If
    
    If Not IsNull(nd.getAttribute("HorizontalAlignment")) Then
        wsRange.HorizontalAlignment = String2HorizontalAlignment(nd.getAttribute("HorizontalAlignment"))
    End If
    
    For Each nd2 In nd.ChildNodes
        If LCase(nd2.BaseName) = "font" Then
            With wsRange.Font
                If Not IsNull(nd2.getAttribute("Color")) Then
                    .Color = CLng(nd2.getAttribute("Color"))
                End If
                If Not IsNull(nd2.getAttribute("Bold")) Then
                    .Bold = String2Boolean(nd2.getAttribute("Bold"))
                End If
            End With
        Else
            With wsRange.Borders(String2BordersIndex(nd2.BaseName))
                If Not IsNull(nd2.getAttribute("LineStyle")) Then
                    .LineStyle = String2LineStyle(nd2.getAttribute("LineStyle"))
                End If
                If Not IsNull(nd2.getAttribute("ColorIndex")) Then
                    .ColorIndex = CLng(nd2.getAttribute("ColorIndex"))
                End If
                If Not IsNull(nd2.getAttribute("TintAndShade")) Then
                    .TintAndShade = CLng(nd2.getAttribute("TintAndShade"))
                End If
                If Not IsNull(nd2.getAttribute("Weight")) Then
                    .Weight = String2BorderWeight(nd2.getAttribute("Weight"))
                End If
            End With
        End If
    Next nd2
End Sub

Sub DeleteInstallerSheet()
    Application.DisplayAlerts = False
    ThisWorkbook.sheets(InstallerName).Delete
    Application.DisplayAlerts = True
End Sub

Private Function VBATrusted() As Boolean
    On Error Resume Next
    VBATrusted = (Application.VBE.VBProjects.count) > 0
End Function

Sub ExportSources()
    Dim i&, exportedFiles$(), listOfFiles$
    
    exportedFiles = ExportModules(BackupDirectory, InstallerName, True)
    listOfFiles = ""
    For i = 1 To UBound(exportedFiles)
        listOfFiles = listOfFiles & exportedFiles(i) & vbCrLf
    Next i
    Call MsgBox("Following files were exported:" & vbCrLf & listOfFiles)
End Sub

Private Function ExportModules(BackupDirectory$, InstallerName$, backupInstaller As Boolean) As String()
    Dim VBComp, VBMod, exportedFiles$()
    
    ReDim exportedFiles(0)
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        Set VBMod = VBComp.CodeModule
        If Not (VBComp.Name = InstallerName And Not backupInstaller) Then
            Select Case VBComp.Type
                Case 1  ' vbext_ct_StdModule
                    VBComp.Export ThisWorkbook.Path & BackupDirectory & VBComp.Name & ".bas"
                    ReDim Preserve exportedFiles(UBound(exportedFiles) + 1)
                    exportedFiles(UBound(exportedFiles)) = VBComp.Name & ".bas"
                Case 2  ' vbext_ct_ClassModule
                    VBComp.Export ThisWorkbook.Path & BackupDirectory & VBComp.Name & ".cls"
                    ReDim Preserve exportedFiles(UBound(exportedFiles) + 1)
                    exportedFiles(UBound(exportedFiles)) = VBComp.Name & ".cls"
                Case 3  ' vbext_ct_UserForm
                    VBComp.Export ThisWorkbook.Path & BackupDirectory & VBComp.Name & ".frm"
                    ReDim Preserve exportedFiles(UBound(exportedFiles) + 1)
                    exportedFiles(UBound(exportedFiles)) = VBComp.Name & ".cls"
            End Select
        End If
    Next VBComp
    
    ExportModules = exportedFiles
    Set VBComp = Nothing: Set VBMod = Nothing
End Function

Function ImportModules() As String()
    Dim cmpComponents, file$, importedFiles$()
    
    ' Get the path to the folder with modules
    If Dir(ThisWorkbook.Path & BackupDirectory) = "" Then
        MsgBox "Import Folder not exist"
        Exit Function
    End If
    Set cmpComponents = ThisWorkbook.VBProject.VBComponents
    
    ReDim importedFiles(0)
    file = Dir(ThisWorkbook.Path & BackupDirectory)
    While (file <> "")
        If (InStr(file, ".cls") > 0 Or InStr(file, ".bas") Or InStr(file, ".frm") > 0) And file <> "Installer.bas" Then
            cmpComponents.Import ThisWorkbook.Path & BackupDirectory & file
            ReDim Preserve importedFiles(UBound(importedFiles) + 1)
            importedFiles(UBound(importedFiles)) = file
        End If
        file = Dir
    Wend
    
    ImportModules = importedFiles
    Set cmpComponents = Nothing
End Function

' Conversions of strings to Excel types:
Function String2BordersIndex(inputString$) As XlBordersIndex
    Select Case LCase(inputString)
        Case LCase("xlEdgeLeft")
            String2BordersIndex = xlEdgeLeft
        Case LCase("xlEdgeTop")
            String2BordersIndex = xlEdgeTop
        Case LCase("xlEdgeBottom")
            String2BordersIndex = xlEdgeBottom
        Case LCase("xlEdgeRight")
            String2BordersIndex = xlEdgeRight
        Case LCase("xlDiagonalUp")
            String2BordersIndex = xlDiagonalUp
        Case LCase("xlDiagonalDown")
            String2BordersIndex = xlDiagonalDown
        Case LCase("xlInsideHorizontal")
            String2BordersIndex = xlInsideHorizontal
        Case LCase("xlInsideVertical")
            String2BordersIndex = xlInsideVertical
    End Select
End Function
Function String2LineStyle(inputString$) As XlLineStyle
    Select Case LCase(inputString)
        Case LCase("xlContinuous")
            String2LineStyle = xlContinuous
        Case LCase("xlDash")
            String2LineStyle = xlDash
        Case LCase("xlDashDot")
            String2LineStyle = xlDashDot
        Case LCase("xlDashDotDot")
            String2LineStyle = xlDashDotDot
        Case LCase("xlDot")
            String2LineStyle = xlDot
        Case LCase("xlDouble")
            String2LineStyle = xlDouble
        Case LCase("xlLineStyleNone")
            String2LineStyle = xlLineStyleNone
        Case LCase("xlSlantDashDot")
            String2LineStyle = xlSlantDashDot
    End Select
End Function
Function String2BorderWeight(inputString$) As XlBorderWeight
    Select Case LCase(inputString)
        Case LCase("xlHairline")
            String2BorderWeight = xlHairline
        Case LCase("xlMedium")
            String2BorderWeight = xlMedium
        Case LCase("xlThick")
            String2BorderWeight = xlThick
        Case LCase("xlThin")
            String2BorderWeight = xlThin
    End Select
End Function
Function String2HorizontalAlignment(inputString$) As Long
    Select Case LCase(inputString)
        Case LCase("xlLeft")
            String2HorizontalAlignment = xlLeft
        Case LCase("xlRight")
            String2HorizontalAlignment = xlRight
        Case LCase("xlCenter")
            String2HorizontalAlignment = xlCenter
    End Select
End Function
Function String2Boolean(inputString$) As Boolean
    Select Case LCase(inputString)
        Case LCase("true")
            String2Boolean = True
        Case LCase("false")
            String2Boolean = False
    End Select
End Function
' Automatically gets size and position of shapes on the sheet to avoid experimental tries to find the best position and size
Sub ShapeHelper()
    Dim ws As Worksheet, shp As Shape, i&
    
    Set ws = ThisWorkbook.Worksheets(InstallerName)
    
    i = 1
    For Each shp In ws.Shapes
        ws.Cells(i, 1) = "Top:"
        ws.Cells(i, 2) = shp.Top
        i = i + 1
        ws.Cells(i, 1) = "Left:"
        ws.Cells(i, 2) = shp.Left
        i = i + 1
        ws.Cells(i, 1) = "Width"
        ws.Cells(i, 2) = shp.Width
        i = i + 1
        ws.Cells(i, 1) = "Height"
        ws.Cells(i, 2) = shp.Height
        i = i + 1
    Next shp
End Sub
