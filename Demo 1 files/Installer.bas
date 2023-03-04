Attribute VB_Name = "Installer"
Option Explicit

Const MainXmlFile As String = "main.xml"
Const backupDirectory As String = "\Demo 1 files\"

Sub Install()
    Dim vbResult As VbMsgBoxResult, result As Boolean, listOfSheets$, queriesSheet As IXMLDOMElement
    Dim mainXml As New MSXML2.DOMDocument60, nd As IXMLDOMElement, nd2 As IXMLDOMElement, newWorkSheet As MSXML2.DOMDocument60
    Dim ws As Worksheet, shp As Button, importedFiles$(), listOfFiles$, i&
    
    vbResult = MsgBox("Would you like to install sheets and modules?", vbYesNo, "Installer")
    
    If vbResult = vbNo Then
        Exit Sub
    End If
    
    If Not VBATrusted() Then
        vbResult = MsgBox("VBA is not in trusted mode. You have two possibilities:" & vbCrLf & _
                          "1: Switch trusted mode ""On""" & vbCrLf & _
                          "(File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust Access...)" & vbCrLf & _
                          "and run the script again." & vbCrLf & _
                          "2: Install sheets and afterwards add modules manually." & vbCrLf & _
                          "Would you like to proceed with option 2?", vbYesNo, "Installer")
        If vbResult = vbNo Then
            Exit Sub
        End If
    End If
    
    result = mainXml.Load(ThisWorkbook.Path & backupDirectory & MainXmlFile)
    If Not result Then
        Call MsgBox("Error at loading of " & ThisWorkbook.Path & backupDirectory & MainXmlFile, vbOKOnly, "Installer")
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
    For Each nd In mainXml.DocumentElement.SelectNodes("/WorkBook/WorkSheets/WorkSheet")
        Set newWorkSheet = New MSXML2.DOMDocument60
        result = newWorkSheet.Load(ThisWorkbook.Path & backupDirectory & nd.getAttribute("Path"))
        If Not result Then
            Call MsgBox("Error at loading of " & ThisWorkbook.Path & backupDirectory & nd.getAttribute("Path"), vbOKOnly, "Installer")
            Exit Sub
        End If
        Set queriesSheet = newWorkSheet.DocumentElement.SelectNodes("/WorkSheet").Item(0)
        listOfSheets = listOfSheets & queriesSheet.getAttribute("Name") & vbCrLf
    Next nd
    
    MsgBox ("Following sheets will be created:" & vbCrLf & listOfSheets)
    
    For Each nd In mainXml.DocumentElement.SelectNodes("/WorkBook/WorkSheets/WorkSheet")
        Set newWorkSheet = New MSXML2.DOMDocument60
        result = newWorkSheet.Load(ThisWorkbook.Path & backupDirectory & nd.getAttribute("Path"))
        If Not result Then
            Call MsgBox("Error at loading of " & ThisWorkbook.Path & backupDirectory & nd.getAttribute("Path"), vbOKOnly, "Installer")
            Exit Sub
        End If
        Set queriesSheet = newWorkSheet.DocumentElement.SelectNodes("/WorkSheet").Item(0)
        
        With ThisWorkbook
            Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            ws.Name = queriesSheet.getAttribute("Name")
        End With
        
        For Each nd2 In queriesSheet.ChildNodes
            Select Case LCase(nd2.BaseName)
                Case "cell"     ' Create cells and elements
                    Call SetCell(ws, nd2)
                Case "range"    ' Range
                    Call SetRange(ws, nd2)
                Case "shape"    ' Create buttons
                    Set shp = ws.Buttons.Add(CDbl(nd2.getAttribute("Left")), CDbl(nd2.getAttribute("Top")), CDbl(nd2.getAttribute("Width")), CDbl(nd2.getAttribute("Height")))
                    With shp
                      .OnAction = nd2.getAttribute("Macro")
                      .Caption = nd2.getAttribute("Text")
                    End With
                Case "run"
                    Run (nd2.getAttribute("Function"))
            End Select
        Next nd2
    Next nd
End Sub

Sub SetCell(ws As Worksheet, nd As IXMLDOMElement)
    ws.Cells(CInt(nd.getAttribute("Row")), CInt(nd.getAttribute("Column"))) = nd.getAttribute("Value")
    'Call SetRange(ws, nd)   ' do formatting
End Sub
Sub SetRange(ws As Worksheet, nd As IXMLDOMElement)
    Dim nd2 As IXMLDOMElement
    
    If Not IsNull(nd.getAttribute("Value")) Then
        ws.Range(nd.getAttribute("Range")) = nd.getAttribute("Value")
    End If
    
'    For Each nd2 In nd.ChildNodes
'        If Not IsNull(nd.getAttribute("xlEdgeLeft")) Then
'            ws.Range().Borders(xlEdgeLeft).LineStyle = xlContinuous
'            With Selection.Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .ColorIndex = 0
'                .TintAndShade = 0
'                .Weight = xlMedium
'            End With
'        End If
'        Select Case LCase(nd2.BaseName)
'            Case "cell"     ' Create cells and elements
'                Call SetCell(ws, nd2)
'            Case "range"    ' Range
'                Call SetRange(ws, nd2)
'            Case "shape"    ' Create buttons
'                Set shp = ws.Buttons.Add(CDbl(nd2.getAttribute("Left")), CDbl(nd2.getAttribute("Top")), CDbl(nd2.getAttribute("Width")), CDbl(nd2.getAttribute("Height")))
'                With shp
'                  .OnAction = nd2.getAttribute("Macro")
'                  .Caption = nd2.getAttribute("Text")
'                End With
'            Case "run"
'                Run (nd2.getAttribute("Function"))
'        End Select
'        With ws.Borders
'        Wend
'    Next nd2
'    If nd.getAttribute("Row") = 1 Then
'    End If
'    <xlEdgeLeft LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
'        <xlEdgeTop LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
'        <xlEdgeBottom LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
'        <xlEdgeRight LineStyle = "xlContinuous" ColorIndex = "0" TintAndShade = "0" Weight = "xlMedium" />
End Sub

Sub DeleteInstallerSheet()
    Dim vbResult As VbMsgBoxResult
    
    vbResult = MsgBox("Do you want to remove ""Installer"" sheet?", vbYesNo, "Installer")
    If vbResult = vbYes Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets("Installer").Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Function VBATrusted() As Boolean
    On Error Resume Next
    VBATrusted = (Application.VBE.VBProjects.Count) > 0
End Function

Sub ExportSources()
    Dim i&, exportedFiles$(), listOfFiles$
    
    exportedFiles = ExportModules(backupDirectory, "Installer", True)
    listOfFiles = ""
    For i = 1 To UBound(exportedFiles)
        listOfFiles = listOfFiles & exportedFiles(i) & vbCrLf
    Next i
    Call MsgBox("Following files were exported:" & vbCrLf & listOfFiles)
End Sub

Private Function ExportModules(backupDirectory$, installerName$, backupInstaller As Boolean) As String()
    Dim VBComp, VBMod, exportedFiles$()
    
    ReDim exportedFiles(0)
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        Set VBMod = VBComp.CodeModule
        If Not (VBComp.Name = installerName And Not backupInstaller) Then
            Select Case VBComp.Type
                Case 1  ' vbext_ct_StdModule
                    VBComp.Export ThisWorkbook.Path & backupDirectory & VBComp.Name & ".bas"
                    ReDim Preserve exportedFiles(UBound(exportedFiles) + 1)
                    exportedFiles(UBound(exportedFiles)) = VBComp.Name & ".bas"
                Case 2  ' vbext_ct_ClassModule
                    VBComp.Export ThisWorkbook.Path & backupDirectory & VBComp.Name & ".cls"
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
    If Dir(ThisWorkbook.Path & backupDirectory) = "" Then
        MsgBox "Import Folder not exist"
        Exit Function
    End If
    Set cmpComponents = ThisWorkbook.VBProject.VBComponents
    
    ReDim importedFiles(0)
    file = Dir(ThisWorkbook.Path & backupDirectory)
    While (file <> "")
        If (InStr(file, ".cls") > 0 Or InStr(file, ".bas") > 0) And file <> "Installer.bas" Then
            cmpComponents.Import ThisWorkbook.Path & backupDirectory & file
            ReDim Preserve importedFiles(UBound(importedFiles) + 1)
            importedFiles(UBound(importedFiles)) = file
        End If
        file = Dir
    Wend
    
    ImportModules = importedFiles
    Set cmpComponents = Nothing
End Function
