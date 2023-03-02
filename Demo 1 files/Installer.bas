Attribute VB_Name = "Installer"
Option Explicit

Const MainXmlFile As String = "main.xml"
Const backupDirectory As String = "\Demo 1 files\"

Sub Install()
    Dim vbResult As VbMsgBoxResult, result As Boolean, listOfSheets$, queriesSheet As IXMLDOMElement
    Dim mainXml As New MSXML2.DOMDocument60, nd, nd2, newWorkSheet As MSXML2.DOMDocument60, ws As Worksheet
    Dim shp As Button, importedFiles$(), listOfFiles$, i&
    
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
                    ws.Cells(CInt(nd2.getAttribute("Row")), CInt(nd2.getAttribute("Column"))) = nd2.getAttribute("Value")
                Case "shape"    ' Create buttons
                    Set shp = ws.Buttons.Add(CDbl(nd2.getAttribute("Left")), CDbl(nd2.getAttribute("Top")), CDbl(nd2.getAttribute("Width")), CDbl(nd2.getAttribute("Height")))
                    With shp
                      .OnAction = nd2.getAttribute("Macro")
                      .Caption = nd2.getAttribute("Text")
                    End With
            End Select
        Next nd2
    Next nd
    
    ' Add modules
    If VBATrusted() Then
        importedFiles = ImportModules
        listOfFiles = ""
        For i = 1 To UBound(importedFiles)
            listOfFiles = listOfFiles & importedFiles(i) & vbCrLf
        Next i
        Call MsgBox("Following files were imported:" & vbCrLf & listOfFiles)
    End If
    
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
