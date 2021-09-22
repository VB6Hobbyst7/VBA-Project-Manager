Attribute VB_Name = "General"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function workbookOfVbcomponent(vbcomp As VBComponent) As Workbook
    Set workbookOfVbcomponent = WorkbookOfProject(vbcomp.Collection.Parent)
End Function

Function WorkbookOfProject(vbProj As VBProject) As Workbook
    tmpstr = vbProj.fileName
    tmpstr = Right(tmpstr, Len(tmpstr) - InStrRev(tmpstr, "\"))
    Set WorkbookOfProject = Workbooks(tmpstr)
End Function

Public Function WorkbookIsOpen(ByRef wname As String) As Boolean
    Dim wb  As Workbook
    On Error Resume Next
    Set wb = Workbooks(wname)
    If Err.Number = 0 Then WorkbookIsOpen = True
End Function

Sub GotoFirstModule(wb As Workbook)
    Application.VBE.MainWindow.Visible = True
    Application.VBE.MainWindow.WindowState = vbext_ws_Maximize
    Dim vbcomp As VBComponent
    For Each vbcomp In wb.VBProject.VBComponents
        'Debug.Print element
        If vbcomp.Type = vbext_ct_StdModule Then
            vbcomp.Activate
            vbcomp.CodeModule.CodePane.SetSelection 1, 1, 1, 1
            Exit Sub
        End If
    Next vbcomp
End Sub

Function getFilePartPath(fileNameWithExtension, Optional IncludeSlash As Boolean) As String
    getFilePartPath = Left(fileNameWithExtension, InStrRev(fileNameWithExtension, "\") - 1 - IncludeSlash)
End Function

Function getFilePartName(fileNameWithExtension As String, Optional IncludeExtension As Boolean) As String
    If InStr(1, fileNameWithExtension, "\") > 0 Then
        getFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "\"))
    Else
        getFilePartName = fileNameWithExtension
    End If
    If IncludeExtension = False Then getFilePartName = Left(getFilePartName, InStr(1, getFilePartName, ".") - 1)
End Function

Public Function getFilePath(Optional fileType As Variant, Optional multiSelect As Boolean) As Variant
    '1 based (not 0 based)
    ' (1) Shows the msoFileDialogFilePicker dialog box.
    ' (2) Checks if the file type parameter was passed and whether the passed parameter is an array.
    ' (3) Sets the dialog box title, file filter and default file according to the parameters passed.
    ' (4) Returns the paths to the selected files in an array, or displays an error message.
    Dim blArray As Boolean
    Dim i As Long
    Dim strErrMsg As String, strTitle As String
    Dim varItem As Variant
    'check whether the file type parameter was passed
    If Not IsMissing(fileType) Then
        'check whether the passed fileType variable is an array
        blArray = IsArray(fileType)
        ' error
        If Not blArray Then strErrMsg = "Please pass an array in the first parameter of this function!"
    End If
    'proceed
    If strErrMsg = vbNullString Then
        ' set title of dialog box
        If multiSelect Then strTitle = "Choose one or more files" Else strTitle = "Choose file"
        ' set dialog properties
        With Application.FileDialog(msoFileDialogFilePicker)
            .InitialFileName = Environ("USERprofile") & "\Desktop\" 'ThisWorkbook.Path 'Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\"))
            .AllowMultiSelect = multiSelect
            .Filters.Clear
            If blArray Then .Filters.Add "File type", "*." & Join(fileType, ", *.")
            .Title = strTitle
            ' show the file picker dialog box
            If .Show <> 0 Then
                ReDim arrResults(1 To .SelectedItems.Count) As Variant
                ' return multiple file paths
                If blArray Then
                    For Each varItem In .SelectedItems
                        i = i + 1
                        arrResults(i) = varItem
                    Next varItem
                    ' return single file path
                Else
                    arrResults(1) = .SelectedItems(1)
                End If
                ' return results
                getFilePath = arrResults
            End If
        End With
        ' error message
    Else
        MsgBox strErrMsg, vbCritical, "Error!"
    End If
End Function

Public Function ModuleExists(name As String, Optional ByVal ExistsInWorkbook As Workbook) As Boolean
    '!!!need to reference: microsoft visual basic for applications extensibility 5.3
    Dim j As Long
    Dim vbcomp As VBComponent
    Dim modules As Collection
    Set modules = New Collection
    ModuleExists = False
    'check if value is set
    If ExistsInWorkbook Is Nothing Then
        Set ExistsInWorkbook = ThisWorkbook
    End If
    If (name = vbNullString) Then
        GoTo errorname
    End If
    'collect names of files
    For Each vbcomp In ExistsInWorkbook.VBProject.VBComponents
        If ((vbcomp.Type = vbext_ct_StdModule) Or (vbcomp.Type = vbext_ct_ClassModule)) Then
            modules.Add vbcomp.name
        End If
    Next vbcomp
    'Compair the file your looking for to the collection
    For j = 1 To modules.Count
        If (name = modules.Item(j)) Then
            ModuleExists = True
        End If
    Next j
    j = 0
    'if Is_module_loaded not true
    If (ModuleExists = False) Then
        GoTo notfound
    End If
    'if error
    If (0 <> 0) Then
errorname:
        MsgBox ("Function BootStrap.Is_Module_Loaded Was not passed a Name of Module")
        Exit Function
        '   Stop
    End If
    If (0 <> 0) Then
notfound:
        '       MsgBox ("MODULE: " & name & " is not installed please add")
        Exit Function
    End If
End Function

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function ArrayIsAllocated(Arr As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = IsArray(Arr) And _
                                    Not IsError(LBound(Arr, 1)) And _
                                    LBound(Arr, 1) <= UBound(Arr, 1)
End Function

Sub FollowLink(folderPath As String)
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.name = "File Explorer" Then
            If Wnd.Document.Folder.Self.path & "\" = folderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=folderPath, NewWindow:=True
End Sub

Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
    Case vbext_ct_ActiveXDesigner
        ComponentTypeToString = "ActiveX Designer"
    Case vbext_ct_ClassModule
        ComponentTypeToString = "Class Module"
    Case vbext_ct_Document
        ComponentTypeToString = "Document Module"
    Case vbext_ct_MSForm
        ComponentTypeToString = "UserForm"
    Case vbext_ct_StdModule
        ComponentTypeToString = "Code Module"
    Case Else
        ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function

Public Function GetSheetByCodeName(wb As Workbook, CodeName As String) As Worksheet
 Dim sh As Worksheet
    For Each sh In wb.Worksheets                 'Run loop.
        If UCase(sh.CodeName) = UCase(CodeName) Then Set GetSheetByCodeName = sh: Exit For 'Check if it's that sheet or not and set if true
    Next sh
End Function

Sub FoldersCreate(folderPath As String)
    'Create all the folders in a folder path
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant
    'Split the folder path into individual folder names
    individualFolders = Split(folderPath, "\")
    'Loop though each individual folder name
    For Each arrayElement In individualFolders
        'Build string of folder path
        tempFolderPath = tempFolderPath & arrayElement & "\"
        'If folder does not exist, then create it
        If Dir(tempFolderPath, vbDirectory) = "" Then
            MkDir tempFolderPath
        End If
    Next arrayElement
End Sub

Function GetCompText(codeMod As CodeModule) As String
    If codeMod.CountOfLines = 0 Then GetCompTextNew = "": Exit Function
    GetCompText = codeMod.Lines(1, codeMod.CountOfLines)
End Function

Function ProcList(codeMod As CodeModule) As Collection
    Dim coll As Collection
    Set coll = New Collection
    Dim lineNum As Long
    Dim NumLines As Long
    Dim procName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    lineNum = codeMod.CountOfDeclarationLines + 1
    Do Until lineNum >= codeMod.CountOfLines
        procName = codeMod.ProcOfLine(lineNum, ProcKind)
        coll.Add procName
        lineNum = codeMod.ProcStartLine(procName, ProcKind) + codeMod.ProcCountLines(procName, ProcKind) + 1
    Loop
    Set ProcList = coll
End Function

Public Function GetProcText(codeMod As CodeModule, _
                            sProcName As String, _
                            Optional bInclHeader As Boolean = True)
    Dim lProcStart            As Long
    Dim lProcBodyStart        As Long
    Dim lProcNoLines          As Long
    Const vbext_pk_Proc = 0
    On Error GoTo Error_Handler
    lProcStart = codeMod.ProcStartLine(sProcName, vbext_pk_Proc)
    lProcBodyStart = codeMod.ProcBodyLine(sProcName, vbext_pk_Proc)
    lProcNoLines = codeMod.ProcCountLines(sProcName, vbext_pk_Proc)
    If bInclHeader = True Then
        GetProcText = codeMod.Lines(lProcStart, lProcNoLines)
    Else
        lProcNoLines = lProcNoLines - (lProcBodyStart - lProcStart)
        GetProcText = codeMod.Lines(lProcBodyStart, lProcNoLines)
    End If
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    'Err 35 is raiised if proc not found
    Debug.Print "The following error has occurred" & vbCrLf & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Source: GetProcText" & vbCrLf & _
                "Error Description: " & Err.Description & _
                Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Line No: " & Erl)
    Resume Error_Handler_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : Txt_Append
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Output Data to an external file (*.txt or other format)
'               If the file does not exist already it will be created automatically
'               ***Do not forget about access' DoCmd.OutputTo Method for
'               exporting objects (queries, report,...)***
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile     : Name of the file that the text is to be output to including the full path
' sText     : Text to be output to the file
'
' Usage:
' ~~~~~~
' Call Txt_Append("C:\temp\text.txt", "This is a new appended line of text.")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2011-06-16              Initial Public Release
' 2         2018-02-24              Updated Copyright
'                                   Updated error handler
'---------------------------------------------------------------------------------------
Function TxtAppend(sFile As String, sText As String)
    On Error GoTo Err_Handler
    Dim iFileNumber           As Integer
 
    iFileNumber = FreeFile                       ' Get unused file number
    Open sFile For Append As #iFileNumber        ' Connect to the file
    Print #iFileNumber, sText                    ' Append our string
    Close #iFileNumber                           ' Close the file
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Txt_Append" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function
