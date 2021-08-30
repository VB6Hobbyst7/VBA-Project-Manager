Attribute VB_Name = "Refreshing"
''''''''''''''''''''''''''''''''''''''''''''
'Ron De Bruin Export - Import Modules START'
''''''''''''''''''''''''''''''''''''''''''''
Sub RefreshComponents(wkbSource As Workbook)
    If wkbSource.name <> ThisWorkbook.name Then
        ExportModules wkbSource
        ImportModules wkbSource
    Else
        MsgBox "Can't run this procedure on myself"
    End If
End Sub

Sub RefreshSelectedBooks()
    'to all selected workbooks
    OptOn
    Dim strFile As String
    Dim Y As Long
    Dim i As Long
    Dim fromWorkbook As Workbook
    Dim varr As Variant

    'create array of selected workbooks for import
    Dim wasOpen As Boolean
    'for each workbook
    Dim ELEMENT As Variant
    For Each ELEMENT In ListboxSelectedIndexes(uCodeArchive.LBooks)
        'For i = LBound(varr) To UBound(varr)
        strFile = uCodeArchive.LBooks.List(ELEMENT, 1)
        wasOpen = False
        'open workbook if closed
        If Not IsWorkBookOpen(uCodeArchive.LBooks.List(ELEMENT, 0)) Then
            On Error GoTo nxt
            Set fromWorkbook = Workbooks.Open(fileName:=strFile, UpdateLinks:=0, ReadOnly:=False)
            fromWorkbook.Windows(1).Visible = False
        Else
            wasOpen = True
            Set fromWorkbook = Workbooks(uCodeArchive.LBooks.List(ELEMENT, 0))
        End If
        '<ACTIONS
        '        Application.DisplayAlerts = False
        If ProtectedVBProject(fromWorkbook) = False And HasProject(fromWorkbook) Then
            Call RefreshComponents(fromWorkbook)
        End If
        '        Application.DisplayAlerts=true
        If wasOpen = False Then
            fromWorkbook.Close savechanges:=True
        End If
nxt:
    Next ELEMENT
    OptOff
End Sub

Public Sub ExportModules(wkbSource As Workbook)
    Dim bExport As Boolean
    '    Dim wkbSource As Excel.Workbook    'change / made wkbSource an arguement
    '    Dim szSourceWorkbook As String     'change / made wkbSource an arguement
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    On Error Resume Next
    Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0
    ''' NOTE: This workbook must be open in Excel.
    '    szSourceWorkbook = ActiveWorkbook.name                     'change / made wkbSource an arguement
    '    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    '    If wkbSource.VBProject.Protection = 1 Then                 'change / will check from caller
    '        MsgBox "The VBA in this workbook is protected," & _
    '               "not possible to export the code"
    '        Exit Sub
    '    End If
    szExportPath = FolderWithVBAProjectFiles & "\"
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        bExport = True
        szFileName = cmpComponent.name
        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
        Case vbext_ct_ClassModule
            szFileName = szFileName & ".cls"
        Case vbext_ct_MSForm
            szFileName = szFileName & ".frm"
        Case vbext_ct_StdModule
            szFileName = szFileName & ".bas"
        Case vbext_ct_Document
            ''' This is a worksheet or workbook object.
            ''' Don't try to export.
            bExport = False
        End Select
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            ''' remove it from the project if you want
            '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        End If
    Next cmpComponent
    '    MsgBox "Export is ready"
End Sub

Public Sub ImportModules(wkbTarget As Workbook)
    'WARNING!
    'DELETES OLD MODULES AND USERFORMS BEFORE IMPORTING NEW
    '    Dim wkbTarget As Excel.Workbook    'change / made wkbTarget an arguement
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    '    Dim szTargetWorkbook As String     'change / made wkbTarget an arguement
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents
    If wkbTarget.name = ThisWorkbook.name Then
        MsgBox "Select another destination workbook" & _
               "Not possible to import in this workbook "
        Exit Sub
    End If
    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If
    ''' NOTE: This workbook must be open in Excel.
    '    szTargetWorkbook = ActiveWorkbook.name                     'change / will check from caller
    '    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    '    If wkbTarget.VBProject.Protection = 1 Then
    '        MsgBox "The VBA in this workbook is protected," & _
    '               "not possible to Import the code"
    '        Exit Sub
    '    End If
    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
        MsgBox "There are no files to import"
        Exit Sub
    End If
    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms(wkbTarget)
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        If (objFSO.GetExtensionName(objFile.name) = "cls") Or _
                                                           (objFSO.GetExtensionName(objFile.name) = "frm") Or _
                                                           (objFSO.GetExtensionName(objFile.name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
    Next objFile
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String
    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")
    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
End Function

Function DeleteVBAModulesAndUserForms(wkbSource As Workbook)
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Set vbProj = wkbSource.VBProject
    For Each vbComp In vbProj.VBComponents
        '    Debug.Print VBComp.Name
        If vbComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            vbProj.VBComponents.Remove vbComp
        End If
    Next vbComp
End Function

''''''''''''''''''''''''''''''''''''''''''
'Ron De Bruin Export - Import Modules END'
''''''''''''''''''''''''''''''''''''''''''





