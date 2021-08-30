Attribute VB_Name = "Imports"
Sub ImportComponents(wb As Workbook)
    Dim varr                                     'As Variant
    Dim ELEMENT                                  'As String
    Dim proceed As Boolean, hasWorksheets As Boolean
    proceed = True
    Dim compName As String
    
    'file dialogue multiselect components to import
    varr = getFilePath(Array("bas", "frm", "cls"), True)
    If Not ArrayIsAllocated(varr) Then Exit Sub
    Dim vbProj As VBProject
    Set vbProj = wb.VBProject
    
    Dim coll As Collection
    Set coll = New Collection
    
    For Each ELEMENT In varr
        compName = getFilePartName(CStr(ELEMENT), False)
        Debug.Print compName
        If compName Like "DocClass*" Then
            compName = Right(compName, Len(compName) - 6)
            hasWorksheets = True
        End If
        If ModuleExists(compName, wb) = True Then
            proceed = False
            coll.Add compName
        End If
    Next ELEMENT
    If proceed = False Then GoTo Errorhandler
    
    Dim wasOpen As Boolean
    Dim wbSource As Workbook
    Dim wbSourceName As String
    Dim basePath As String
    basePath = getFilePartPath(varr(1), True)
    If hasWorksheets = True Then
        
        wbSourceName = Dir(basePath & "*.xl*")
        If wbSourceName <> "" Then
            wasOpen = WorkbookIsOpen(wbSourceName)
            If wasOpen = False Then
                Set wbSource = Workbooks.Open(basePath & wbSourceName)
            Else
                Set wbSource = Workbooks(wbSourceName)
            End If
        End If
    End If
    
    For Each ELEMENT In varr
        compName = getFilePartName(CStr(ELEMENT), False)
        If Not compName Like "DocClass*" Then
            vbProj.VBComponents.Import ELEMENT
        Else
            compName = Right(compName, Len(compName) - 9)
            If compName <> "ThisWorkbook" Then
                wbSource.Sheets(compName).Copy Before:=wb.Sheets(1)
            End If
        End If
    Next ELEMENT
    GoTo exitHandler
Errorhandler:
    Dim str As String
    str = "The following components already exist. All import canceled."
    For Each ELEMENT In coll
        str = str & vbNewLine & ELEMENT
    Next ELEMENT
    MsgBox str
    Exit Sub
exitHandler:
    If wasOpen = False And WorkbookIsOpen(wbSourceName) Then wbSource.Close False
    Set vbProj = Nothing
    Set coll = Nothing
    Set wbSource = Nothing
    MsgBox "Import successful"
    
End Sub

