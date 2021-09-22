Attribute VB_Name = "UF"
Public wb As Workbook

Sub ShowMe()
    If Not IsLoaded("ProjectManager") Then
        ProjectManager.Show
    End If
End Sub

Sub addCompsList()
    RemoveComps.LComponents.Clear
    Dim vbcomp As VBComponent
    For Each vbcomp In wb.VBProject.VBComponents
        RemoveComps.LComponents.AddItem
        RemoveComps.LComponents.List(RemoveComps.LComponents.ListCount - 1, 0) = ComponentTypeToString(vbcomp.Type)
        RemoveComps.LComponents.List(RemoveComps.LComponents.ListCount - 1, 1) = vbcomp.name
    Next
    SortListboxOnColumn RemoveComps.LComponents, 0
    RemoveComps.Repaint
End Sub

Function IsLoaded(formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function

Sub AddCommandbar()
    'Delete any existing menu item that may have been left.
    Call DeleteCommandBar
    '    Dim cControl As Office.CommandBarButton
    'Add the new menu item and set a CommandBarButton variable to it
    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add(Temporary:=True)
    With cControl
        '        .Picture = LoadPicture(picturePath) 'need to set cControl as office.commandbarbutton
        .Caption = "ProjectManager"
        .Style = msoButtonIconAndCaption
        .FaceId = 4181
        .OnAction = "ShowMe"                     'Macro stored in a Standard Module
    End With
    On Error GoTo 0
End Sub

Sub DeleteCommandBar()
    On Error Resume Next
    Dim bar As CommandBarControl
    For Each bar In Application.CommandBars("Worksheet Menu Bar").Controls
        If bar.Caption = "ProjectManager" Then bar.Delete
        'Debug.Print bar.Caption
    Next
End Sub

Function PickExcelFile() As String
    Dim strFile As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xl*", 1
        .Title = "Choose an Excel file"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERprofile") & "\Desktop\"
        If .Show = True Then
            strFile = .SelectedItems(1)
            PickExcelFile = strFile
        End If
    End With
End Function

Sub SortListboxOnColumn(lBox As MSForms.ListBox, OnColumn As Long)
    Dim vntData As Variant
    Dim vntTempItem As Variant
    Dim lngOuterIndex As Long
    Dim lngInnerIndex As Long
    Dim lngSubItemIndex As Long
    'Store the list in an array for sorting
    vntData = lBox.List
    'Bubble sort the array on the first value
    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
            If vntData(lngOuterIndex, OnColumn) > vntData(lngInnerIndex, OnColumn) Then
                'Swap values
                For lngSubItemIndex = 0 To lBox.ColumnCount - 1
                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
                Next
            End If
        Next lngInnerIndex
    Next lngOuterIndex
    'Remove the contents of the listbox
    lBox.Clear
    'Repopulate with the sorted list
    lBox.List = vntData
End Sub

