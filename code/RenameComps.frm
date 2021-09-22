VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameComps 
   Caption         =   "COMPONENT RENAMER"
   ClientHeight    =   7080
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5700
   OleObjectBlob   =   "RenameComps.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'NEW
Private Sub RenameComponents_Click()

        Dim NewNames As Variant
        Dim i As Long
        NewNames = Split(RenameComps.textboxNewName, vbNewLine)
        For i = 0 To UBound(NewNames)
            If NewNames(i) = vbNullString Then
                NewNames(i) = LRenameListbox.List(i)
            End If
        Next i
        
        For i = 0 To UBound(NewNames) ' RenameComps.LRenameListbox.ListCount - 1
continue:
            On Error GoTo EH
            Select Case RenameComps.LRenameListbox.List(i, 0)
            'rename component
            Case Is = "Module", "Class", "UserForm"
                If RenameComps.LRenameListbox.List(i, 1) <> NewNames(i) Then
                    wb.VBProject.VBComponents(RenameComps.LRenameListbox.List(i, 1)).name = NewNames(i)
                End If
            'rename document (new worksheet name)
            Case Is = "Document"
                'can't rename ThisWorkbook
                If RenameComps.LRenameListbox.List(i, 1) <> "ThisWorkbook" And NewNames(i) <> "ThisWorkbook" Then
                    If RenameComps.LRenameListbox.List(i, 1) <> NewNames(i) Then
                        wb.Sheets(LDocument.List(RenameComps.LRenameListbox.List(i, 1))).name = NewNames(i)
                    End If
                Else
                'nothing
                End If
            Case Else
                'nothing
            End Select
        Next
        'replace old names with new names in listbox
        For i = 0 To RenameComps.LRenameListbox.ListCount - 1
            RenameComps.LRenameListbox.List(i, 1) = NewNames(i)
        Next i
        'update user's new names in textbox with actual new names
        RenameComps.textboxNewName.Text = vbNullString
        Dim str As String
        str = Join(NewNames, vbNewLine)
        RenameComps.textboxNewName.Text = str
       
        MsgBox "Components renamed"
        Exit Sub
EH:
        'hanlde user giving duplicate name by incrementing
        NewNames(i) = NewNames(i) & i + 1
        Resume continue
End Sub


Private Sub UserForm_Initialize()
    Dim vbcomp As VBComponent
    For Each vbcomp In ActiveWorkbook.VBProject.VBComponents
        RenameComps.LRenameListbox.AddItem
        RenameComps.LRenameListbox.List(RenameComps.LRenameListbox.ListCount - 1, 0) = ComponentTypeToString(vbcomp.Type)
        If vbcomp.Type <> vbext_ct_Document Then
            RenameComps.LRenameListbox.List(RenameComps.LRenameListbox.ListCount - 1, 1) = vbcomp.name
        Else
            If vbcomp.name = "ThisWorkbook" Then
                RenameComps.LRenameListbox.List(RenameComps.LRenameListbox.ListCount - 1, 1) = vbcomp.name
            Else
                RenameComps.LRenameListbox.List(RenameComps.LRenameListbox.ListCount - 1, 1) = GetSheetByCodeName(wb, vbcomp.name).name
            End If
        End If
    Next
    SortListboxOnColumn LRenameListbox, 0
    Dim str As String
    str = LRenameListbox.List(0, 1)
    For i = 1 To LRenameListbox.ListCount - 1
    str = str & vbNewLine & LRenameListbox.List(i, 1)
    Next
    textboxNewName.Text = str
End Sub
Public Function SetSheetByCodeName(wb As Workbook, CodeName As String) As Worksheet

    Dim sh As Worksheet
    For Each sh In wb.Worksheets                 'Run loop.
        If UCase(sh.CodeName) = UCase(CodeName) Then Set SetSheetByCodeName = sh: Exit For 'Check if it's that sheet or not and set if true
    Next sh
End Function


