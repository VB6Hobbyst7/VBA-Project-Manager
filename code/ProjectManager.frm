VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectManager 
   Caption         =   "INPUT OUTPUT"
   ClientHeight    =   4896
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   1992
   OleObjectBlob   =   "ProjectManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    chExportSheets.Value = ThisWorkbook.Sheets("SETTINGS").Range("ExportSheets").Value
    chExportForms.Value = ThisWorkbook.Sheets("SETTINGS").Range("ExportForms").Value
End Sub

Private Sub chExportSheets_Click()
    ThisWorkbook.Sheets("SETTINGS").Range("ExportSheets").Value = chExportSheets.Value
End Sub

Private Sub chExportForms_Click()
    ThisWorkbook.Sheets("SETTINGS").Range("ExportForms").Value = chExportForms.Value
End Sub

Private Sub ActiveFile_Click()
    ActiveFile.SpecialEffect = fmSpecialEffectSunken
    ActiveFile.Width = 90
    DoEvents
    Sleep 50
    ActiveFile.SpecialEffect = fmSpecialEffectFlat
    ActiveFile.BorderStyle = fmBorderStyleSingle
    ActiveFile.Width = 90

    Set wb = ActiveWorkbook
    SelectAction
End Sub

Private Sub ActiveFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ActiveFile.BackColor = RGB(255, 187, 120)    ''' sample colour
    MakeAllElementsWhite ActiveFile.name
End Sub

Sub SelectAction()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SETTINGS")
    Select Case True
    Case oExport.Value = True
        Me.Hide
        ExportProject wb, ws.Range("ExportSheets"), ws.Range("ExportForms")
        Me.Show
    Case oImport.Value = True
        ImportComponents wb
    Case oRefresh.Value = True
        RefreshComponents wb
    Case chDelete.Value = True
        Comps.Show
    Case Else
        '
    End Select
End Sub

Private Sub cInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExportSettings.Show
End Sub

Private Sub SelectFile_Click()
    SelectFile.SpecialEffect = fmSpecialEffectSunken
    SelectFile.Width = 90
    DoEvents
    Sleep 50                                     ' in module pu Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    SelectFile.SpecialEffect = fmSpecialEffectFlat
    SelectFile.BorderStyle = fmBorderStyleSingle
    SelectFile.Width = 90

    Dim fPath As String
    fPath = PickExcelFile
    If fPath = "" Then Exit Sub
    Set wb = Workbooks.Open(fileName:=fPath, UpdateLinks:=0, ReadOnly:=False)
    SelectAction
    Set wb = Nothing
End Sub

Private Sub SelectFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SelectFile.BackColor = RGB(255, 187, 120)    ''' sample colour
    MakeAllElementsWhite SelectFile.name
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single) ''' example of the code to make all elements of the user form white
    MakeAllElementsWhite
End Sub

Private Sub MakeAllElementsWhite(Optional NameOfEleToExcludeAsString As String)
    Dim ele As Variant
    On Error Resume Next
    For Each ele In Me.Controls                  ''' me is the userform
        If ele.name <> NameOfEleToExcludeAsString Then ele.BackColor = vbWhite
    Next ele
End Sub

