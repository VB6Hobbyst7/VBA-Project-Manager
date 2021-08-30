Attribute VB_Name = "AddinMaker"

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : F_AddInInstall
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Modified   : 24-08-2021 by Anastasiou Alex to be standalone for any project
'*                         #vbaCodeArchive - https://github.com/alexofrhodes - anastasioualex@gmail.com
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public Sub AddinCreate()
    Dim AddFolder As String
    On Error GoTo InstallationAdd_Err
    AddFolder = Replace(Application.UserLibraryPath & "\", "\\", "\")
    If Dir(AddFolder, vbDirectory) = vbNullString Then
        Call MsgBox("Unfortunately, the program cannot install the add-in on this computer." _
                  & vbCrLf & "The settings directory is missing." & vbCrLf & _
                    "Contact the program developer.", vbCritical, _
                    "Add-in installation failed")
        Exit Sub
    End If
    Dim addinsPath As String
    addinsPath = Application.UserLibraryPath
    Dim partName As String
    partName = Right(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - InStrRev(ThisWorkbook.FullName, "\"))
    partName = Left(partName, InStr(1, partName, ".") - 1)
      
    If Dir(addinsPath & partName & ".xlam") <> "" Then AddIns(partName).Installed = False
    If WorkbookIsOpen(partName & ".xlam") Then
        Call MsgBox("The file with the add-in is already open." & vbCrLf & _
                    "It may have already been installed earlier.", vbCritical, _
                    "Program installation failed")
        Exit Sub
    End If
    
    Application.EnableEvents = 0
    Application.DisplayAlerts = False
    If Workbooks.Count = 0 Then Workbooks.Add
    ThisWorkbook.SaveAs addinsPath & partName & ".xlam", FileFormat:=xlOpenXMLAddIn
    AddIns.Add fileName:=addinsPath & partName & ".xlam"
    AddIns(partName).Installed = True
    Application.EnableEvents = 1
    Application.DisplayAlerts = True
    Call MsgBox("The program is installed successfully!" & vbCrLf & _
                "Just open or create a new document.", vbInformation, _
                "Installing the add-in:" & partName)
    ThisWorkbook.Close False
    Exit Sub
InstallationAdd_Err:
    If Err.Number = 1004 Then
        MsgBox "To install the add-in, please close this file and run it again.", _
               64, "Installation"
    Else
        MsgBox Err.Description & vbCrLf & " Addin installation failed "
    End If
End Sub

