Attribute VB_Name = "modAPI"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit
Option Compare Text

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_SHOWNORMAL = 1
Private Const mcModuleName = "modAPI"

Public Sub ShellOpen(ByVal strCommand As String, Optional ByVal strFolder As String = "")
    On Error GoTo ErrorTag
    
    If strFolder = "" Then strFolder = App.Path
    
    ShellExecute vbNull, vbNullString, strCommand, vbNullString, App.Path, SW_SHOWNORMAL
    Exit Sub

ErrorTag:
    Call LogError(Err.Number, Err.Description, Err.Source, _
                  mcModuleName, "ShellOpen")
End Sub

