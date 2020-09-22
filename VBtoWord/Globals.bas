Attribute VB_Name = "Globals"
Option Explicit
Option Compare Text
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal Hwnd As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_RESTORE = 9
 
Public FldRng(1 To 13) As Range

Public OpenDoc As String
Public Ctrl As Control
Public FormChanged As Boolean
Public Checked As Boolean
Public ExitFlag As Boolean
Public FileLoaded As Boolean




Public FileSaved As Boolean
Public Watching As Boolean
Public Function CheckWord(Optional Quit As Boolean = False) As Boolean
If AppWD.Documents.Count = 0 Then
DisableButtons
    If Not Quit Then
    OpenForm
    End If
CheckWord = False
Else
Watching = False
CheckWord = True
End If
End Function



Public Sub ClearAllForms()
Dim i As Integer
Dim frm As Form
Dim Ctrl As Control
For Each frm In Forms
For Each Ctrl In frm
If TypeOf Ctrl Is TextBox Then
Ctrl.Text = ""
End If
Next
Next
End Sub


Public Function Random_X(Digits As Integer) As Long
Dim UpperBound As Long
Dim LowerBound As Long
Randomize Timer
    If Digits > 8 Then Digits = 8
UpperBound = (10 ^ Digits) - 1
LowerBound = (10 ^ (Digits - 1))
Random_X = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function
Public Sub Adjust_For_Resolution(frm As Form)
For Each Ctrl In frm.Controls
If TypeOf Ctrl Is Menu Or TypeOf Ctrl Is Form Or _
 TypeOf Ctrl Is Timer Or TypeOf Ctrl Is VScrollBar Or _
 TypeOf Ctrl Is Line Then GoTo Skip:
With Ctrl
 .Left = Ctrl.Left * Resize
 .Top = Ctrl.Top * Resize
 .Height = Ctrl.Height * Resize
 .Width = Ctrl.Width * Resize
 .FontSize = Ctrl.FontSize * Resize
End With
Skip:
Next
End Sub
Public Sub EnableControls()
Dim frm As Form
For Each frm In Forms
    For Each Ctrl In frm.Controls
        If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is OptionButton _
            Or TypeOf Ctrl Is Label Then
        Ctrl.Enabled = True
        End If
    Next
Next
End Sub

Public Sub DisableControls()
Dim frm As Form
For Each frm In Forms
    For Each Ctrl In frm.Controls
        If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is OptionButton _
            Or TypeOf Ctrl Is Label Then
        Ctrl.Enabled = False
        End If
    Next
Next
End Sub
Public Sub DisableButtons()
Dim frm As Form
Dim Ctrl As Control
For Each frm In Forms
    For Each Ctrl In frm
        If TypeOf Ctrl Is CommandButton Then
        Ctrl.Enabled = False
        End If
    Next
Next
End Sub
Public Sub EnableButtons()
Dim frm As Form
Dim Ctrl As Control
For Each frm In Forms
    For Each Ctrl In frm
        If TypeOf Ctrl Is CommandButton Then
        Ctrl.Enabled = True
        End If
    Next
Next
End Sub

Public Sub BringToFront()
Dim Hwnd As Long
Hwnd = FindWindow("OpusApp", vbNullString)
BringWindowToTop Hwnd
End Sub

Public Sub HideForms()
Dim frm As Form
For Each frm In Forms
frm.Visible = False
Next
End Sub

Public Sub NewForm()
Dim Rtn As VbMsgBoxResult
If FormChanged Then
Rtn = MsgBox("Some text fields have changed. Do you wish to save these changes?", vbYesNoCancel + vbCritical, "Save Changes?")
    If Rtn = vbYes Then
        If FileCreated Then
        SavePage1
        Else
        Exit Sub
        End If
    ElseIf Rtn = vbCancel Then
    EnableButtons
    Exit Sub
    End If
End If
Screen.MousePointer = vbHourglass
Dialog.Message = "Loading Please Wait"
Dialog.Show
Page1.NewDelay.Enabled = True
End Sub
Public Sub OpenForm()
Dim Rtn As VbMsgBoxResult
If FormChanged And Not Watching Then
Rtn = MsgBox("Some text fields have changed. Do you wish to save these changes?", vbYesNoCancel + vbCritical, "Save Changes?")
    If Rtn = vbYes Then
        If FileCreated Then
        Else
        Exit Sub
        End If
    ElseIf Rtn = vbCancel Then
    EnableButtons
    Exit Sub
    End If
End If

Filebox = SelectedFile
If Filebox <> "" Then
    If Dir(Filebox) <> "" Then
    Watching = False
    Dialog.Message = "Loading Please Wait"
    Dialog.Progress.Value = 0
    Dialog.Show
    FileSaved = True
    Page1.DialogDelay.Enabled = True
    Else
    MsgBox "No valid file selected."
    End If
Else
    If FileLoaded Then
    EnableButtons
    Else
    EnablePartial
    End If
End If
Exit Sub

Err1:
MsgBox Err.Number & " " & Err.Description & "  Error in opening Word document.", vbCritical
End Sub

Private Sub EnablePartial()
Page1.cmdOpen.Enabled = True
Page1.cmdNew.Enabled = True
Page1.cmdExit.Enabled = True
End Sub





Public Sub UnloadForms()
Dim frm As Form
For Each frm In Forms
    If Not frm Is Page1 Then
    Unload frm
    Set frm = Nothing
    End If
Next
End Sub
Public Function FileCreated() As Boolean
On Error GoTo Err1:
Dim Name As String
Dim File_Name As String
If OpenDoc = App.Path & "\Template" Then
Name = Trim(Page1.Text(5))
    If Name = "" Then
    MsgBox "Please enter first name for a filename.", vbCritical
    FileCreated = False
    Exit Function
    End If
    
    If Len(Name) > 8 Then
    Name = Left(Name, 8)
    End If

File_Name = Name & Random_X(12 - Len(Name))
TheDoc = File_Name
File_Name = App.Path & "\" & File_Name
Pages.SaveAs FileName:=File_Name & ".net"
OpenDoc = File_Name & ".net"
Else
FileSaved = True
End If

FileCreated = True
Exit Function
Err1:
FileCreated = False
MsgBox "Error creating file."
End Function









Public Sub OpenTemplate()
Dim i As Integer
On Error Resume Next
    If AppWD.Documents.Count <> 0 Then
    Pages.Saved = True
    Pages.Close
    End If
On Error GoTo 0
On Error GoTo Err1:
    If OpenWord Then
    FileSaved = False
    OpenDocument App.Path & "\Template"
    OpenDoc = App.Path & "\Template"
    AppWD.Visible = False
    Dialog.Progress.Value = 10
    SetRanges1
    EnableControls
    End If
Exit Sub
Err1:
MsgBox Err.Number & " " & Err.Description & ". Error in opening template.", vbCritical
End Sub
Public Sub ClearText()
Dim frm As Form
Dim Ctrl As Control
For Each frm In Forms
    For Each Ctrl In frm
        If TypeOf Ctrl Is TextBox Then
        Ctrl.Text = ""
        End If
    Next
Next
End Sub
Public Function Resize() As Single
On Error GoTo Err1:

Select Case Screen.Width
Case 9600
Resize = 1
Case 12000
Resize = 1.25
Case 15360
Resize = 1.6
Case 19200
Resize = 2
Case Else
Resize = 1
End Select

Exit Function
Err1:
MsgBox Err.Number & " " & Err.Description & ". Error in resizing for resolution.", vbCritical
End Function

Public Sub ReadPage1()
Dim i As Integer
    For i = 1 To 13
    Page1.Text(i) = FldRng(i)
    Next
FormChanged = False
End Sub
Public Sub SavePage1()
Dim i As Integer
    For i = 1 To 13
    FldRng(i) = Page1.Text(i)
    Next
Pages.Save
FormChanged = False
End Sub

Public Sub SetRanges1()
Dim i As Integer
For i = 1 To 13
Set FldRng(i) = Pages.Range(Start:=Pages.Bookmarks(i * 2 - 1).Range.End, _
      End:=Pages.Bookmarks(i * 2).Range.Start - 1)
Dialog.Progress.Value = i / 13 * 100
Next
End Sub

