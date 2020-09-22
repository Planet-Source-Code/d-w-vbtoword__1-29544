Attribute VB_Name = "OpenFile"
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public TheDoc As String
Public Filebox As String
Public Function SelectedFile() As String
Dim Ret As String
Dim Ofn As OPENFILENAME
Dim L As Long
With Ofn
   .lStructSize = Len(Ofn)
   .hwndOwner = Screen.ActiveForm.Hwnd
   .hInstance = App.hInstance
   .lpstrFilter = "Survey Forms(*.net)" + Chr(0) + "*.net" + Chr(0)
   .lpstrFile = Space(254)
   .nMaxFile = 255
   .lpstrFileTitle = Space(254)
   .nMaxFileTitle = 255
   .lpstrInitialDir = App.Path
   .lpstrTitle = "Open an existing set of forms."
   .flags = 0
End With
L = GetOpenFileName(Ofn)
    If L Then
    Ret = Trim(Ofn.lpstrFile)
    Else
    Ret = ""
    End If
TheDoc = Ofn.lpstrFileTitle
SelectedFile = Ret
End Function

