Attribute VB_Name = "WordAuto"
Option Explicit
Private Declare Function BringWindowToTop Lib "user32" _
    (ByVal Hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" _
    (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" _
    (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal Hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" _
    (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type ServerStatus
    ServerNotCreated As Integer
    ServerIsBusy As Integer
    ServerIsReady As Integer
End Type
Private Type MessageItem
   WordIsBusy As Integer
End Type

Public Const ServerNotCreated = 1
Public Const ServerIsBusy = 2
Public Const ServerIsReady = 3
Public Const WordIsBusy = 2

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const APP_CAPTION = "Microsoft Word"
Private Const APP_EXIT_DELAY = 3000
Private Const SERVER_BUSY_TIMEOUT = 500
Private Const ERR_SERVER_BUSY = -2147418111
Private Const SM_CXForm = 32
Private Const SM_CYForm = 33
Private Const SM_CYCAPTION = 4
Private Const TYPENAME_APPLICATION = "Application"
Private Const TYPENAME_OBJECT = "Object"
Private Const TYPENAME_DOCUMENT = "Document"


Public AppWD As Word.Application
Public Pages As Word.Document
Public wdHwnd As Long
Private wdOriginalRect As RECT
Private wdOriginalStatusBar As Boolean
Public WdFrm As Form


Public Function QueryUnloadWord(WdFrm As Form) As Boolean
If Not QuitWord Then QueryUnloadWord = False: Exit Function
WdFrm.Visible = False
DoEvents
Sleep APP_EXIT_DELAY
Unload WordApp
QueryUnloadWord = True
End Function





Private Function CloseWordDocument(Pages As Word.Document) As Boolean
Dim intAnswer As Integer
On Error GoTo ErrorHandler
If Pages.Saved Then
Pages.Close
ElseIf OpenDoc = App.Path & "\Template" Then
Pages.Saved = True
Pages.Close
Else
intAnswer = MsgBox("Do you want to save the changes you made to " _
    & Pages.Name & "?", vbYesNoCancel + vbExclamation, _
    "Microsoft Word")
    If intAnswer = vbYes Then
    Pages.Save
    Pages.Close
    ElseIf intAnswer = vbNo Then
    Pages.Saved = True
    Pages.Close
    Else
    Exit Function
    End If
End If
CloseWordDocument = True
ErrorHandler:
End Function
Private Sub DoMsgBox(msg As Integer)
Dim strMsg As String
If msg = 2 Then
strMsg = "Cannot automate Microsoft Word at this time. "
strMsg = strMsg & "Please make sure Microsoft Word is not busy "
strMsg = strMsg & "before you attempt this action."
End If
MsgBox strMsg, vbCritical, APP_CAPTION
End Sub
Public Function GetServerStatus(AppObj As Object) As Integer
Dim strTypeName As String
Dim strTest As String
strTypeName = TypeName(AppObj)
If strTypeName = TYPENAME_APPLICATION Then
GetServerStatus = ServerIsReady
ElseIf strTypeName = TYPENAME_OBJECT Then
On Error Resume Next
strTest = AppObj.Name
    If Err.Number = ERR_SERVER_BUSY Or Err.Number = 0 Then
    GetServerStatus = ServerIsBusy
    Else
    GetServerStatus = ServerNotCreated
    End If
Else
GetServerStatus = ServerNotCreated
End If
End Function
Public Sub OpenDocument(PathAndName As String)
Dim strPath As String
Dim lngRetVal As Long
On Error GoTo ErrorHandler

If GetServerStatus(AppWD) = ServerIsBusy Then
DoMsgBox WordIsBusy
Exit Sub
End If
lngRetVal = ShowWord
    If lngRetVal <> 0 Then Err.Raise lngRetVal

    If TypeName(Pages) = TYPENAME_DOCUMENT Then _
        If Not CloseWordDocument(Pages) Then Exit Sub

SetForegroundWindow wdHwnd
Set Pages = AppWD.Documents.Open(PathAndName)

Exit Sub
ErrorHandler:
MsgBox "Error " & Err.Number & ":" & vbCrLf & _
Err.Description, vbExclamation, APP_CAPTION
End Sub
Public Function OpenWord() As Boolean
App.OleServerBusyTimeout = SERVER_BUSY_TIMEOUT
App.OleServerBusyRaiseError = True
Dim lngRetVal As Long
lngRetVal = ShowWord

If lngRetVal = 0 Then
SetForegroundWindow wdHwnd
Else
MsgBox "Error " & lngRetVal & ":" & vbCrLf & _
Error(lngRetVal), vbExclamation, APP_CAPTION
OpenWord = False
Exit Function
End If

OpenWord = True
End Function
Private Function QuitWord() As Boolean

Dim WordStatus As Integer
Dim Pages As Word.Document
On Error GoTo ErrorHandler
WordStatus = GetServerStatus(AppWD)
    If WordStatus = ServerIsBusy Then
    OpenWord
    DoMsgBox WordIsBusy
    Exit Function
    End If
    
    If WordStatus = ServerIsReady Then
    OpenWord
    AppWD.ScreenUpdating = False
        For Each Pages In AppWD.Documents
        If Not CloseWordDocument(Pages) Then
        AppWD.ScreenUpdating = True
        Exit Function
        End If
        Next
    AppWD.ScreenUpdating = True
    End If

If wdHwnd <> 0 Then

    If WordStatus = ServerNotCreated Then
    Set AppWD = Nothing
    Set AppWD = CreateObject("Word.Application")

    AppWD.Caption = "besuretofindthisinstance"
    wdHwnd = FindWindow("OpusApp", AppWD.Caption)
    AppWD.Caption = "Microsoft Word"
    End If

    If AppWD.WindowState <> wdWindowStateNormal Then
    AppWD.Visible = True
    AppWD.WindowState = wdWindowStateNormal
    End If

AppWD.Visible = False

Set Pages = AppWD.Documents.Add
AppWD.DisplayStatusBar = wdOriginalStatusBar
Pages.Saved = True
Pages.Close

With AppWD.CommandBars("Menu Bar")
    .Controls("&File").Controls("&New...").Enabled = True
    .Controls("&File").Controls("&Close").Enabled = True
    .Controls("&File").Controls("E&xit").Enabled = True
    .Controls("&File").Controls("&Open...").Enabled = True
    .Controls("&File").Controls("&Save").Enabled = True
    .Controls("&File").Controls("Save &As...").Enabled = True
    .Controls("&File").Controls("Save as Web Page...").Enabled = True
End With

    If WordStatus = ServerIsReady Then _
        SetParent wdHwnd, 0

With wdOriginalRect
    MoveWindow wdHwnd, .Left, .Top, .Right - .Left, .Bottom - .Top, True
End With

AppWD.DisplayAlerts = wdAlertsNone
AppWD.Quit wdDoNotSaveChanges
End If
Set Pages = Nothing
Set AppWD = Nothing
QuitWord = True
ErrorHandler:
End Function
Public Sub SetAppSize(lngHwnd As Long, WdFrm As Form)
Dim lngX As Long
Dim lngY As Long
Dim lngW As Long
Dim lngH As Long
Dim AppRect As RECT
GetWindowRect WdFrm.Hwnd, AppRect
lngX = -GetSystemMetrics(SM_CXForm)
lngY = -GetSystemMetrics(SM_CYForm)
lngW = AppRect.Right - AppRect.Left - lngX * 2
lngH = AppRect.Bottom - AppRect.Top - lngY * 2
lngY = lngY - GetSystemMetrics(SM_CYCAPTION)
lngH = lngH + GetSystemMetrics(SM_CYCAPTION)
MoveWindow lngHwnd, lngX, lngY, lngW, lngH, True
End Sub

Private Function ShowWord() As Long
Static blnGetRect As Boolean
Dim wrdAppTemp As Word.Application
On Error GoTo ErrorHandler

If GetServerStatus(AppWD) = ServerNotCreated Then
Set AppWD = Nothing
Set wrdAppTemp = CreateObject("Word.Application")
Set AppWD = CreateObject("Word.Application")
wrdAppTemp.Quit
Set wrdAppTemp = Nothing
AppWD.Caption = "besuretofindthisinstance"
wdHwnd = FindWindow("OpusApp", AppWD.Caption)
AppWD.Caption = "Microsoft Word"
Set Pages = AppWD.Documents.Add
wdOriginalStatusBar = AppWD.DisplayStatusBar
AppWD.DisplayStatusBar = True
Pages.Saved = True
Pages.Close
Set Pages = Nothing
With AppWD.CommandBars("Menu Bar")
    .Controls("&File").Controls("&New...").Enabled = False
    .Controls("&File").Controls("&Close").Enabled = False
    .Controls("&File").Controls("E&xit").Enabled = False
    .Controls("&File").Controls("&Open...").Enabled = False
    .Controls("&File").Controls("&Save").Enabled = False
    .Controls("&File").Controls("Save &As...").Enabled = False
    .Controls("&File").Controls("Save as Web Page...").Enabled = False
End With
AppWD.NormalTemplate.Saved = True
AppWD.Visible = False
If AppWD.WindowState <> wdWindowStateNormal Then _
AppWD.WindowState = wdWindowStateNormal
    If Not blnGetRect Then
    GetWindowRect wdHwnd, wdOriginalRect
    blnGetRect = True
    End If
BringWindowToTop wdHwnd
SetParent wdHwnd, WdFrm.Hwnd
SetAppSize wdHwnd, WdFrm
End If

ShowWord = 0
Exit Function
ErrorHandler:
ShowWord = Err.Number
End Function






