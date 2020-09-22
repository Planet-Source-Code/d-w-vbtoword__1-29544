VERSION 5.00
Begin VB.Form Page1 
   Caption         =   "Vb to Word Sample"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9480
   Icon            =   "Page1.frx":0000
   LinkTopic       =   "Page1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   13
      Left            =   3855
      TabIndex        =   30
      Text            =   "Text(13)"
      Top             =   5430
      Width           =   3135
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8775
      Picture         =   "Page1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3030
      Width           =   675
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   825
      Left            =   8775
      Picture         =   "Page1.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4950
      Width           =   675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8775
      Picture         =   "Page1.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2070
      Width           =   675
   End
   Begin VB.CommandButton cmdWord 
      Caption         =   "WORD"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8775
      Picture         =   "Page1.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3990
      Width           =   675
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "OPEN"
      Height          =   825
      Left            =   8775
      Picture         =   "Page1.frx":1502
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   150
      Width           =   675
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "NEW"
      Height          =   825
      Left            =   8775
      Picture         =   "Page1.frx":1944
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1110
      Width           =   675
   End
   Begin VB.Timer DialogDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   180
      Top             =   2055
   End
   Begin VB.Timer NewDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   165
      Top             =   1500
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   12
      Left            =   3855
      TabIndex        =   11
      Text            =   "Text(12)"
      Top             =   5085
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   11
      Left            =   3855
      TabIndex        =   10
      Text            =   "Text(11)"
      Top             =   4735
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   10
      Left            =   3855
      TabIndex        =   9
      Text            =   "Text(10)"
      Top             =   4392
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   9
      Left            =   3855
      TabIndex        =   8
      Text            =   "Text(9)"
      Top             =   4049
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   8
      Left            =   3855
      TabIndex        =   7
      Text            =   "Text(9)"
      Top             =   3706
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   7
      Left            =   3855
      TabIndex        =   6
      Text            =   "Text(7)"
      Top             =   3363
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   6
      Left            =   3855
      TabIndex        =   5
      Text            =   "Text(6)"
      Top             =   3020
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   5
      Left            =   3855
      TabIndex        =   4
      Text            =   "Text(5)"
      Top             =   2677
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   4
      Left            =   3855
      TabIndex        =   3
      Text            =   "Text(4)"
      Top             =   2334
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   3855
      TabIndex        =   2
      Text            =   "Text(3)"
      Top             =   1991
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   3855
      TabIndex        =   1
      Text            =   "Text(2)"
      Top             =   1648
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   3855
      TabIndex        =   0
      Text            =   "Text(1)"
      Top             =   1305
      Width           =   3135
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1320
      TabIndex        =   32
      Top             =   285
      Width           =   6120
   End
   Begin VB.Label Label 
      Caption         =   "Label(13)"
      Height          =   270
      Index           =   13
      Left            =   1080
      TabIndex        =   31
      Top             =   5430
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(12)"
      Height          =   270
      Index           =   12
      Left            =   1080
      TabIndex        =   23
      Top             =   5085
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(11)"
      Height          =   270
      Index           =   11
      Left            =   1080
      TabIndex        =   22
      Top             =   4735
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(10)"
      Height          =   270
      Index           =   10
      Left            =   1080
      TabIndex        =   21
      Top             =   4392
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(9)"
      Height          =   270
      Index           =   9
      Left            =   1080
      TabIndex        =   20
      Top             =   4049
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(8)"
      Height          =   270
      Index           =   8
      Left            =   1080
      TabIndex        =   19
      Top             =   3706
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(7)"
      Height          =   270
      Index           =   7
      Left            =   1080
      TabIndex        =   18
      Top             =   3363
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(6)"
      Height          =   270
      Index           =   6
      Left            =   1080
      TabIndex        =   17
      Top             =   3020
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(5)"
      Height          =   270
      Index           =   5
      Left            =   1080
      TabIndex        =   16
      Top             =   2677
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(4)"
      Height          =   270
      Index           =   4
      Left            =   1080
      TabIndex        =   15
      Top             =   2334
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(3)"
      Height          =   270
      Index           =   3
      Left            =   1080
      TabIndex        =   14
      Top             =   1991
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(2)"
      Height          =   270
      Index           =   2
      Left            =   1080
      TabIndex        =   13
      Top             =   1648
      Width           =   2640
   End
   Begin VB.Label Label 
      Caption         =   "Label(1)"
      Height          =   270
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Top             =   1305
      Width           =   2640
   End
End
Attribute VB_Name = "Page1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



























Private Sub WriteLabels()
Label1 = "Internet Addiction Survey"
Label(1) = "Favorite browser:"
Label(2) = "Favorite website:"
Label(3) = "Type of internet connection:"
Label(4) = "Spouse or friend's first name:"
Label(5) = "Your first name:"
Label(6) = "Your email server:"
Label(7) = "Favorite news website:"
Label(8) = "Father's first name"
Label(9) = "Mother's first name"
Label(10) = "Your weight:"
Label(11) = "Your favorite graphics format:"
Label(12) = "Spouse or friend's hair color:"
Label(13) = "Favorite college:"
End Sub

Private Sub cmdExit_Click()
Unload Page1
End Sub

Private Sub cmdNew_Click()
DisableButtons
NewForm
End Sub

Private Sub cmdOpen_Click()
DisableButtons
OpenForm
End Sub

Private Sub cmdPrint_Click()
If Watching Then
    If CheckWord = False Then Exit Sub
End If
cmdPrint.Enabled = False
SavePage1
Pages.Save
AppWD.PrintOut Range:=wdPrintAllPages
cmdPrint.Enabled = True
End Sub

Private Sub cmdSave_Click()

If Watching Then
    If CheckWord = False Then Exit Sub
End If

If FileCreated Then
SavePage1
Pages.Save
End If

End Sub

Private Sub cmdWord_Click()
On Error Resume Next
Watching = True
If FileCreated Then
WordApp.ZOrder
WordApp.Caption = "Microsoft Word - " & TheDoc
WordApp.Visible = True
AppWD.Visible = True
End If

End Sub

Private Sub DialogDelay_Timer()
On Error GoTo Err1:
DialogDelay.Enabled = False
FileLoaded = False
If AppWD.Documents.Count > 0 Then
Pages.Close
End If
OpenDocument Filebox
OpenDoc = Filebox
Screen.MousePointer = vbHourglass
HideForms
ClearAllForms
Dialog.Progress.Refresh
Dialog.Progress.Value = 10
Dialog.Show
SetRanges1
ReadPage1
Me.Show
EnableButtons
EnableControls
FileLoaded = True
FormChanged = False
Dialog.Hide
Dialog.Progress.Value = 0
Screen.MousePointer = vbDefault
Exit Sub
Err1:
FileLoaded = True
End Sub

Private Sub Form_Load()
Dim Hwnd As Long
Dim Ret As Long
If App.PrevInstance Then
Hwnd = FindWindow(vbNullString, Caption)
Ret = ShowWindow(Hwnd, 3)
End
End If
Adjust_For_Resolution Me
WriteLabels
Load WordApp
DisableControls
ClearText
FormChanged = False
OpenWord
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If Watching Then
If CheckWord(True) = False Then GoTo TheExit:
End If
DisableControls
DisableButtons
Dim Rtn As VbMsgBoxResult
If FormChanged Then
Rtn = MsgBox("Some text fields have changed. Do you wish to save these changes?", vbYesNoCancel + vbCritical, "Save Changes?")
If Rtn = vbYes Then
If FileCreated Then
SavePage1
Else
Cancel = True
EnableControls
EnableButtons
Exit Sub
End If
ElseIf Rtn = vbCancel Then
Cancel = True
EnableControls
EnableButtons
Exit Sub
End If
End If
TheExit:
Checked = True
If Not QueryUnloadWord(WordApp) Then
Cancel = True
EnableControls
EnableButtons
Exit Sub
End If

ExitFlag = True
Me.Hide
UnloadForms
Set WdFrm = Nothing
End Sub

Private Sub NewDelay_Timer()
Dim i As Integer
On Error GoTo Err1:
NewDelay.Enabled = False
FileLoaded = False
If AppWD.Documents.Count > 0 Then
Pages.Close
End If
ClearAllForms
OpenTemplate
Dialog.Hide
Dialog.Progress.Value = 0
Me.Show
EnableButtons
FileLoaded = True
FormChanged = False
Screen.MousePointer = vbDefault
Exit Sub
Err1:
End Sub

Private Sub Text_Change(Index As Integer)
FormChanged = True
End Sub


