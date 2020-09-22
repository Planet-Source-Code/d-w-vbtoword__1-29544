VERSION 5.00
Begin VB.Form WordApp 
   Caption         =   "Microsoft Word"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9480
   Icon            =   "WordApp.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6780
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "WordApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Set WdFrm = WordApp
End Sub









Private Sub Form_Paint()
Caption = "Microsoft Word - " & TheDoc
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Frm As Form
If Not ExitFlag Then
Cancel = True

Me.Hide

For Each Frm In Forms
If Frm.Visible Then
Frm.Show
Frm.ZOrder
End If
Next
End If
End Sub


Private Sub Form_Resize()
SetAppSize wdHwnd, WdFrm
End Sub


