VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Moves"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1470
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   1470
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   4005
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SWW_HPARENT = (-8)

Private Sub Form_Load()
SetWindowLong Me.hwnd, SWW_HPARENT, Form1.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong Me.hwnd, SWW_HPARENT, 0
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub
