VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get picture from...."
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Get internal picture"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load a file"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Tag = "1"
Visible = False
End Sub

Private Sub Command2_Click()
Tag = "0"
Visible = False
End Sub
