VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Okay"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHESS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "http://vbaccelerator.com/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Thanks to"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Created by"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Gary Eyles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Created with"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3060
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   3060
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Label4.Caption <> "Gary Eyles" Then
Label4.Caption = "Gary Eyles"
End If
End Sub

