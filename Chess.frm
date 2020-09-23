VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Chess"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6000
   Icon            =   "Chess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin Chess.vbalImageList brdPeices 
      Left            =   1440
      Top             =   5280
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   16
      Size            =   41808
      Images          =   "Chess.frx":068A
      KeyCount        =   12
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5280
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList MnuIcons 
      Left            =   5280
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chess.frx":A9FA
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chess.frx":AD16
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chess.frx":B032
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chess.frx":B34E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chess.frx":B66A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox wCheck 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5280
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox bCheck 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5280
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5400
      Top             =   3120
   End
   Begin VB.PictureBox TempDrag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5280
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Enabled         =   0   'False
      Height          =   4815
      Left            =   3840
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   15
         Left            =   600
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   14
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   13
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   12
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   11
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   10
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   9
         Left            =   600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   8
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   7
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   6
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   5
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   4
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   3
         Left            =   600
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   1
         Left            =   600
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Wpeices 
         Height          =   615
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4815
      Left            =   2280
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   15
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   14
         Left            =   600
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   13
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   12
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   11
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   10
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   9
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   8
         Left            =   600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   7
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   6
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   5
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   4
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   2
         Left            =   600
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Bpeices 
         Height          =   615
         Index           =   0
         Left            =   600
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox temp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5280
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PictTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   5280
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   630
   End
   Begin MSComDlg.CommonDialog LoadSave 
      Left            =   5280
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Chess"
      Filter          =   "Chess Files (*.csh) | *.csh"
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   20
      Left            =   2520
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   55
      Left            =   4320
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   840
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Info 
      BackStyle       =   0  'Transparent
      Caption         =   "name peice"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Black"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current player :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   19
      Left            =   1920
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   54
      Left            =   3720
      Top             =   3720
      Width           =   615
   End
   Begin VB.Shape Shape3 
      Height          =   5055
      Left            =   120
      Top             =   120
      Width           =   5055
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   63
      Left            =   4320
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   62
      Left            =   3720
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   61
      Left            =   3120
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   60
      Left            =   2520
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   59
      Left            =   1920
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   58
      Left            =   1320
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   57
      Left            =   720
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   56
      Left            =   120
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   53
      Left            =   3120
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   52
      Left            =   2520
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   51
      Left            =   1920
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   50
      Left            =   1320
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   49
      Left            =   720
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   48
      Left            =   120
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   47
      Left            =   4320
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   46
      Left            =   3720
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   45
      Left            =   3120
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   44
      Left            =   2520
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   43
      Left            =   1920
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   42
      Left            =   1320
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   41
      Left            =   720
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   40
      Left            =   120
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   39
      Left            =   4320
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   38
      Left            =   3720
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   37
      Left            =   3120
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   36
      Left            =   2520
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   35
      Left            =   1920
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   34
      Left            =   1320
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   33
      Left            =   720
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   32
      Left            =   120
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   31
      Left            =   4320
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   30
      Left            =   3720
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   29
      Left            =   3120
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   28
      Left            =   2520
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   27
      Left            =   1920
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   26
      Left            =   1320
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   25
      Left            =   720
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   24
      Left            =   120
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   23
      Left            =   4320
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   22
      Left            =   3720
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   21
      Left            =   3120
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   18
      Left            =   1320
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   17
      Left            =   720
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   16
      Left            =   120
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   15
      Left            =   4320
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   14
      Left            =   3720
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   13
      Left            =   3120
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   12
      Left            =   2520
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   11
      Left            =   1920
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   10
      Left            =   1320
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   9
      Left            =   720
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   8
      Left            =   120
      Top             =   720
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   7
      Left            =   4320
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   6
      Left            =   3720
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   5
      Left            =   3120
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   4
      Left            =   2520
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   3
      Left            =   1920
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   2
      Left            =   1320
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   1
      Left            =   720
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Board 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   4920
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   240
      Top             =   4920
      Width           =   4815
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu newgame 
         Caption         =   "New"
      End
      Begin VB.Menu none2 
         Caption         =   "-"
      End
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu none1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu viewpeices 
         Caption         =   "View taken peices"
      End
      Begin VB.Menu viewmoves 
         Caption         =   "View moves"
      End
      Begin VB.Menu viewshadow 
         Caption         =   "View Shadow"
         Checked         =   -1  'True
      End
      Begin VB.Menu viewinfo 
         Caption         =   "View peice info"
         Checked         =   -1  'True
      End
      Begin VB.Menu none3 
         Caption         =   "-"
      End
      Begin VB.Menu bcolour 
         Caption         =   "Choose black check colour"
         Checked         =   -1  'True
      End
      Begin VB.Menu btexture 
         Caption         =   "Choose black check texture"
      End
      Begin VB.Menu wcolour 
         Caption         =   "Choose white check colour"
         Checked         =   -1  'True
      End
      Begin VB.Menu wtexture 
         Caption         =   "Choose white check texture"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu PawntoQueen 
         Caption         =   "Pawn to Queen"
         Checked         =   -1  'True
      End
      Begin VB.Menu none4 
         Caption         =   "-"
      End
      Begin VB.Menu saveonexit 
         Caption         =   "Save settings on exit"
      End
   End
   Begin VB.Menu about 
      Caption         =   "&About"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare all peices
Const bPawn = 1
Const bRook = 2
Const bKnight = 3
Const bBishop = 4
Const bQueen = 5
Const bKing = 6
Const wPawn = 7
Const wRook = 8
Const wKnight = 9
Const wBishop = 10
Const wQueen = 11
Const wKing = 12
Const pBlack = 13
Const pWhite = 14

Dim wDc As Long
Dim Draging As Boolean
Dim NeedRedraw As Boolean
Dim Timer2Counter As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const HTBORDER = 18
Private Const HTCAPTION = 2
Private Const HTMENU = 5

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Const ILD_TRANSPARENT = 1&
Private Const ILD_BLEND25 = 2&
Private Const CLR_NONE = -1

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" _
     (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" _
     (ByVal hMenu As Long) _
     As Long
Private Declare Function RemoveMenu Lib "user32" _
     (ByVal hMenu As Long, ByVal _
     nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

Dim xyCol(9, 9) As Boolean
Dim xyPcs(9, 9) As Integer
Dim ClickPeice As Integer
Dim HasCleared As Boolean
Dim WhitePeices, BlackPeices As Integer
Dim rWpeices(16) As Integer
Dim rBpeices(16) As Integer
Dim BlackColour As Long
Dim WhiteColour As Long
Dim GoingAbout As Boolean
Dim Winner As Boolean
Dim DoneSizing As Boolean

Function wColor(pc As Integer) As Integer
'Function tells you whether it is a
'Black or White peice
If pc < 7 And pc > 0 Then
    wColor = pBlack
ElseIf pc > 6 Then
    wColor = pWhite
End If
End Function

Function IsCheck(WorB As Integer) As Boolean
Dim xx, yy As Integer
Dim KingX, KingY As Integer
Dim Found As Boolean
Dim look As Boolean

For yy = 1 To 8
  For xx = 1 To 8
    If xyPcs(xx, yy) = bKing Or xyPcs(xx, yy) = wKing Then
      If wColor(xyPcs(xx, yy)) = WorB Then
        KingX = xx
        KingY = yy
        Found = True
        Exit For
      End If
    End If
  Next
  If Found Then Exit For
Next

For yy = 1 To 8
    For xx = 1 To 8
        If wColor(xyPcs(xx, yy)) <> WorB Then
            look = CheckIfValid(xx, yy, wColor(xyPcs(xx, yy)), KingX, KingY)
            If look = True Then
                IsCheck = True
                Exit Function
            End If
        End If
    Next
Next
End Function

Sub DrawCheck(xx, yy)
If xyCol(xx, yy) Then
    If bcolour.Checked Then
        PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), BlackColour, BF
    Else
        'PictTemp.PaintPicture bCheck, 0, 0, PictTemp.ScaleWidth, PictTemp.ScaleHeight - 1, 0, 0, bCheck.ScaleWidth, bCheck.ScaleHeight, vbSrcCopy
        StretchBlt PictTemp.hdc, 0, 0, PictTemp.ScaleWidth, PictTemp.ScaleHeight, bCheck.hdc, 0, 0, bCheck.ScaleWidth, bCheck.ScaleHeight, vbSrcCopy
    End If
Else
    If wcolour.Checked Then
        PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), WhiteColour, BF
    Else
        'PictTemp.PaintPicture wCheck, 0, 0, PictTemp.ScaleWidth, PictTemp.ScaleHeight - 1, 0, 0, wCheck.ScaleWidth, wCheck.ScaleHeight, vbSrcCopy
        StretchBlt PictTemp.hdc, 0, 0, PictTemp.ScaleWidth, PictTemp.ScaleHeight, wCheck.hdc, 0, 0, wCheck.ScaleWidth, wCheck.ScaleHeight, vbSrcCopy
    End If
End If
End Sub

Sub SaveGame(TheFile As String)
Dim xx, yy As Integer

If UCase(Mid(TheFile, Len(TheFile) - 3, 4)) <> ".CSH" Then
    TheFile = TheFile & ".CSH"
End If

Open TheFile For Output As #1
For yy = 1 To 8
    For xx = 1 To 8
        If Len(Trim(Str(xyPcs(xx, yy)))) = 1 Then
            Print #1, "0" & Trim(Str(xyPcs(xx, yy)));
        Else
            Print #1, Trim(Str(xyPcs(xx, yy)));
        End If
    Next
Print #1,
Next

Print #1, BlackPeices
For xx = 0 To BlackPeices - 1
    If Len(Trim(Str(rBpeices(xx)))) = 1 Then
        Print #1, "0" & Trim(Str(rBpeices(xx)));
    Else
        Print #1, Trim(Str(rBpeices(xx)));
    End If
Next
Print #1,
Print #1, WhitePeices
For xx = 0 To WhitePeices - 1
    If Len(Trim(Str(rWpeices(xx)))) = 1 Then
        Print #1, "0" & Trim(Str(rWpeices(xx)));
    Else
        Print #1, Trim(Str(rWpeices(xx)));
    End If
Next
Print #1,
Print #1, Label2.Caption

Print #1, Form3.Text1.Text
Close #1
End Sub

Sub OpenGame(TheFile As String)
If UCase(Mid(TheFile, Len(TheFile) - 3, 4)) <> ".CSH" Then
    TheFile = TheFile & ".CSH"
End If

Dim xcnt2, ycnt2 As Integer
Dim gline As String * 16
Dim glinetmp As Integer
Dim counting As Integer

Form3.Text1.Text = ""

Open TheFile For Input As #1

Dim tt As Integer
'Clear out taken peices
For tt = 0 To 15
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vbButtonFace, BF
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vb3DShadow, B
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, 0), vb3DHighlight
    PictTemp.Line (0, 0)-(0, PictTemp.ScaleHeight), vb3DHighlight
    Bpeices(tt).Picture = PictTemp.Image
    Wpeices(tt).Picture = PictTemp.Image
Next

For ycnt2 = 1 To 8
    Line Input #1, gline
For xcnt2 = 1 To 8

glinetmp = Val(Mid(gline, (xcnt2 * 2) - 1, 2))

Board(counting).DragIcon = Nothing

If wColor(glinetmp) = pBlack Then

If glinetmp = bPawn Then
    temp.Picture = LoadResPicture("bPawn", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("bPawn", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = bRook Then
    temp.Picture = LoadResPicture("bCastle", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("bCastle", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = bBishop Then
    temp.Picture = LoadResPicture("bBishop", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("bBishop", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = bKnight Then
    temp.Picture = LoadResPicture("bHorse", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("bHorse", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = bQueen Then
    temp.Picture = LoadResPicture("bQueen", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("bQueen", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = bKing Then
    temp.Picture = LoadResPicture("bKing", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("bKing", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
End If

Call DrawCheck(xcnt2, ycnt2)

DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3

Board(counting).Picture = PictTemp.Image
Else

If glinetmp = wPawn Then
    temp.Picture = LoadResPicture("wPawn", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("wPawn", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = wRook Then
    temp.Picture = LoadResPicture("wCastle", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("wCastle", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = wBishop Then
    temp.Picture = LoadResPicture("wBishop", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("wBishop", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = wKnight Then
    temp.Picture = LoadResPicture("wHorse", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("wHorse", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = wQueen Then
    temp.Picture = LoadResPicture("wQueen", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("wQueen", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
ElseIf glinetmp = wKing Then
    temp.Picture = LoadResPicture("wKing", vbResIcon)
'    Board(counting).DragIcon = LoadResPicture("wKing", vbResIcon)
    Board(counting).DragIcon = LoadResPicture(101, vbResIcon)
End If

Call DrawCheck(xcnt2, ycnt2)

DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3

Board(counting).Picture = PictTemp.Image
End If

If glinetmp = 0 Then
    Call DrawCheck(xcnt2, ycnt2)
    Board(counting).Picture = PictTemp.Image
End If


xyPcs(xcnt2, ycnt2) = glinetmp
counting = counting + 1
Next
Next

Dim BP As Integer
Dim WP As Integer
Dim BPs As String
Dim WPs As String
Line Input #1, gline
BP = Val(gline)
BlackPeices = BP
Line Input #1, gline
BPs = gline
Line Input #1, gline
WP = Val(gline)
WhitePeices = WP
Line Input #1, gline
WPs = gline
Call DrawRest(BP, BPs, WP, WPs)

Line Input #1, gline
Label2.Caption = Trim(gline)

Do While Not EOF(1)
Line Input #1, gline
If Trim(gline) <> "" Then
Form3.Text1.Text = Form3.Text1.Text & gline & vbCrLf
End If
Loop

Close #1
End Sub

Sub DisableSizing(frm As Form)
Dim hMenu, nCount, IDnumber As Long
hMenu = GetSystemMenu(frm.hwnd, 0)
nCount = GetMenuItemCount(hMenu)
Call RemoveMenu(hMenu, nCount - 5, MF_REMOVE Or MF_BYPOSITION)
End Sub

Sub DrawRest(BP As Integer, BPs As String, WP As Integer, WPs As String)
Dim cnt As Integer
Dim TmpStore As Integer
Dim FindP As Integer

For cnt = 1 To BP
FindP = (cnt * 2) - 1
TmpStore = Val(Mid(BPs, FindP, 2))

If TmpStore = bPawn Then
    temp.Picture = LoadResPicture("bPawn", vbResIcon)
ElseIf TmpStore = bRook Then
    temp.Picture = LoadResPicture("bCastle", vbResIcon)
ElseIf TmpStore = bBishop Then
    temp.Picture = LoadResPicture("bBishop", vbResIcon)
ElseIf TmpStore = bKnight Then
    temp.Picture = LoadResPicture("bHorse", vbResIcon)
ElseIf TmpStore = bQueen Then
    temp.Picture = LoadResPicture("bQueen", vbResIcon)
ElseIf TmpStore = bKing Then
    temp.Picture = LoadResPicture("bKing", vbResIcon)
End If

PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vbButtonFace, BF
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vb3DShadow, B
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, 0), vb3DHighlight
PictTemp.Line (0, 0)-(0, PictTemp.ScaleHeight), vb3DHighlight

DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3

Bpeices(cnt - 1).Picture = PictTemp.Image
rBpeices(cnt - 1) = TmpStore
Next

For cnt = 1 To WP
FindP = (cnt * 2) - 1
TmpStore = Val(Mid(WPs, FindP, 2))

If TmpStore = wPawn Then
    temp.Picture = LoadResPicture("wPawn", vbResIcon)
ElseIf TmpStore = wRook Then
    temp.Picture = LoadResPicture("wCastle", vbResIcon)
ElseIf TmpStore = wBishop Then
    temp.Picture = LoadResPicture("wBishop", vbResIcon)
ElseIf TmpStore = wKnight Then
    temp.Picture = LoadResPicture("wHorse", vbResIcon)
ElseIf TmpStore = wQueen Then
    temp.Picture = LoadResPicture("wQueen", vbResIcon)
ElseIf TmpStore = wKing Then
    temp.Picture = LoadResPicture("wKing", vbResIcon)
End If

PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vbButtonFace, BF
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vb3DShadow, B
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, 0), vb3DHighlight
PictTemp.Line (0, 0)-(0, PictTemp.ScaleHeight), vb3DHighlight

DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3

Wpeices(cnt - 1).Picture = PictTemp.Image
rWpeices(cnt - 1) = TmpStore
Next
End Sub

Sub DrawFinishedPeice(Index, xcnt, ycnt, xcnt2, ycnt2)
If wColor(xyPcs(xcnt, ycnt)) <> wColor(xyPcs(xcnt2, ycnt2)) Then
If wColor(xyPcs(xcnt2, ycnt2)) = pBlack Then
If xyPcs(xcnt2, ycnt2) = bPawn Then
    temp.Picture = LoadResPicture("bPawn", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = bRook Then
    temp.Picture = LoadResPicture("bCastle", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = bBishop Then
    temp.Picture = LoadResPicture("bBishop", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = bKnight Then
    temp.Picture = LoadResPicture("bHorse", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = bQueen Then
    temp.Picture = LoadResPicture("bQueen", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = bKing Then
    temp.Picture = LoadResPicture("bKing", vbResIcon)
End If

PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vbButtonFace, BF
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vb3DShadow, B
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, 0), vb3DHighlight
PictTemp.Line (0, 0)-(0, PictTemp.ScaleHeight), vb3DHighlight

DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3

Bpeices(BlackPeices).Picture = PictTemp.Image
rBpeices(BlackPeices) = xyPcs(xcnt2, ycnt2)
BlackPeices = BlackPeices + 1
Else

If xyPcs(xcnt2, ycnt2) = wPawn Then
    temp.Picture = LoadResPicture("wPawn", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = wRook Then
    temp.Picture = LoadResPicture("wCastle", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = wBishop Then
    temp.Picture = LoadResPicture("wBishop", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = wKnight Then
    temp.Picture = LoadResPicture("wHorse", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = wQueen Then
    temp.Picture = LoadResPicture("wQueen", vbResIcon)
ElseIf xyPcs(xcnt2, ycnt2) = wKing Then
    temp.Picture = LoadResPicture("wKing", vbResIcon)
End If

PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vbButtonFace, BF
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vb3DShadow, B
PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, 0), vb3DHighlight
PictTemp.Line (0, 0)-(0, PictTemp.ScaleHeight), vb3DHighlight

DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3

Wpeices(WhitePeices).Picture = PictTemp.Image
rWpeices(WhitePeices) = xyPcs(xcnt2, ycnt2)
WhitePeices = WhitePeices + 1
End If
End If
End Sub

Function CheckIfValid(xx, yy, col, sx, sy) As Boolean
Dim cctmp As Integer
Dim sClear As Boolean
Dim sValid As Boolean
Dim ctmp As Integer
Dim StillBlank As Boolean

'PAWNS information
If xyPcs(xx, yy) = bPawn Then
'BLACK PAWN data here
If (yy = 2 And sy - yy = 2) And sx - xx = 0 Then
    If xyPcs(sx, sy) = 0 And xyPcs(sx, sy - 1) = 0 Then
        CheckIfValid = True
        Exit Function
    End If
ElseIf sy - yy = 1 And sx - xx = 0 Then
    If xyPcs(sx, sy) = 0 Then
        CheckIfValid = True
        Exit Function
End If
End If

If sy - yy = 1 And sx - xx = 1 Then
    If wColor(xyPcs(sx, sy)) = pWhite Then
        CheckIfValid = True
        Exit Function
    End If
ElseIf sy - yy = 1 And sx - xx = -1 Then
    If wColor(xyPcs(sx, sy)) = pWhite Then
    CheckIfValid = True
    Exit Function
    End If
End If

ElseIf xyPcs(xx, yy) = wPawn Then
'WHITE PAWN data here
If (yy = 7 And yy - sy = 2) And sx - xx = 0 Then
    If xyPcs(sx, sy) = 0 And xyPcs(sx, sy + 1) = 0 Then
        CheckIfValid = True
        Exit Function
    End If
ElseIf yy - sy = 1 And xx - sx = 0 Then
    If xyPcs(sx, sy) = 0 Then
        CheckIfValid = True
        Exit Function
    End If
End If

If yy - sy = 1 And xx - sx = 1 Then
    If wColor(xyPcs(sx, sy)) = pBlack Then
        CheckIfValid = True
        Exit Function
    End If
ElseIf yy - sy = 1 And xx - sx = -1 Then
    If wColor(xyPcs(sx, sy)) = pBlack Then
        CheckIfValid = True
        Exit Function
    End If
End If
End If

'CASTLE information
If xyPcs(xx, yy) = wRook Or xyPcs(xx, yy) = bRook Then
    If sx - xx = 0 And sy - yy > 0 Then
        StillBlank = True
        For ctmp = yy + 1 To sy
            If xyPcs(xx, ctmp) <> 0 Then
                If wColor(xyPcs(xx, sy)) <> col And ctmp = sy Then
                    CheckIfValid = True
                    Exit Function
                Else
                    StillBlank = False
                    Exit For
                End If
            End If
        Next
ElseIf sx - xx = 0 And sy - yy < 0 Then
    StillBlank = True
    For ctmp = yy - 1 To sy Step -1
        If xyPcs(xx, ctmp) <> 0 Then
            If wColor(xyPcs(xx, sy)) <> col And ctmp = sy Then
                CheckIfValid = True
                Exit Function
            Else
                StillBlank = False
                Exit For
            End If
        End If
    Next
End If

If sx - xx > 0 And sy - yy = 0 Then
    StillBlank = True
    For ctmp = xx + 1 To sx
        If xyPcs(ctmp, yy) <> 0 Then
            If wColor(xyPcs(sx, yy)) <> col And ctmp = sx Then
                CheckIfValid = True
                Exit Function
            Else
                StillBlank = False
                Exit For
            End If
        End If
    Next
ElseIf sx - xx < 0 And sy - yy = 0 Then
    StillBlank = True
    For ctmp = xx - 1 To sx Step -1
        If xyPcs(ctmp, yy) <> 0 Then
            If wColor(xyPcs(sx, yy)) <> col And ctmp = sx Then
                CheckIfValid = True
                Exit Function
            Else
                StillBlank = False
                Exit For
            End If
        End If
    Next
End If

If StillBlank Then
CheckIfValid = True
Exit Function
End If
End If

'HORSE information
If xyPcs(xx, yy) = wKnight Or xyPcs(xx, yy) = bKnight Then
    If sx - xx = 1 And sy - yy = 2 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If
    If sx - xx = -1 And sy - yy = 2 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If
    If sx - xx = 1 And sy - yy = -2 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If
    If sx - xx = -1 And sy - yy = -2 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If

'dif
    If sx - xx = 2 And sy - yy = 1 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If
    If sx - xx = -2 And sy - yy = 1 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If
    If sx - xx = 2 And sy - yy = -1 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If
    If sx - xx = -2 And sy - yy = -1 Then
        If wColor(xyPcs(sx, sy)) <> col Then
            CheckIfValid = True
            Exit Function
        End If
    End If
End If

'BISHOP information
If xyPcs(xx, yy) = bBishop Or xyPcs(xx, yy) = wBishop Then
    If sx - xx = sy - yy And sx - xx > 0 And sy - yy > 0 Then
        sValid = True
        For cctmp = 1 To sx - xx
            If xyPcs(xx + cctmp, yy + cctmp) <> 0 Then
                If cctmp = sx - xx And _
                    wColor(xyPcs(xx, yy)) <> _
                    wColor(xyPcs(xx + cctmp, yy + cctmp)) _
                    Then Exit For
                sClear = True
                Exit For
            End If
        Next
    End If
    If sx - xx < sy - yy And (sx - xx) * -1 = sy - yy And sy - yy > 0 Then
        sValid = True
        For cctmp = 1 To sy - yy
            If xyPcs(xx - cctmp, yy + cctmp) <> 0 Then
                If cctmp = sy - yy And _
                    wColor(xyPcs(xx - cctmp, yy + cctmp)) <> _
                    wColor(xyPcs(xx, yy)) Then Exit For
                sClear = True
                Exit For
            End If
        Next
    End If
    If sy - yy < 0 And sx - xx = sy - yy Then
        sValid = True
        For cctmp = 1 To (sy - yy) * -1
            If xyPcs(xx - cctmp, yy - cctmp) <> 0 Then
                If cctmp = (sy - yy) * -1 And _
                    wColor(xyPcs(xx - cctmp, yy - cctmp)) <> _
                    wColor(xyPcs(xx, yy)) Then Exit For
                sClear = True
            End If
        Next
    End If
    If sy - yy < 0 And sx - xx > 0 And sx - xx = (sy - yy) * -1 Then
        sValid = True
        For cctmp = 1 To sx - xx
            If xyPcs(xx + cctmp, yy - cctmp) <> 0 Then
                If cctmp = sx - xx And _
                    wColor(xyPcs(xx + cctmp, yy - cctmp)) <> _
                    wColor(xyPcs(xx, yy)) Then Exit For
                sClear = True
            End If
        Next
    End If
    If sClear = False And sValid = True Then
        CheckIfValid = True
        Exit Function
    End If
End If

'QUEEN information
If xyPcs(xx, yy) = wQueen Or xyPcs(xx, yy) = bQueen Then
    If sx - xx = sy - yy And sx - xx > 0 And sy - yy > 0 Then
        sValid = True
        For cctmp = 1 To sx - xx
            If xyPcs(xx + cctmp, yy + cctmp) <> 0 Then
                If cctmp = sx - xx And _
                    wColor(xyPcs(xx, yy)) <> _
                    wColor(xyPcs(xx + cctmp, yy + cctmp)) Then Exit For
                sClear = True
                Exit For
            End If
        Next
    End If
    If sx - xx < sy - yy And (sx - xx) * -1 = sy - yy And sy - yy > 0 Then
        sValid = True
        For cctmp = 1 To sy - yy
            If xyPcs(xx - cctmp, yy + cctmp) <> 0 Then
                If cctmp = sy - yy And _
                    wColor(xyPcs(xx - cctmp, yy + cctmp)) <> _
                    wColor(xyPcs(xx, yy)) Then Exit For
                sClear = True
                Exit For
            End If
        Next
    End If
    If sy - yy < 0 And sx - xx = sy - yy Then
        sValid = True
        For cctmp = 1 To (sy - yy) * -1
            If xyPcs(xx - cctmp, yy - cctmp) <> 0 Then
                If cctmp = (sy - yy) * -1 And _
                    wColor(xyPcs(xx - cctmp, yy - cctmp)) <> _
                    wColor(xyPcs(xx, yy)) Then Exit For
                sClear = True
            End If
        Next
    End If
    If sy - yy < 0 And sx - xx > 0 And sx - xx = (sy - yy) * -1 Then
        sValid = True
        For cctmp = 1 To sx - xx
            If xyPcs(xx + cctmp, yy - cctmp) <> 0 Then
                If cctmp = sx - xx And _
                    wColor(xyPcs(xx + cctmp, yy - cctmp)) <> _
                    wColor(xyPcs(xx, yy)) Then Exit For
                sClear = True
            End If
        Next
    End If

    If sClear = False And sValid = True Then
        CheckIfValid = True
        Exit Function
    End If

'CASTLE information for QUEEN
    If sx - xx = 0 And sy - yy > 0 Then
        StillBlank = True
        For ctmp = yy + 1 To sy
            If xyPcs(xx, ctmp) <> 0 Then
                If wColor(xyPcs(xx, sy)) <> col _
                    And ctmp = sy Then
                CheckIfValid = True
                Exit Function
                Else
                    StillBlank = False
                    Exit For
                End If
            End If
        Next
    ElseIf sx - xx = 0 And sy - yy < 0 Then
        StillBlank = True
        For ctmp = yy - 1 To sy Step -1
            If xyPcs(xx, ctmp) <> 0 Then
                If wColor(xyPcs(xx, sy)) <> col _
                    And ctmp = sy Then
                CheckIfValid = True
                Exit Function
                Else
                    StillBlank = False
                    Exit For
                End If
            End If
        Next
    End If
    If sx - xx > 0 And sy - yy = 0 Then
        StillBlank = True
        For ctmp = xx + 1 To sx
            If xyPcs(ctmp, yy) <> 0 Then
                If wColor(xyPcs(sx, yy)) <> col _
                    And ctmp = sx Then
                CheckIfValid = True
                Exit Function
                Else
                    StillBlank = False
                    Exit For
                End If
            End If
        Next
    ElseIf sx - xx < 0 And sy - yy = 0 Then
        StillBlank = True
        For ctmp = xx - 1 To sx Step -1
            If xyPcs(ctmp, yy) <> 0 Then
                If wColor(xyPcs(sx, yy)) <> col _
                    And ctmp = sx Then
                CheckIfValid = True
                Exit Function
                Else
                    StillBlank = False
                    Exit For
                End If
            End If
        Next
    End If

    If StillBlank Then
        CheckIfValid = True
        Exit Function
        End If
    End If


'King Information
If xyPcs(xx, yy) = wKing Or xyPcs(xx, yy) = bKing Then

If sx - xx = 0 And sy - yy = -1 Then
'Debug.Print "0 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

If sx - xx = 1 And sy - yy = -1 Then
'Debug.Print "45 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

If sx - xx = 1 And sy - yy = 0 Then
'Debug.Print "90 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

If sx - xx = 1 And sy - yy = 1 Then
'Debug.Print "135 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

If sx - xx = 0 And sy - yy = 1 Then
'Debug.Print "180 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

If sx - xx = -1 And sy - yy = 1 Then
'Debug.Print "225 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

If sx - xx = -1 And sy - yy = 0 Then
'Debug.Print "270 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

If sx - xx = -1 And sy - yy = -1 Then
'Debug.Print "315 Degrees"
    If wColor(xyPcs(sx, sy)) <> col Then
        CheckIfValid = True
        Exit Function
    End If
End If

End If
End Function

Private Sub about_Click()
GoingAbout = True
frmAbout.Show 1, Me
End Sub


Private Sub bcolour_Click()
On Error GoTo nearend
LoadSave.flags = cdlCCRGBInit Or cdlCCFullOpen
LoadSave.Color = BlackColour
LoadSave.ShowColor

BlackColour = LoadSave.Color

If bcolour.Checked = False Then
bcolour.Checked = True
btexture.Checked = False
End If

SaveGame "C:\GangGreen.csh"
OpenGame "C:\GangGreen.csh"
Kill "C:\GangGreen.csh"

Exit Sub
nearend:
End Sub

Private Sub Board_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
Dim count, xcnt, ycnt As Integer
Dim xcnt2, ycnt2 As Integer
Dim Changed As Boolean

Draging = False
Picture1.Picture = Image1.Picture
BitBlt wDc, 0, 0, Width / 15, Height / 15, Picture1.hdc, 0, 0, vbSrcCopy
Picture1.Picture = LoadPicture("")

ycnt = 1
For count = 0 To ClickPeice
    If xcnt = 8 Then
        xcnt = 0
        ycnt = ycnt + 1
    End If
    xcnt = xcnt + 1
Next

ycnt2 = 1
For count = 0 To Index
    If xcnt2 = 8 Then
        xcnt2 = 0
        ycnt2 = ycnt2 + 1
    End If
    xcnt2 = xcnt2 + 1
Next

Dim test As Boolean
'NEW NEW NEW NEW
test = CheckIfValid(xcnt, ycnt, wColor(xyPcs(xcnt, ycnt)), xcnt2, ycnt2)
'test = True


'New
Dim TmpxyPcs(9, 9) As Integer
Dim xtmp, ytmp As Integer
Dim IsInCheck As Boolean

For ytmp = 0 To 9
    For xtmp = 0 To 9
        TmpxyPcs(xtmp, ytmp) = xyPcs(xtmp, ytmp)
    Next
Next

If test Then
xyPcs(xcnt2, ycnt2) = xyPcs(xcnt, ycnt)
xyPcs(xcnt, ycnt) = 0

'NEW NEW NEW NEW NEW
If Label2.Caption = "White" Then
    IsInCheck = IsCheck(pWhite)
Else
    IsInCheck = IsCheck(pBlack)
End If

For ytmp = 0 To 9
    For xtmp = 0 To 9
        xyPcs(xtmp, ytmp) = TmpxyPcs(xtmp, ytmp)
    Next
Next

If IsInCheck Then
    MsgBox "Can't move here, this would put you in check!", , "Warning"
    test = False
End If
End If
'NEW


If test Then
Form3.Text1.Text = Form3.Text1.Text & _
            Chr$(64 + xcnt) & Trim(Str(9 - ycnt)) & _
            " - " & _
            Chr$(64 + xcnt2) & Trim(Str(9 - ycnt2)) & _
            vbCrLf
End If

If test = True And xyPcs(xcnt2, ycnt2) <> 0 Then
    Call DrawFinishedPeice(Index, xcnt, ycnt, xcnt2, ycnt2)
End If

Dim tmpPeice As Integer
If PawntoQueen.Checked Then
    If xyPcs(xcnt, ycnt) = bPawn And ycnt2 = 8 And test = True Then
        Changed = True
        tmpPeice = bQueen
'        Board(Index).DragIcon = LoadResPicture("bQueen", vbResIcon)
        Board(Index).DragIcon = LoadResPicture(101, vbResIcon)
    ElseIf xyPcs(xcnt, ycnt) = wPawn And ycnt2 = 1 And test = True Then
        Changed = True
        tmpPeice = wQueen
'        Board(Index).DragIcon = LoadResPicture("wQueen", vbResIcon)
        Board(Index).DragIcon = LoadResPicture(101, vbResIcon)
    Else
    tmpPeice = xyPcs(xcnt, ycnt)
    End If
Else
    tmpPeice = xyPcs(xcnt, ycnt)
End If

If wColor(tmpPeice) = pBlack Then
If tmpPeice = bPawn Then
    temp.Picture = LoadResPicture("bPawn", vbResIcon)
ElseIf tmpPeice = bRook Then
    temp.Picture = LoadResPicture("bCastle", vbResIcon)
ElseIf tmpPeice = bBishop Then
    temp.Picture = LoadResPicture("bBishop", vbResIcon)
ElseIf tmpPeice = bKnight Then
    temp.Picture = LoadResPicture("bHorse", vbResIcon)
ElseIf tmpPeice = bQueen Then
    temp.Picture = LoadResPicture("bQueen", vbResIcon)
ElseIf tmpPeice = bKing Then
    temp.Picture = LoadResPicture("bKing", vbResIcon)
End If

ElseIf wColor(tmpPeice) = pWhite Then

If tmpPeice = wPawn Then
    temp.Picture = LoadResPicture("wPawn", vbResIcon)
ElseIf tmpPeice = wRook Then
    temp.Picture = LoadResPicture("wCastle", vbResIcon)
ElseIf tmpPeice = wBishop Then
    temp.Picture = LoadResPicture("wBishop", vbResIcon)
ElseIf tmpPeice = wKnight Then
    temp.Picture = LoadResPicture("wHorse", vbResIcon)
ElseIf tmpPeice = wQueen Then
    temp.Picture = LoadResPicture("wQueen", vbResIcon)
ElseIf tmpPeice = wKing Then
    temp.Picture = LoadResPicture("wKing", vbResIcon)
End If
End If

If test Then
    Call DrawCheck(xcnt2, ycnt2)
End If

DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3


If test Then
Board(Index).Picture = PictTemp.Image
'If Mid(xyPcs(xcnt2, ycnt2), 2, 1) = "K" Then
'Dim message As String
'message = Label2.Caption & " is the winner!"
'Call MsgBox(message, vbInformation, "Winner")
'Winner = True
'Exit Sub
'End If

xyPcs(xcnt2, ycnt2) = tmpPeice
If Not Changed Then
    Board(Index).DragIcon = Board(ClickPeice).DragIcon
End If
Board(ClickPeice).DragIcon = Nothing
xyPcs(xcnt, ycnt) = 0
Else
Board(ClickPeice).Picture = PictTemp.Image
End If

Dim Checked As Boolean

If Label2.Caption = "Black" And test = True Then
    Label2.Caption = "White"
    Checked = IsCheck(pWhite)
ElseIf Label2.Caption = "White" And test = True Then
    Label2.Caption = "Black"
    Checked = IsCheck(pBlack)
End If

If Checked Then
MsgBox "Check or Checkmate, be careful!!", vbExclamation, "Warning"
End If

HasCleared = False
End Sub

Private Sub Board_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
'Me.Caption = State & " " & Source
Dim count, xcnt, ycnt As Integer
Dim xcnt2, ycnt2 As Integer

ycnt = 1
For count = 0 To Index
    If xcnt = 8 Then
        xcnt = 0
        ycnt = ycnt + 1
    End If
    xcnt = xcnt + 1
Next
    
ycnt2 = 1
For count = 0 To ClickPeice
    If xcnt2 = 8 Then
        xcnt2 = 0
        ycnt2 = ycnt2 + 1
    End If
    xcnt2 = xcnt2 + 1
Next
   
Dim xyP As POINTAPI
Dim xx, yy As Integer
Call GetCursorPos(xyP)
xx = (xyP.x - left / 15) - 32 / 2
yy = (xyP.y - tOp / 15) - 32 / 2

Dim xad, yad As Long
yad = 0 + GetSystemMetrics(33)
yad = yad + GetSystemMetrics(15)
yad = yad + GetSystemMetrics(4)
xad = 0 + GetSystemMetrics(32)

If HasCleared = False Then
Call DrawCheck(xcnt2, ycnt2)
'Board(ClickPeice).DragIcon = LoadResPicture(101, vbResIcon)
Board(ClickPeice).Picture = PictTemp.Image
'Board(ClickPeice).Refresh

HasCleared = True

If Draging = False Then
'NEW BIT
Picture1.Width = Width / 15
Picture1.Height = Height / 15
BitBlt Picture1.hdc, 0, 0, Width / 15, Height / 15, wDc, 0, 0, vbSrcCopy

BitBlt Picture1.hdc, Board(ClickPeice).left + xad, Board(ClickPeice).tOp + yad, Board(0).Width - 1, Board(0).Height - 4, PictTemp.hdc, 0, 0, vbSrcCopy

Image1.Picture = Picture1.Image
DoEvents
Draging = True
'NEW BIT
End If

Timer2.Enabled = True
End If

Call GetCursorPos(xyP)
xx = (xyP.x - left / 15) - 32 / 2
yy = (xyP.y - tOp / 15) - 32 / 2

Picture1.Picture = Image1.Picture
        ImageList_DrawEx _
              brdPeices.hIml, _
              xyPcs(xcnt2, ycnt2) - 1, _
              Picture1.hdc, _
              xx, yy, 0, 0, _
              CLR_NONE, CLR_NONE, _
              ILD_TRANSPARENT Or ILD_BLEND25
BitBlt wDc, 0, 0, Width / 15, Height / 15, Picture1.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub Board_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Winner = True Then Exit Sub

Dim count, xcnt, ycnt As Integer

ycnt = 1
For count = 0 To Index
    If xcnt = 8 Then
        xcnt = 0
        ycnt = ycnt + 1
    End If
xcnt = xcnt + 1
Next

If wColor(xyPcs(xcnt, ycnt)) = pBlack And Label2.Caption = "White" Then Exit Sub
If wColor(xyPcs(xcnt, ycnt)) = pWhite And Label2.Caption = "Black" Then Exit Sub

If xyPcs(xcnt, ycnt) <> 0 Then
    TempDrag.Picture = Board(Index).Picture
    ClickPeice = Index
    Board(Index).Drag 1
End If
End Sub

Private Sub Board_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim count, xcnt, ycnt As Integer

ycnt = 1
For count = 0 To Index
    If xcnt = 8 Then
        xcnt = 0
        ycnt = ycnt + 1
    End If
xcnt = xcnt + 1
Next

Dim first As Integer
Dim second As Integer
Dim ResultT As String

first = wColor(xyPcs(xcnt, ycnt))
second = xyPcs(xcnt, ycnt)

If first = pBlack Then
    ResultT = "Black"
ElseIf first = pWhite Then
    ResultT = "White"
End If

If second = wPawn Or second = bPawn Then
    ResultT = ResultT & " Pawn"
ElseIf second = wRook Or second = bRook Then
    ResultT = ResultT & " Rook"
ElseIf second = wBishop Or second = bBishop Then
    ResultT = ResultT & " Bishop"
ElseIf second = wKnight Or second = bKnight Then
    ResultT = ResultT & " Knight"
ElseIf second = wQueen Or second = bQueen Then
    ResultT = ResultT & " Queen"
ElseIf second = wKing Or second = bKing Then
    ResultT = ResultT & " King"
End If

Info.Caption = ResultT
Timer1.Enabled = True

Exit Sub
If Button = 1 Then
Dim xx, yy As Integer
xx = x  ' - 48 / 2
yy = y  ' - 48 / 2
Picture1.Picture = Image1.Picture
If xx > 0 And yy > 0 And xx < 230 Then
        ImageList_DrawEx _
              brdPeices.hIml, _
              1, _
              Picture1.hdc, _
              xx, yy, 0, 0, _
              CLR_NONE, CLR_NONE, _
              ILD_TRANSPARENT Or ILD_BLEND25

End If
BitBlt wDc, 0, 0, Width / 15, Height / 15, Picture1.hdc, 0, 0, vbSrcCopy
End If
End Sub

Private Sub btexture_Click()
On Error GoTo nearend

Dim s As Boolean
s = ShowForm

If s = False Then
    GoingAbout = True
    Form2.Caption = "Choose Black checker"
    Form2.Show 1, Me
    GoTo skipbit
End If

LoadSave.Filter = "Bitmap Files (*.bmp) | *.bmp"
LoadSave.ShowOpen

bCheck.Picture = LoadPicture(LoadSave.Filename)

skipbit:

If bcolour.Checked Then
bcolour.Checked = False
btexture.Checked = True
End If

SaveGame "C:\GangGreen.csh"
OpenGame "C:\GangGreen.csh"
Kill "C:\GangGreen.csh"

LoadSave.Filter = "Chess Files (*.csh) | *.csh"

Exit Sub
nearend:
LoadSave.Filter = "Chess Files (*.csh) | *.csh"
End Sub

Private Sub exit_Click()
Unload Form2
Unload frmAbout
'Unload Form3
Unload Form4
Unload Form1
End Sub

Private Sub Form_Activate()
On Error GoTo nearend

wDc = GetWindowDC(hwnd)

If GoingAbout Then
    GoingAbout = False
    Exit Sub
End If

If Reset = False Then
Exit Sub
End If
Reset = False

Dim col, cc, cctmp, yy, xx As Integer

col = 1

Dim tt As Integer
For tt = 0 To 15
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vbButtonFace, BF
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, PictTemp.ScaleHeight - 1), vb3DShadow, B
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth - 1, 0), vb3DHighlight
    PictTemp.Line (0, 0)-(0, PictTemp.ScaleHeight), vb3DHighlight
    Bpeices(tt).Picture = PictTemp.Image
    Wpeices(tt).Picture = PictTemp.Image
Next

WhitePeices = 0
BlackPeices = 0
Label2.Caption = "White"

For yy = 1 To 8
For xx = 1 To 8
 
  If col = 0 Then
If btexture.Checked = False Then
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth, PictTemp.ScaleHeight), BlackColour, BF
Else
    StretchBlt PictTemp.hdc, 0, 0, PictTemp.ScaleWidth, PictTemp.ScaleHeight, bCheck.hdc, 0, 0, bCheck.ScaleWidth, bCheck.ScaleHeight, vbSrcCopy
End If
    xyCol(xx, yy) = True
    col = 1
  ElseIf col = 1 Then
If wtexture.Checked = False Then
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth, PictTemp.ScaleHeight), WhiteColour, BF
Else
    StretchBlt PictTemp.hdc, 0, 0, PictTemp.ScaleWidth, PictTemp.ScaleHeight, wCheck.hdc, 0, 0, wCheck.ScaleWidth, wCheck.ScaleHeight, vbSrcCopy
End If
    xyCol(xx, yy) = False
    col = 0
  End If

If yy = 2 Then
temp.Picture = LoadResPicture("bPawn", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("bPawn", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = bPawn
End If

If yy = 7 Then
temp.Picture = LoadResPicture("wPawn", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("wPawn", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = wPawn
End If

If (xx = 1 Or xx = 8) And yy = 1 Then
temp.Picture = LoadResPicture("bCastle", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("bCastle", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = bRook
End If

If (xx = 1 Or xx = 8) And yy = 8 Then
temp.Picture = LoadResPicture("wCastle", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("wCastle", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = wRook
End If

If (xx = 2 Or xx = 7) And yy = 1 Then
temp.Picture = LoadResPicture("bHorse", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("bHorse", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = bKnight
End If

If (xx = 2 Or xx = 7) And yy = 8 Then
temp.Picture = LoadResPicture("wHorse", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("wHorse", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = wKnight
End If

If (xx = 3 Or xx = 6) And yy = 1 Then
temp.Picture = LoadResPicture("bBishop", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("bBishop", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = bBishop
End If

If (xx = 3 Or xx = 6) And yy = 8 Then
temp.Picture = LoadResPicture("wBishop", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("wBishop", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = wBishop
End If

If xx = 4 And yy = 1 Then
temp.Picture = LoadResPicture("bQueen", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("bQueen", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = bQueen
End If

If xx = 5 And yy = 1 Then
temp.Picture = LoadResPicture("bKing", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("bKing", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = bKing
End If

If xx = 4 And yy = 8 Then
temp.Picture = LoadResPicture("wQueen", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("wQueen", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = wQueen
End If

If xx = 5 And yy = 8 Then
temp.Picture = LoadResPicture("wKing", vbResIcon)
'Board(cc).DragIcon = LoadResPicture("wKing", vbResIcon)
Board(cc).DragIcon = LoadResPicture(101, vbResIcon)
xyPcs(xx, yy) = wKing
End If


If yy > 2 And yy < 7 Then
    Board(cc).DragIcon = Nothing
    Board(cc).DragMode = 0
    xyPcs(xx, yy) = 0
End If

If yy <= 2 Or yy >= 7 Then
DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (32 / 2), _
    PictTemp.ScaleHeight / 2 - (32 / 2), _
    temp, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3

'This is the old one.
'DrawIconEx PictTemp.hdc, _
    PictTemp.ScaleWidth / 2 - (36 / 2), _
    PictTemp.ScaleHeight / 2 - (36 / 2), _
    temp, 36, 36, ByVal 0&, ByVal 0&, &H8 Or &H3
End If

Board(cc).Picture = PictTemp.Image
cc = cc + 1
Next
  If col = 0 Then
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth, PictTemp.ScaleHeight), BlackColour, BF
    xyCol(xx, yy) = True
    col = 1
  ElseIf col = 1 Then
    PictTemp.Line (0, 0)-(PictTemp.ScaleWidth, PictTemp.ScaleHeight), WhiteColour, BF
    xyCol(xx, yy) = False
    col = 0
  End If
Next

Form1.SetFocus

If Not DoneSizing Then
    DisableSizing Me
    DoneSizing = True
End If

nearend:
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
Call Board_DragDrop(ClickPeice, Source, x, y)
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
'Timer2.Enabled = True

If State = 0 Then
Board(ClickPeice).Picture = TempDrag.Image

temp.Picture = TempDrag.Image
Picture1.Picture = Image1.Picture

Dim xad, yad As Long
yad = 0 + GetSystemMetrics(33)
yad = yad + GetSystemMetrics(15)
yad = yad + GetSystemMetrics(4)
xad = 0 + GetSystemMetrics(32)

BitBlt Picture1.hdc, Board(ClickPeice).left + xad, Board(ClickPeice).tOp + yad, temp.Width - 1, temp.Height - 4, temp.hdc, 0, 0, vbSrcCopy
Image1.Picture = Picture1.Image

Dim xyP As POINTAPI
Dim xx, yy As Integer
Call GetCursorPos(xyP)
xx = (xyP.x - left / 15) - 32 / 2
yy = (xyP.y - tOp / 15) - 32 / 2
Picture1.Picture = Image1.Picture
        ImageList_DrawEx _
              brdPeices.hIml, _
              xyPcs(xcnt2, ycnt2) - 1, _
              Picture1.hdc, _
              xx, yy, 0, 0, _
              CLR_NONE, CLR_NONE, _
              ILD_TRANSPARENT Or ILD_BLEND25


BitBlt wDc, 0, 0, Width / 15, Height / 15, Picture1.hdc, 0, 0, vbSrcCopy

HasCleared = False
'NeedRedraw = True
Draging = False
End If

End Sub

Private Sub Form_Load()
Reset = True

Dim ct As Integer
Dim xx, yy As Integer

BlackColour = QBColor(7)
WhiteColour = QBColor(15)

Info.Caption = ""

'Black Peices
ct = 0
For yy = 1 To 8
    For xx = 1 To 2
        Bpeices(ct).Width = 42 * 15
        Bpeices(ct).Height = 42 * 15
        Bpeices(ct).left = (xx - 1) * 43 * 15
        Bpeices(ct).tOp = (yy - 1) * 43 * 15
        Bpeices(ct).Stretch = False
        ct = ct + 1
    Next
Next

Frame1.Width = (43 * 2)
Frame1.Height = (43 * 8)

ct = 0
For yy = 1 To 8
    For xx = 1 To 2
        Wpeices(ct).Width = 42 * 15
        Wpeices(ct).Height = 42 * 15
        Wpeices(ct).left = (xx - 1) * 43 * 15
        Wpeices(ct).tOp = (yy - 1) * 43 * 15
        Wpeices(ct).Stretch = False
        ct = ct + 1
    Next
Next

Frame2.Width = (43 * 2)
Frame2.Height = (43 * 8)

Dim r As Boolean
r = GetSetting("Chess", "Settings", "Save_on_Exit", False)
saveonexit.Checked = r
r = GetSetting("Chess", "Settings", "Taken_Peices", True)
If r = True Then viewpeices_Click
r = GetSetting("Chess", "Settings", "View_Moves", False)
If r = True Then viewmoves_Click
r = GetSetting("Chess", "Settings", "View_Shadow", True)
If r = False Then viewshadow_Click
r = GetSetting("Chess", "Settings", "View_Info", True)
If r = False Then viewinfo_Click

Dim rX As Long
rX = GetSetting("Chess", "Settings", "cLeft", (Screen.Width / 2) - (Width / 2))
Me.left = rX
rX = GetSetting("Chess", "Settings", "cTop", (Screen.Height / 2) - (Height / 2))
Me.tOp = rX

If viewmoves.Checked Then
Load Form3
Form3.Visible = True

rX = GetSetting("Chess", "Settings", "mLeft", Form1.left - Form3.Width)
Form3.left = rX
rX = GetSetting("Chess", "Settings", "mTop", Form1.tOp + 50)
Form3.tOp = rX
End If

'Center form
'Me.Left = (Screen.Width / 2) - (Width / 2)
'Me.Top = (Screen.Height / 2) - (Height / 2)
   
Call CreateMenus(Me.hwnd)
OldWindowProc = SetWindowLong(Me.hwnd, _
        GWL_WNDPROC, AddressOf NewWindowProc)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Info.Caption = ""
Timer1.Enabled = False
End Sub

Private Sub Form_Resize()
Dim ct As Integer
Dim xx, yy As Integer

If WindowState = 1 Then Exit Sub

Me.Visible = False

For yy = 1 To 8
For xx = 1 To 8
Board(ct).Width = 42
Board(ct).Height = 42
Board(ct).left = (xx - 1) * 41 + (Me.ScaleWidth / 2) - ((8 * 41) / 2)
Board(ct).tOp = (yy - 1) * 41 + (Me.ScaleHeight / 2) - ((8 * 41) / 2)
ct = ct + 1
Next
Next

Shape1.Move Board(7).left + Board(7).Width, _
            Board(7).tOp + 10, _
            10, _
            (Board(63).tOp + Board(63).Height) - (Board(7).tOp + 4)

Shape2.Move Board(56).left + 10, _
            Board(56).tOp + Board(56).Height, _
            Board(63).left + Board(63).Width - (Board(56).left), _
            10

Shape3.Move Board(0).left - 1, Board(0).tOp - 1, _
            42 * 8 - 5, _
            42 * 8 - 5

Line (0, 0)-(ScaleWidth, 0), vb3DShadow
Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
Refresh

Label1.left = Board(0).left + 10
Label2.left = Label1.left + Label1.Width + 5
Info.left = Board(6).left

'If viewmoves.Checked Then
'Text1.Width = ScaleWidth
'Text1.Height = ScaleHeight
'End If

Me.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If saveonexit.Checked Then
Call SaveSetting("Chess", "Settings", "Save_on_Exit", saveonexit.Checked)
Call SaveSetting("Chess", "Settings", "Taken_Peices", viewpeices.Checked)
Call SaveSetting("Chess", "Settings", "View_Moves", viewmoves.Checked)
Call SaveSetting("Chess", "Settings", "View_Shadow", viewshadow.Checked)
Call SaveSetting("Chess", "Settings", "View_Info", viewinfo.Checked)
Call SaveSetting("Chess", "Settings", "cLeft", Me.left)
Call SaveSetting("Chess", "Settings", "cTop", Me.tOp)
Call SaveSetting("Chess", "Settings", "mLeft", Form3.left)
Call SaveSetting("Chess", "Settings", "mTop", Form3.tOp)
End If

Unload Form2
Unload Form3
Unload frmAbout
Call OnDestroy
End Sub

Private Sub Info_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Info.Caption = ""
Timer1.Enabled = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Info.Caption = ""
Timer1.Enabled = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Info.Caption = ""
Timer1.Enabled = False
End Sub

Private Sub newgame_Click()
HasCleared = False
Winner = False
Form3.Text1.Text = ""
Reset = True
Form_Activate
End Sub

Private Sub open_Click()
On Error GoTo canceled
LoadSave.ShowOpen

OpenGame LoadSave.Filename
Winner = False
HasCleared = False

Exit Sub
canceled:
End Sub

Private Sub PawntoQueen_Click()
If PawntoQueen.Checked Then
PawntoQueen.Checked = False
Else
PawntoQueen.Checked = True
End If
End Sub

Private Sub save_Click()
On Error GoTo canceled
LoadSave.ShowSave

SaveGame LoadSave.Filename

Exit Sub
canceled:
End Sub

Private Sub saveonexit_Click()
If saveonexit.Checked Then
    saveonexit.Checked = False
Else
    saveonexit.Checked = True
End If
End Sub

Private Sub Timer1_Timer()
Dim mc As POINTAPI
GetCursorPos mc
If WindowFromPoint(mc.x, mc.y) <> Me.hwnd Then
Info.Caption = ""
End If
End Sub

Private Sub Timer2_Timer()
Dim r As Long
r = GetAsyncKeyState(vbKeyLButton)

If Timer2Counter < 3 Then
    Timer2Counter = Timer2Counter + 1
End If

If r <> -32768 And r <> 1 And r <> -32767 Then

If Timer2Counter < 3 Then
    Dim gogo As Integer
    For gogo = 0 To Board.count - 1
        Board(gogo).Refresh
    Next
End If

Timer2.Enabled = False
HasCleared = False
Draging = False
Timer2Counter = 0
End If

Dim mP As POINTAPI
Dim yad As Long
Call GetCursorPos(mP)
yad = 0 + GetSystemMetrics(33)
yad = yad + GetSystemMetrics(15)
yad = yad + GetSystemMetrics(4)

If mP.x < left / 15 Or _
    mP.x > (left + Width) / 15 Or _
    mP.y < (tOp / 15) + yad Or _
    mP.y > (tOp + Height) / 15 Then
'Debug.Print "Out " & Rnd
Form_DragOver Board(ClickPeice), -10, -10, 0
End If
End Sub

Private Sub viewinfo_Click()
If viewinfo.Checked Then
viewinfo.Checked = False
Info.Visible = False
Else
viewinfo.Checked = True
Info.Visible = True
End If
End Sub

Private Sub viewmoves_Click()
If viewmoves.Checked Then
viewmoves.Checked = False
'Text1.Visible = False
Form3.Visible = False
Else
viewmoves.Checked = True
'Text1.Left = 0
'Text1.Top = 0
'Text1.Width = ScaleWidth
'Text1.Height = ScaleHeight
'Text1.Visible = True
Form3.Visible = True
End If
End Sub

Private Sub viewpeices_Click()
If viewpeices.Checked = False Then
Me.Width = Me.Width + (42 * 4 * 15)
Frame1.left = 10
Frame1.tOp = Board(0).tOp
Frame1.Visible = True
Frame2.left = Me.ScaleWidth - Frame2.Width - 10
Frame2.tOp = Board(0).tOp
Frame2.Visible = True
viewpeices.Checked = True
Else
Me.Width = Me.Width - (42 * 4 * 15)
Frame1.Visible = False
Frame2.Visible = False
viewpeices.Checked = False
End If
End Sub

Private Sub viewshadow_Click()
If viewshadow.Checked Then
    viewshadow.Checked = False
    Shape1.Visible = False
    Shape2.Visible = False
Else
    viewshadow.Checked = True
    Shape1.Visible = True
    Shape2.Visible = True
End If
End Sub

Private Sub wcolour_Click()
On Error GoTo nearend
LoadSave.flags = cdlCCRGBInit Or cdlCCFullOpen
LoadSave.Color = WhiteColour
LoadSave.ShowColor

WhiteColour = LoadSave.Color

If wcolour.Checked = False Then
wcolour.Checked = True
wtexture.Checked = False
End If

SaveGame "C:\GangGreen.csh"
OpenGame "C:\GangGreen.csh"
Kill "C:\GangGreen.csh"

Exit Sub
nearend:
End Sub

Private Sub wtexture_Click()
On Error GoTo nearend

Dim s As Boolean
s = ShowForm

If s = False Then
    GoingAbout = True
    Form2.Caption = "Choose White checker"
    Form2.Show 1, Me
    GoTo skipbit
End If

LoadSave.Filter = "Bitmap Files (*.bmp) | *.bmp"
LoadSave.ShowOpen
wCheck.Picture = LoadPicture(LoadSave.Filename)

skipbit:

If wcolour.Checked Then
wcolour.Checked = False
wtexture.Checked = True
End If

SaveGame "C:\GangGreen.csh"
OpenGame "C:\GangGreen.csh"
Kill "C:\GangGreen.csh"

LoadSave.Filter = "Chess Files (*.csh) | *.csh"

Exit Sub
nearend:
LoadSave.Filter = "Chess Files (*.csh) | *.csh"
End Sub
