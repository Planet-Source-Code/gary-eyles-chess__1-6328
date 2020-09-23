VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Choose check"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   1  'CenterOwner
   Begin Chess.vbalImageList Checkers1 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   42
      IconSizeY       =   42
      ColourDepth     =   16
      Size            =   209304
      Images          =   "Pictures.frx":0000
      KeyCount        =   36
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Okay"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Image Check 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Old As Integer

Private Sub Check_Click(Index As Integer)
Check(Old).BorderStyle = 0
Check(Index).BorderStyle = 1
Old = Index
End Sub

Private Sub Command1_Click()
If Me.Caption = "Choose White checker" Then
Form1.wCheck.Picture = Check(Old).Picture
Else
Form1.bCheck.Picture = Check(Old).Picture
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Old = 0
Check(0).BorderStyle = 1
Me.Width = ((Check(5).Width + Check(5).left) * 15) + ((Me.Width / 15) - Me.ScaleWidth + 2) * 15
End Sub

Private Sub Form_Load()
Dim c As Integer

'For c = 1 To Checkers.ListImages.count
For c = 1 To Checkers1.ImageCount
Load Check(c)
Check(c).Stretch = True
'Check(c - 1).Picture = Checkers.ListImages(c).Picture
Check(c - 1).Picture = Checkers1.ItemPicture(c)

Check(c).left = Check(c - 1).left + Check(c - 1).Width + 1
Check(c).tOp = Check(c - 1).tOp
Check(c).Width = Check(c - 1).Width
Check(c).Height = Check(c - 1).Height

If c Mod 6 = 1 Then
Check(c - 1).left = 0
Check(c - 1).tOp = Check(c - 1).tOp + Check(c - 1).Height + 1
Check(c).left = Check(c - 1).left + Check(c - 1).Width + 1
Check(c).tOp = Check(c - 1).tOp
End If

Next

'For c = 1 To Checkers.ListImages.count
For c = 1 To Checkers1.ImageCount
Check(c - 1).tOp = Check(c - 1).tOp - Check(c - 1).Height
Check(c - 1).Visible = True
Next
End Sub

Private Sub Form_Resize()
Command1.left = (ScaleWidth / 2) - (Command1.Width / 2)
End Sub
