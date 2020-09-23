Attribute VB_Name = "Module1"
Option Explicit
DefLng A-Z

Global Reset As Boolean

Const MFT_STRING = 0
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCHITTEST = &H84
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCRBUTTONUP = &HA5

'// DrawEdge() constants
'Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Public Const BDR_RAISEDOUTER = &H1
'Public Const BDR_SUNKENOUTER = &H2
'Public Const BDR_RAISEDINNER = &H4
'Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
'Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_DIAGONAL = &H10

Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31

'Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'Public Enum SysColors
'    COLOR_SCROLLBAR = 0
'    COLOR_BACKGROUND = 1
'    COLOR_ACTIVECAPTION = 2
'    COLOR_INACTIVECAPTION = 3
'    COLOR_MENU = 4
'    COLOR_WINDOW = 5
'    COLOR_WINDOWFRAME = 6
'    COLOR_MENUTEXT = 7
'    COLOR_WINDOWTEXT = 8
'    COLOR_CAPTIONTEXT = 9
'    COLOR_ACTIVEBORDER = 10
'    COLOR_INACTIVEBORDER = 11
'    COLOR_APPWORKSPACE = 12
'    COLOR_HIGHLIGHT = 13
'    COLOR_HIGHLIGHTTEXT = 14
'    COLOR_BTNFACE = 15
'    COLOR_BTNSHADOW = 16
'    COLOR_GRAYTEXT = 17
'    COLOR_BTNTEXT = 18
'    COLOR_INACTIVECAPTIONTEXT = 19
'    COLOR_BTNHIGHLIGHT = 20
'End Enum

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_WORD_ELLIPSIS = &H40000
    
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'Type RECT
'    left As Long
'    tOp As Long
'    Right As Long
'    Bottom As Long
'End Type

Type Size
    cx As Long
    cy As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Declare Function GetMenu Lib "user32" _
   (ByVal hwnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" _
   (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function GetMenuItemCount Lib "user32" _
   (ByVal hMenu As Long) As Long

Public Declare Function GetMenuItemInfo Lib "user32" _
    Alias "GetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpmii As MENUITEMINFO) As Long

Declare Function GetMenuItemID Lib "user32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function SetMenuItemInfo Lib "user32" _
    Alias "SetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal uItem As Long, _
    ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

Declare Function AppendMenu Lib "user32" _
    Alias "AppendMenuA" (ByVal hMenu As Long, _
    ByVal wFlags As Long, ByVal wIDNewItem As Long, _
    ByVal lpNewItem As Any) As Long

Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long

Declare Function CreateFont Lib "gdi32" _
    Alias "CreateFontA" (ByVal H As Long, _
    ByVal W As Long, ByVal E As Long, ByVal O As Long, _
    ByVal W As Long, ByVal i As Long, ByVal U As Long, _
    ByVal s As Long, ByVal c As Long, ByVal OP As Long, _
    ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, _
    ByVal F As String) As Long

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

Public Const MF_STRING = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&

Public Const ETO_OPAQUE = 2

Public Const ODS_SELECTED = &H1
Public Const ODS_GRAYED = &H2
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8
Public Const ODS_FOCUS = &H10

Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MENUSELECT = &H11F
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const WM_USER = &H400
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_SYSCOLORCHANGE = &H15

Declare Sub MemCopy Lib "kernel32" Alias _
        "RtlMoveMemory" (dest As Any, src As Any, _
        ByVal numbytes As Long)

Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function TextOut Lib "gdi32" Alias "TextOutA" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal lpString As String, ByVal nCount As Long) As Long

Declare Function ExtTextOut Lib "gdi32" Alias _
    "ExtTextOutA" (ByVal hdc As Long, ByVal x As _
    Long, ByVal y As Long, ByVal wOptions As Long, _
    lpRect As RECT, ByVal lpString As String, _
    ByVal nCount As Long, lpDx As Long) As Long

Declare Function GetDC Lib "user32" _
    (ByVal hwnd As Long) As Long

'Declare Function ReleaseDC Lib "user32" _
    (ByVal hwnd As Long, ByVal hdc As Long) As Long

Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function SetBkColor Lib "gdi32" _
    (ByVal hdc As Long, ByVal crColor As Long) As Long

'Declare Function SetTextColor Lib "gdi32" _
    (ByVal hdc As Long, ByVal crColor As Long) As Long

'Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long

Declare Function GetTextExtentPoint Lib "gdi32" _
    Alias "GetTextExtentPointA" (ByVal hdc As Long, _
    ByVal lpszString As String, ByVal cbString As Long, _
    lpSize As Size) As Long

'Public Const COLOR_MENU = 4
'Public Const COLOR_MENUTEXT = 7
'Public Const COLOR_HIGHLIGHT = 13
'Public Const COLOR_HIGHLIGHTTEXT = 14
'Public Const COLOR_GRAYTEXT = 17

Public Const IDM_CHARACTER = 10
Public Const IDM_REGULAR = 11
Public Const IDM_BOLD = 12
Public Const IDM_ITALIC = 13
Public Const IDM_UNDERLINE = 14

Type myItemType
    hFont As Long
    cchItemText As Integer
    szItemText As String * 32
End Type

Public OldWindowProc
Public hMenu, hSubMenu
Public mnuItemCount, MyItem() As myItemType
Public clrPrevText, clrPrevBkgnd
Public hfntPrev

Public Function ShowForm() As Boolean
Form4.Show 1
If Form4.Tag = "1" Then
    ShowForm = True
Else
    ShowForm = False
End If
Unload Form4
End Function

Public Sub DrawGradient( _
      ByVal hdc As Long, _
      ByRef rct As RECT, _
      ByVal lEndColour As Long, _
      ByVal lStartColour As Long, _
      ByVal bVertical As Boolean _
   )
Dim lStep As Long
Dim lPos As Long, lSize As Long
Dim bRGB(1 To 3) As Integer
Dim bRGBStart(1 To 3) As Integer
Dim dR(1 To 3) As Double
Dim dPos As Double, d As Double
Dim hBr As Long
Dim tR As RECT
   
   LSet tR = rct
   If bVertical Then
      lSize = (tR.Bottom - tR.tOp)
   Else
      lSize = (tR.Right - tR.left)
   End If
   lStep = lSize \ 255
   If (lStep < 3) Then
       lStep = 3
   End If
       
   bRGB(1) = lStartColour And &HFF&
   bRGB(2) = (lStartColour And &HFF00&) \ &H100&
   bRGB(3) = (lStartColour And &HFF0000) \ &H10000
   bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
   dR(1) = (lEndColour And &HFF&) - bRGB(1)
   dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
   dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)
        
   For lPos = lSize To 0 Step -lStep
      ' Draw bar:
      If bVertical Then
         tR.tOp = tR.Bottom - lStep
      Else
         tR.left = tR.Right - lStep
      End If
      If tR.tOp < rct.tOp Then
         tR.tOp = rct.tOp
      End If
      If tR.left < rct.left Then
         tR.left = rct.left
      End If
      
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, tR, hBr
      DeleteObject hBr
            
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tR.Bottom = tR.tOp
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tR.Right = tR.left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If
      
   Next lPos

End Sub

Public Function NewWindowProc(ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, _
    lParam As Long) As Long

    Dim mM As MEASUREITEMSTRUCT
    Dim dM As DRAWITEMSTRUCT

'Debug.Print msg

'If msg = 20 Then
'Debug.Print "HERE " & Rnd
'End If

'If msg = WM_NCMOUSEMOVE Then
'Debug.Print "WM_NCMOUSEMOVE"
'End If

    Select Case msg

        Case WM_DRAWITEM

            MemCopy dM, lParam, Len(dM)
            OnDrawMenuItem hwnd, dM
            
        Case WM_MEASUREITEM

            MemCopy mM, lParam, Len(mM)
            mM = OnMeasureItem(hwnd, mM)
            MemCopy lParam, mM, Len(mM)

        Case WM_COMMAND

           'Put your Menu Command here.

        Case WM_SYSCOLORCHANGE

           'Put your code here.

        Case Else


    End Select

NewWindowProc = CallWindowProc(OldWindowProc, _
hwnd, msg, wParam, VarPtr(lParam))

End Function

Sub CreateMenus(hwnd As Long)

    'get Menus
    hMenu = GetMenu(hwnd)
    hSubMenu = GetSubMenu(hMenu, 0)

    'remove original menu item
'    RemoveMenu hSubMenu, 0, MF_BYPOSITION

    'creates string menus
'    AppendMenu hSubMenu, MF_STRING, IDM_REGULAR, "Regular"
'    AppendMenu hSubMenu, MF_STRING, IDM_BOLD, "Bold"
'    AppendMenu hSubMenu, MF_STRING, IDM_ITALIC, "Italic"
'    AppendMenu hSubMenu, MF_STRING, IDM_UNDERLINE, "Underline"

    'call to make OwnerDrawMenus
    CreateOwnerDrawMenus

End Sub

Sub CreateOwnerDrawMenus()

Dim minfo As MENUITEMINFO, id As Integer
Dim MainMenu, MainMenuCount As Long
Dim MainCount As Integer
Dim cc As Integer
ReDim MyItem(0 To 30) As myItemType
  
  'get the menuitem handle
   MainMenu = GetMenu(Form1.hwnd)
   MainMenuCount = GetMenuItemCount(MainMenu)

For MainCount = 0 To MainMenuCount - 1

   hSubMenu = GetSubMenu(GetMenu(Form1.hwnd), MainCount)
   mnuItemCount = GetMenuItemCount(hSubMenu)
   
'   ReDim MyItem(0 To mnuItemCount - 1) As myItemType
   Dim r As Long

   'loop to fill array
   For id = 0 To mnuItemCount - 1
    minfo.cbSize = Len(minfo)
    minfo.fMask = MIIM_TYPE
    minfo.fType = MFT_STRING
    minfo.dwTypeData = Space$(256)
    minfo.cch = Len(minfo.dwTypeData)

    'get menuitem data
    r = GetMenuItemInfo(hSubMenu, id, True, minfo)

    'and save into user array
    MyItem(cc).cchItemText = minfo.cch 'menuitem length
    MyItem(cc).szItemText = Trim(minfo.dwTypeData) 'text
    'MyItem(0).hFont = CreateMenuItemFont 'font
    'MyItem(id).hFont = CreateMenuItemFont 'font

cc = cc + 1

    'change menu type
If minfo.fType <> 2048 Then
    minfo.fType = MF_OWNERDRAW
    minfo.fMask = MIIM_TYPE Or MIIM_DATA
    minfo.dwItemData = id

    'into MF_OWNERDRAW

    r = SetMenuItemInfo(hSubMenu, id, True, minfo)
End If

   Next

cc = cc + 1
Next

End Sub


Function OnMeasureItem(hwnd As Long, lpmis As MEASUREITEMSTRUCT) As MEASUREITEMSTRUCT
Debug.Print "Measure ";
Debug.Print lpmis.itemID

Dim xM2 As MEASUREITEMSTRUCT

If lpmis.itemID < 8 Then
xM2.itemHeight = 20
xM2.itemWidth = 80
ElseIf lpmis.itemID = 19 Then
xM2.itemHeight = 20
xM2.itemWidth = 120
Else
xM2.itemHeight = 20
xM2.itemWidth = 150
End If

LSet OnMeasureItem = xM2

'Dim lpmis As MEASUREITEMSTRUCT
'On Error Resume Next
'CopyMemory lpmis, ByVal lp, Len(lpmis)
'With lpmis
'    .itemWidth = 200
'    Select Case .itemID
'        Case IDM_SEPARATOR      '// Special case
'            .itemHeight = 2
'        Case Else
'            .itemHeight = m_lngcItemHeight
'    End Select
'End With
'CopyMemory ByVal lp, lpmis, Len(lpmis)

Exit Function

    Dim xM As MEASUREITEMSTRUCT, hfntOld As Long
    Dim s As Size, hdc As Long
    hdc = GetDC(hwnd)
    hfntOld = SelectObject(hdc, MyItem(lpmis.itemData).hFont)
    GetTextExtentPoint hdc, MyItem(lpmis.itemData).szItemText, _
            MyItem(lpmis.itemData).cchItemText, s
    xM.itemWidth = s.cx + 10
    xM.itemHeight = s.cy
    SelectObject hdc, hfntOld
    ReleaseDC hwnd, hdc
    LSet OnMeasureItem = xM
End Function

Sub OnDrawMenuItem(hwnd As Long, lpdis As DRAWITEMSTRUCT)
Dim sText As String
Dim FillColour As Long
Dim MeColor
Dim HiColor

Dim hFont As Long
hFont = SendMessage(Form1.hwnd, WM_GETFONT, 0, ByVal 0&)
Call SelectObject(lpdis.hdc, hFont)

MeColor = GetSysColor(COLOR_BTNFACE)
HiColor = GetSysColor(COLOR_HIGHLIGHT)

    If (lpdis.itemState And ODS_SELECTED) Then 'if selected
        clrPrevText = SetTextColor(lpdis.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT))
'        clrPrevBkgnd = SetBkColor(lpdis.hdc, GetSysColor(COLOR_HIGHLIGHT))
    '    FillColour = GetSysColor(SysColors.COLOR_HIGHLIGHT)
        FillColour = HiColor
        'FillColour = QBColor(14)
    Else
        clrPrevText = SetTextColor(lpdis.hdc, GetSysColor(COLOR_MENUTEXT))
'        clrPrevBkgnd = SetBkColor(lpdis.hdc, GetSysColor(COLOR_MENU))
    '    FillColour = GetSysColor(SysColors.COLOR_BTNFACE)
        FillColour = MeColor
    End If

Dim wPicture As Integer
Dim TmpText As String
TmpText = Mid(MyItem(lpdis.itemID - 2).szItemText, 1, MyItem(lpdis.itemID - 2).cchItemText)
If TmpText = "New" Then
wPicture = 3
ElseIf TmpText = "Open" Then
wPicture = 2
ElseIf TmpText = "Save" Then
wPicture = 1
ElseIf lpdis.itemState = 8 Then
wPicture = 4
ElseIf lpdis.itemState = 9 Then
wPicture = 4
End If

Dim tmprect As RECT
Dim hBr As Long
hBr = CreateSolidBrush(MeColor)
If wPicture > 0 And (lpdis.itemState And ODS_SELECTED) Then lpdis.rcItem.left = lpdis.rcItem.left + 25
If (lpdis.itemState And ODS_SELECTED) Then
DrawGradient lpdis.hdc, lpdis.rcItem, HiColor, MeColor, False
Else
FillRect lpdis.hdc, lpdis.rcItem, hBr
End If
If wPicture > 0 And (lpdis.itemState And ODS_SELECTED) Then lpdis.rcItem.left = lpdis.rcItem.left - 25
    
tmprect = lpdis.rcItem
tmprect.left = tmprect.left + 30
tmprect.tOp = tmprect.tOp + 3

SetBkMode lpdis.hdc, 0
DrawText lpdis.hdc, MyItem(lpdis.itemID - 2).szItemText, -1, tmprect, DT_WORDBREAK

If wPicture > 0 Then
DrawIconEx lpdis.hdc, _
    lpdis.rcItem.left - 4, _
    lpdis.rcItem.tOp - 5, _
    Form1.MnuIcons.ListImages(wPicture).Picture, 32, 32, ByVal 0&, ByVal 0&, &H8 Or &H3
End If

Dim rcSep As RECT
rcSep = lpdis.rcItem
rcSep.left = 2
rcSep.Right = 23
rcSep.Bottom = rcSep.tOp + 20
If (lpdis.itemState And ODS_SELECTED) And wPicture > 0 Then
Call DrawEdge(lpdis.hdc, rcSep, BDR_RAISEDINNER, BF_RECT)
End If

ReleaseDC lpdis.hwndItem, lpdis.hdc
DeleteObject hBr

'    SetTextColor lpdis.hdc, clrPrevText
'    SetBkColor lpdis.hdc, clrPrevBkgnd
End Sub
Function CreateMenuItemFont() As Long
Dim Weight As Long
Dim use_italic As Long
Dim use_underline As Long
Dim use_strikethrough As Long

'   Select Case uID + 11
'        Case IDM_BOLD
'            Weight = 700
'        Case IDM_ITALIC
'            use_italic = True
'        Case IDM_UNDERLINE
'            use_underline = True
'     End Select

'CreateMenuItemFont = CreateFont(20, 0, _
        0, 0, Weight, _
        use_italic, use_underline, _
        use_strikethrough, 136, 0, _
        16, 0, 0, "Times New Roman")

Weight = 0

CreateMenuItemFont = CreateFont(Form1.Font.Size, 0, _
        0, 0, Weight, _
        use_italic, use_underline, _
        use_strikethrough, 136, 0, _
        16, 0, 0, Form1.Font.Name)

End Function

Sub OnDestroy()
Dim r As Long

   'do some clean works
   Dim minfo As MENUITEMINFO, id As Integer

   hSubMenu = GetSubMenu(GetMenu(Form1.hwnd), 0)
   mnuItemCount = GetMenuItemCount(hSubMenu)

  For id = 0 To mnuItemCount - 1
   minfo.fMask = MIIM_DATA

   r = GetMenuItemInfo(hSubMenu, id, True, minfo)

   DeleteObject minfo.dwItemData

   r = SetMenuItemInfo(hSubMenu, id, True, minfo)
  Next

  Erase MyItem

End Sub


