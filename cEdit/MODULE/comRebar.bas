Attribute VB_Name = "comRebar"
Option Explicit

'============COMRebar.bas====================
'Visual Basic declarations for Windows
'Rebar common control
'============================================

Public Const REBARCLASSNAME = "ReBarWindow32"

Public Const RBIM_IMAGELIST = &H1

Public Const RBS_TOOLTIPS = &H100
Public Const RBS_VARHEIGHT = &H200
Public Const RBS_BANDBORDERS = &H400
Public Const RBS_FIXEDORDER = &H800

Type REBARINFO
    cbSize As Long
    fMask As Long
    himl As Long
End Type

Public Const RBBS_BREAK = &H1      '// break to new line
Public Const RBBS_FIXEDSIZE = &H2  '// band can't be sized
Public Const RBBS_CHILDEDGE = &H4  '// edge around top & bottom of child
Public Const RBBS_HIDDEN = &H8     '// don't show
Public Const RBBS_NOVERT = &H10    '// don't show when vertical
Public Const RBBS_FIXEDBMP = &H20  '// bitmap doesn't move during resize

Public Const RBBIM_STYLE = &H1
Public Const RBBIM_COLORS = &H2
Public Const RBBIM_TEXT = &H4
Public Const RBBIM_IMAGE = &H8
Public Const RBBIM_CHILD = &H10
Public Const RBBIM_CHILDSIZE = &H20
Public Const RBBIM_SIZE = &H40
Public Const RBBIM_BACKGROUND = &H80
Public Const RBBIM_ID = &H100

Type REBARBANDINFOA
    cbSize As Long
    fMask As Long
    fStyle As Long
    colorFore As Long
    colorBack As Long
    lpText As String
    cch As Long
    iImage As Integer 'Image
    hWndChild As Long
    cxMinChild As Long
    cyMinChild As Long
    cx As Long
    hbmBack As Long 'hBitmap for background
    wID As Long
End Type

Public Const RB_INSERTBANDA = (WM_USER + 1)
Public Const RB_DELETEBAND = (WM_USER + 2)
Public Const RB_GETBARINFO = (WM_USER + 3)
Public Const RB_SETBARINFO = (WM_USER + 4)
Public Const RB_GETBANDINFO = (WM_USER + 5)
Public Const RB_SETBANDINFOA = (WM_USER + 6)
Public Const RB_SETPARENT = (WM_USER + 7)
Public Const RB_INSERTBANDW = (WM_USER + 10)
Public Const RB_SETBANDINFOW = (WM_USER + 11)
Public Const RB_GETBANDCOUNT = (WM_USER + 12)
Public Const RB_GETROWCOUNT = (WM_USER + 13)
Public Const RB_GETROWHEIGHT = (WM_USER + 14)

Public Const RB_INSERTBAND = RB_INSERTBANDA
Public Const RB_SETBANDINFO = RB_SETBANDINFOA

'=======================================
Public hWndRebar As Long 'Rebar's hWnd


Public Property Get BandCount() As Long
     BandCount = SendMessage(hWndRebar, _
          RB_GETBANDCOUNT, 0&, ByVal 0&)
End Property

Public Function CreateCoolbar(ByVal hWndParent As Long, _
ByVal Width As Long, ByVal Height As Long, _
Optional ByVal bVertical As Boolean = False) As Long

     Dim cStyle As Long

     cStyle = WS_CHILD Or WS_BORDER Or _
     WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or _
     WS_VISIBLE Or RBS_VARHEIGHT Or _
     RBS_BANDBORDERS

     If bVertical = True Then _
          cStyle = cStyle Or CCS_VERT

     hWndRebar = CreateWindowEx(0&, _
     REBARCLASSNAME, "", cStyle, 0, 0, Width, _
     Height, hWndParent, ByVal 0&, App.hInstance, ByVal 0&)

     'Check to see if we were successful
     If hWndRebar = 0 Then
          MsgBox "Rebar not created!", vbOKOnly
          CreateCoolbar = 0
          Exit Function
     End If

     CreateCoolbar = hWndRebar
End Function

Public Function RBAddBandByhWnd( _
Optional ByVal CtlChild As Long = 0, _
Optional ByVal BandText As String = "", _
Optional ByVal hBMP As Long = 0, _
Optional ByVal BreakLine As Boolean = True, _
Optional ByVal NoMove As Boolean = False) As Long

     On Error Resume Next

     If hWndRebar = 0 Then
          MsgBox "No hWndRebar!"
          Exit Function
     End If

     Dim ClassName As String
     Dim hWndReal As Long

     Dim Band As REBARBANDINFOA
     Dim rct As Rect

     hWndReal = CtlChild

     If Not (CtlChild = 0) Then
          'Check to see if it's a toolbar (so we can
          'make if flat)
          Band.fMask = RBBIM_CHILD Or RBBIM_CHILDSIZE
          ClassName = Space$(255)
          Call GetClassName(CtlChild, ClassName, 255)
          'see if it's a real Windows toolbar
          If InStr(UCase$(ClassName), "TOOLBARWINDOW32") Then
               SetWindowLong CtlChild, GWL_STYLE, 1442875725
          End If
          'Could be a VB Toolbar -- make it flat anyway.
          If InStr(UCase$(ClassName), "TOOLBARWNDCLASS") Then
               hWndReal = GetWindow(CtlChild, GW_CHILD)
               SetWindowLong hWndReal, GWL_STYLE, 1442875725
          End If
     End If

     Call GetWindowRect(hWndReal, rct)
     rct.bottom = rct.bottom + 2

     If hBMP <> 0 Then _
          Band.fMask = Band.fMask Or RBBIM_BACKGROUND

     Band.fMask = Band.fMask Or RBBIM_STYLE _
          Or RBBIM_ID _
          Or RBBIM_COLORS Or RBBIM_SIZE

     If BandText <> "" Then
          Band.fMask = Band.fMask Or RBBIM_TEXT
     End If

     Band.fStyle = RBBS_CHILDEDGE Or RBBS_FIXEDBMP
     If BreakLine = True Then _
          Band.fStyle = Band.fStyle Or RBBS_BREAK
     If NoMove = True Then
          Band.fStyle = Band.fStyle Or RBBS_FIXEDSIZE
     Else
          Band.fStyle = Band.fStyle And Not RBBS_FIXEDSIZE
     End If

     If BandText <> "" Then Band.lpText = BandText
     If BandText <> "" Then Band.cch = LenB(BandText)
     'Only set if there's a child window
     If hWndReal <> 0 Then
          Band.hWndChild = hWndReal
          Band.cxMinChild = rct.Right - rct.Left
          Band.cyMinChild = rct.bottom - rct.Top
     End If
     'Set the rest OK
     Band.wID = BandCount + 1
     Band.colorBack = GetSysColor(COLOR_BTNFACE)
     Band.colorFore = GetSysColor(COLOR_BTNTEXT)
     Band.cx = 200
     Band.hbmBack = hBMP
     'The length of the type
     Band.cbSize = LenB(Band)

     'non zero (<> 0) means success!
     RBAddBandByhWnd = SendMessage(hWndRebar, RB_INSERTBAND, -1, Band)

     If BandCount < 2 Then
          Call MoveWindow(hWndRebar, 0, 0, 200, 10, True)
     End If
End Function

Public Sub RBRemoveBand(ByVal BandNum As Integer)
     On Error Resume Next
     Call SendMessage(hWndRebar, RB_DELETEBAND, BandNum, 0&)
End Sub
'--end block--'



