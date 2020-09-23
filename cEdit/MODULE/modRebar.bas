Attribute VB_Name = "modRebar"
Option Explicit
'General declarations we need for the example
Public Const WM_USER = &H400

Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type
Declare Sub GetWindowRect Lib "user32" (ByVal hwnd As Long, _
     lpRect As RECT)

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
     (ByVal dwExStyle As Long, ByVal lpClassName As String, _
     ByVal lpWindowName As String, ByVal dwStyle As Long, _
     ByVal x As Long, ByVal y As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long, _
     ByVal hWndParent As Long, ByVal hMenu As Long, _
     ByVal hInstance As Long, ByVal lpParam As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
     ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
     ByVal hWndNewParent As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal hwnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, lParam As Any) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
     (ByVal hwnd As Long, ByVal lpClassName As String, _
     ByVal nMaxCount As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
     (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
     (ByVal hwnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
Public Const GWL_STYLE = (-16)

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
     ByVal wCmd As Integer) As Long
Public Const GW_CHILD = 5

Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Long
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNTEXT = 18

'WINDOW STYLES
Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_VISIBLE = &H10000000
'--end block--'

