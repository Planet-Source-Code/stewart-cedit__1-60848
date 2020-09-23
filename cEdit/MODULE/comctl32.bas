Attribute VB_Name = "comctl32"
Option Explicit

'============COMCTL32.bas====================
'Visual Basic declarations for Windows
'common controls...
'============================================
Type CommonControlsEx
        dwSize As Long '// size of this structure
        dwICC As Long  '// flags indicating which classes to be initialized
End Type
Public Const ICC_COOL_CLASSES = &H400&   '// rebar (coolbar) control
Declare Function InitCommonControlsEx Lib "COMCTL32" _
     (LPINITCOMMONCONTROLSEX As CommonControlsEx) As Boolean

'//====== Generic Common Control Styles ===========

Public Const CCS_TOP = &H1
Public Const CCS_NOMOVEY = &H2
Public Const CCS_BOTTOM = &H3
Public Const CCS_NORESIZE = &H4
Public Const CCS_NOPARENTALIGN = &H8
Public Const CCS_ADJUSTABLE = &H20
Public Const CCS_NODIVIDER = &H40
Public Const CCS_VERT = &H80
Public Const CCS_LEFT = (CCS_VERT Or CCS_TOP)
Public Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM)
Public Const CCS_NOMOVEX = (CCS_VERT Or CCS_NOMOVEY)


Public Function LoadCommCtls() As Boolean
     Dim ctEx As CommonControlsEx

     ctEx.dwSize = Len(ctEx)
     ctEx.dwICC = ICC_COOL_CLASSES

     LoadCommCtls = InitCommonControlsEx(ctEx)
End Function
'--end block--'


