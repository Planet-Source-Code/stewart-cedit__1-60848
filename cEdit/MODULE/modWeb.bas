Attribute VB_Name = "modWeb"
Option Explicit
Public Const SW_SHOWDEFAULT = 10


Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub OpenURL(strURL As String, lngHwnd As Long)
    ShellExecute lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", SW_SHOWDEFAULT
End Sub

