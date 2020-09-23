Attribute VB_Name = "modAPI"
Option Explicit
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


Public Langs As String      'Used to load additional langs into memory
Public FilterB As String    'Used to store the filter
Public AllSupport As String 'Used to store a list of all supported files for
                            'the open dialog.
Public ClrString As String  'This will store in a simple format all extensions
                            'supported by this editor and what language to
                            'assign them to. It's format is like this:
                            'html:html c:c/c++ htm:html
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long


Private Const MF_BYPOSITION = &H400&

Private Const HH_DISPLAY_TOC = &H1
Public StopClose As Boolean

Public Sub DisableX(TheForm As Form)
    '** Description:
    '** Disable X in upper right corner of the form
    Dim lngMenu As Long
    lngMenu = GetSystemMenu(TheForm.hwnd, False)
    DeleteMenu lngMenu, 6, MF_BYPOSITION
End Sub

'----------SHOW HTMLHELP CONTENTS----------'
Public Sub HHShowContents(lhWnd As Long)
    HTMLHelp lhWnd, App.path & "\cEdit.chm" & "", HH_DISPLAY_TOC, 0
End Sub
