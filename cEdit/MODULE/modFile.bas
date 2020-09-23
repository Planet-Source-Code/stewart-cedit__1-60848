Attribute VB_Name = "modFile"
Option Explicit

Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type

Public Const HKEY_CLASSES_ROOT = &H80000000

Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1                                                   ' Unicode nul terminated string
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Public Function MakeFileType(ByVal Extension As String, ByVal NameOfType As String, ByVal DefaultIcon As String, ByVal NameOfAction As String, ByVal ActionPath As String, Optional ByVal ShellNew As Boolean, Optional ByVal QuickView As Boolean, Optional logs As Boolean) As Boolean
    'On Error GoTo Oops
    Dim dotExtension As String, Extensionfile As String
    Dim correctNameOfAction As String
    Dim writes As String
    dotExtension = "." & Extension
    Extensionfile = Extension & "file"
    correctNameOfAction = ReplaceChars(NameOfAction, " ", "_")
    
    If logs = True Then
      writes = GetString(HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command", "")
      writeini "Reg", Extension & "exe", writes, App.path & "\reg.ini"
      writes = GetString(HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon", "")
      writeini "Reg", Extension & "ico", writes, App.path & "\reg.ini"
      writes = GetString(HKEY_CLASSES_ROOT, Extensionfile, "shell", "")
      writeini "Reg", Extension & "act", writes, App.path & "\reg.ini"
    End If
    
    CreateKey HKEY_CLASSES_ROOT, dotExtension
    CreateKey HKEY_CLASSES_ROOT, Extensionfile
    CreateKey HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon"
    CreateKey HKEY_CLASSES_ROOT, Extensionfile, "Shell"
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell", correctNameOfAction
    CreateKey HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command"
        
        
    SaveString HKEY_CLASSES_ROOT, dotExtension, "", "", Extensionfile
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "", "", NameOfType
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "DefaultIcon", "", DefaultIcon
    SaveString HKEY_CLASSES_ROOT, Extensionfile, "Shell", "", correctNameOfAction
    SaveString HKEY_CLASSES_ROOT, Extensionfile & "\Shell", correctNameOfAction, "", "&" & NameOfAction
    SaveString HKEY_CLASSES_ROOT, Extensionfile & "\Shell\" & correctNameOfAction, "command", "", ActionPath
    
    
    
    
    If Not IsMissing(ShellNew) Then
        EnableShellNew Extension, ShellNew
    End If
    
    If Not IsMissing(QuickView) Then
        EnableQuickView Extension, QuickView
    End If
    MakeFileType = True
    Exit Function
Oops:
    MakeFileType = False
    Exit Function
    Resume Next
End Function


'Sample call:
'    EnableQuickView "txt", True

Public Function EnableQuickView(ByVal Extension As String, ByVal QuickView As Boolean) As Boolean
    On Error GoTo QuickViewOops
    Dim Extensionfile As String
    Extensionfile = Extension & "file"
    
    If QuickView = True Then
        'enable QuickView
        CreateKey HKEY_CLASSES_ROOT, Extensionfile, "QuickView"
        SaveString HKEY_CLASSES_ROOT, Extensionfile, "QuickView", "", "*"
      Else
        'disable QuickView
        DeleteKey HKEY_CLASSES_ROOT, Extensionfile & "\QuickView"
    End If
    
    EnableQuickView = True
    Exit Function
    
QuickViewOops:
    EnableQuickView = False
    Exit Function
    Resume Next
    
End Function


Public Function EnableShellNew(ByVal Extension As String, ByVal ShellNew As Boolean) As Boolean
    On Error GoTo OopsShellN
    Dim dotExtension As String
    dotExtension = "." & Extension
    
    If ShellNew = True Then
        'enable
        CreateKey HKEY_CLASSES_ROOT, dotExtension, "ShellNew"
        SaveString HKEY_CLASSES_ROOT, dotExtension, "ShellNew", "NullFile", ""
      Else
        'disable
        DeleteKey HKEY_CLASSES_ROOT, dotExtension & "\ShellNew"
    End If
    EnableShellNew = True
    Exit Function
    
OopsShellN:
    EnableShellNew = False
    Exit Function
    Resume Next
End Function

'Sample call:
'    Replaced = ReplaceChars("Hello there. Happy New Year", " ", "_")
'   Returns "Hello_there. Happy_New_Year."


' Win32 Registry Access Module
'
' WINREG32.BAS - Copyright <C> 1998, 1999 Randy Mcdowell.
'
' If you modify this code please send me a copy, it's not commented
' really well so you'll have to bear with me here. I have included some
' sample subroutines and  functions to  access the registry. I have  a
' more complex  module  much  more  rich in  code if you want it you
' will need to Email me and ask for it.  mcdowellrandy@hotmail.com.

Private Sub CreateKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant)

    Dim hHnd As Long
    If Not IsMissing(SubKey) Then
        RegCreateKey hKey, Key & "\" & SubKey, hHnd
        RegCloseKey hHnd
    Else
        RegCreateKey hKey, Key, hHnd
        RegCloseKey hHnd
    End If

End Sub

Private Function GetString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As String

    Dim hHnd As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lValueType As Long
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim KeyPath As String
    
    KeyPath = Key + "\" + SubKey
    RegOpenKey hKey, KeyPath, hHnd
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufferSize, " ")
        lResult = RegQueryValueEx(hHnd, ValueName, 0&, 0&, ByVal strBuf, lDataBufferSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        
        End If
    End If
End Function

Public Sub SaveString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As String)

    Dim hHnd As Long
    Dim KeyPath As String
    
    KeyPath = Key + "\" + SubKey
    RegCreateKey hKey, KeyPath, hHnd
    RegSetValueEx hHnd, ValueTitle, 0, REG_SZ, ByVal ValueData, Len(ValueData)
    RegCloseKey hHnd

End Sub

Public Sub DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW")
    '
    RegDeleteKey hKey, strKey
End Sub

