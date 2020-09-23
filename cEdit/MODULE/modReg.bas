Attribute VB_Name = "modReg"
Option Explicit
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_CURRENT_USER = &H80000001


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1 ' Unicode nul terminated String
'MakeFileType "txt", "Text Document", "C:\windows\notepad.exe,0", "open", "C:\windows\notepad.exe %1", False, True

Public Function ReplaceChars(ByVal Text As String, ByVal Char As String, ReplaceChar As String) As String
    Dim counter As Integer
    
    counter = 1
    Do
        counter = InStr(counter, Text, Char)
        If counter <> 0 Then
            Mid(Text, counter, Len(ReplaceChar)) = ReplaceChar
          Else
            ReplaceChars = Text
            Exit Do
        End If
    Loop

    ReplaceChars = Text
End Function


Private Function GetString(hKey As Long, strPath As String, strValue As String, DefaultStr As Long) As String
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    RegOpenKey hKey, strPath, keyhand
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
    If strBuf = "" Then GetString = DefaultStr
End Function


Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    RegCreateKey hKey, strPath, keyhand
    RegSetValueEx keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata)
    RegCloseKey keyhand
End Sub




'This is the section to read\write all the options :)

Public Sub WriteOptions()
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "comment", frmDoc.rt.GetColor(cmClrComment)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "bookmark", frmDoc.rt.GetColor(cmClrBookmark)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "bookmarkbk", frmDoc.rt.GetColor(cmClrBookmarkBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "commentbk", frmDoc.rt.GetColor(cmClrCommentBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "divider", frmDoc.rt.GetColor(cmClrHDividerLines)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "highlight", frmDoc.rt.GetColor(cmClrHighlightedLine)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "keyword", frmDoc.rt.GetColor(cmClrKeyword)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "keywordbk", frmDoc.rt.GetColor(cmClrKeywordBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "left", frmDoc.rt.GetColor(cmClrLeftMargin)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "linenum", frmDoc.rt.GetColor(cmClrLineNumber)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "linenumbk", frmDoc.rt.GetColor(cmClrLineNumberBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "number", frmDoc.rt.GetColor(cmClrNumber)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "numberbk", frmDoc.rt.GetColor(cmClrNumberBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "operator", frmDoc.rt.GetColor(cmClrOperator)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "operatorbk", frmDoc.rt.GetColor(cmClrOperatorBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "scope", frmDoc.rt.GetColor(cmClrScopeKeyword)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "scopebk", frmDoc.rt.GetColor(cmClrScopeKeywordBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "string", frmDoc.rt.GetColor(cmClrString)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "stringbk", frmDoc.rt.GetColor(cmClrStringBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagattrib", frmDoc.rt.GetColor(cmClrTagAttributeName)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagattribbk", frmDoc.rt.GetColor(cmClrTagAttributeNameBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagele", frmDoc.rt.GetColor(cmClrTagElementName)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagelebk", frmDoc.rt.GetColor(cmClrTagElementNameBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagent", frmDoc.rt.GetColor(cmClrTagEntity)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagentbk", frmDoc.rt.GetColor(cmClrTagEntityBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagtxt", frmDoc.rt.GetColor(cmClrTagText)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "tagtxtbk", frmDoc.rt.GetColor(cmClrTagTextBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "text", frmDoc.rt.GetColor(cmClrText)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "textbk", frmDoc.rt.GetColor(cmClrTextBk)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "vdivider", frmDoc.rt.GetColor(cmClrVDividerLines)
  SaveString HKEY_CLASSES_ROOT, "cEdit\colors\", "window", frmDoc.rt.GetColor(cmClrWindow)
  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "selbounds", frmDoc.rt.SelBounds
  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "numbering", frmDoc.rt.LineNumbering
  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "lttips", frmDoc.rt.LineToolTips
  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "numberingstyle", frmDoc.rt.LineNumberStyle
  SaveString HKEY_CLASSES_ROOT, "cEdit\data\", "numberingstart", frmDoc.rt.LineNumberStart
  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "leftmargin", frmDoc.rt.DisplayLeftMargin
  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "bold", frmDoc.rt.Font.Bold
  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "italic", frmDoc.rt.Font.Italic
  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "size", frmDoc.rt.Font.Size
  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "strike", frmDoc.rt.Font.Strikethrough
  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "under", frmDoc.rt.Font.Underline
  SaveString HKEY_CLASSES_ROOT, "cEdit\font\", "name", frmDoc.rt.Font.Name
  'savestring HKEY_CLASSES_ROOT, "cEdit\data\", "leftmargin", frmdoc.rt.
End Sub

Public Sub ReadOptions(rt As CodeSense)
  Dim FntName As String
  Call rt.SetColor(cmClrComment, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "comment", 32896))
  Call rt.SetColor(cmClrBookmark, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "bookmark", -1))
  Call rt.SetColor(cmClrBookmarkBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "bookmarkbk", -1))
  Call rt.SetColor(cmClrCommentBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "commentbk", -1))
  Call rt.SetColor(cmClrHDividerLines, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "divider", -1))
  Call rt.SetColor(cmClrHighlightedLine, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "highlight", 65535))
  Call rt.SetColor(cmClrKeyword, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "keyword", 16711680))
  Call rt.SetColor(cmClrKeywordBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "keywordbk", -1))
  Call rt.SetColor(cmClrLeftMargin, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "left", 14473948))
  Call rt.SetColor(cmClrLineNumber, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "linenum", 0))
  Call rt.SetColor(cmClrLineNumberBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "linenumbk", 14473948))
  Call rt.SetColor(cmClrNumber, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "number", 0))
  Call rt.SetColor(cmClrNumberBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "numberbk", -1))
  Call rt.SetColor(cmClrOperator, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "operator", 255))
  Call rt.SetColor(cmClrOperatorBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "operatorbk", -1))
  Call rt.SetColor(cmClrScopeKeyword, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "scope", 16711680))
  Call rt.SetColor(cmClrScopeKeywordBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "scopebk", -1))
  Call rt.SetColor(cmClrString, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "string", 8388736))
  Call rt.SetColor(cmClrStringBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "stringbk", -1))
  Call rt.SetColor(cmClrTagAttributeName, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagattrib", 16711680))
  Call rt.SetColor(cmClrTagAttributeNameBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagattribbk", -1))
  Call rt.SetColor(cmClrTagElementName, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagele", 128))
  Call rt.SetColor(cmClrTagElementNameBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagelebk", -1))
  Call rt.SetColor(cmClrTagEntity, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagent", 255))
  Call rt.SetColor(cmClrTagEntityBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagentbk", -1))
  Call rt.SetColor(cmClrTagText, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagtxt", 0))
  Call rt.SetColor(cmClrTagTextBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "tagtxtbk", -1))
  Call rt.SetColor(cmClrText, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "text", 0))
  Call rt.SetColor(cmClrTextBk, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "textbk", -1))
  Call rt.SetColor(cmClrVDividerLines, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "vdivider", 8421504))
  Call rt.SetColor(cmClrWindow, GetString(HKEY_CLASSES_ROOT, "cEdit\colors\", "window", -1))
  rt.SelBounds = GetString(HKEY_CLASSES_ROOT, "cEdit\options\", "selbounds", True)
  rt.DisplayLeftMargin = GetString(HKEY_CLASSES_ROOT, "cEdit\options\", "leftmargin", False)
  rt.LineNumbering = GetString(HKEY_CLASSES_ROOT, "cEdit\data\", "numbering", True)
  rt.LineToolTips = GetString(HKEY_CLASSES_ROOT, "cEdit\options\", "lttips", True)
  rt.LineNumberStyle = GetString(HKEY_CLASSES_ROOT, "cEdit\data\", "numberingstyle", 1)
  rt.LineNumberStart = GetString(HKEY_CLASSES_ROOT, "cEdit\data\", "numberingstart", 1)
  rt.Font.Bold = GetString(HKEY_CLASSES_ROOT, "cEdit\font\", "bold", 0)
  rt.Font.Italic = GetString(HKEY_CLASSES_ROOT, "cEdit\font\", "italic", 0)
  rt.Font.Size = GetString(HKEY_CLASSES_ROOT, "cEdit\font\", "size", 8)
  rt.Font.Strikethrough = GetString(HKEY_CLASSES_ROOT, "cEdit\font\", "strike", 0)
  rt.Font.Underline = GetString(HKEY_CLASSES_ROOT, "cEdit\font\", "under", 0)
  FntName = GetString(HKEY_CLASSES_ROOT, "cEdit\font\", "name", 0)
  If FntName <> "0" Then rt.Font.Name = FntName
End Sub

Public Sub SaveFormData(frm As Form)
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Width", frm.Width
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Height", frm.Height
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Top", frm.Top
  SaveString HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Left", frm.Left
End Sub

Public Sub LoadFormData(frm As Form)
  frm.Width = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Width", frm.Width)
  frm.Height = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Height", frm.Height)
  frm.Top = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Top", frm.Top)
  frm.Left = GetString(HKEY_CURRENT_USER, "Software\cEdit\Forms\" + frm.Name, "Left", frm.Left)
End Sub

'Window Data

Public Sub WriteData()
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "windowstate", frmMain.WindowState
  frmMain.WindowState = vbNormal
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "left", frmMain.Left
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "top", frmMain.Top
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "width", frmMain.Width
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "height", frmMain.Height
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "toolbar", frmMain.tBar.Visible
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "mactoolbar", frmMain.tbMacro.Visible
  SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "statusbar", frmMain.stBar.Visible
  'SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "quicknav", frmMain.Picture2.Visible
  'SaveString HKEY_CLASSES_ROOT, "cEdit\window\", "quicknavwidth", frmMain.Picture2.Width
End Sub

Public Sub ReadData()
  Dim m As Boolean
  frmMain.Left = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "left", 1980)
  frmMain.Top = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "top", 1980)
  frmMain.Width = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "width", 10080)
  frmMain.Height = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "height", 5640)
  frmMain.WindowState = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "windowstate", 0)
  'frmMain.Picture2.Width = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "quicknavwidth", 3005)
'  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "quicknav", True)
'  frmMain.quicknav.Checked = m
  'frmMain.Picture2.Visible = m
  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "toolbar", True)
  frmMain.tBar.Visible = m
  frmMain.toolbar.Checked = m
  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "mactoolbar", True)
  frmMain.tbMacro.Visible = m
  frmMain.mnuMacBar.Checked = m
  m = GetString(HKEY_CLASSES_ROOT, "cEdit\window\", "statusbar", True)
  frmMain.stBar.Visible = m
  frmMain.statusbar2.Checked = m
End Sub

Public Sub WriteInput()
  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "whitespace", frmMain.whitespace.Checked
  SaveString HKEY_CLASSES_ROOT, "cEdit\options\", "hlline", frmMain.hlline.Checked
End Sub

Public Sub ReadInput()
  WhiteSpaced = GetString(HKEY_CLASSES_ROOT, "cEdit\options\", "whitespace", False)
  frmMain.whitespace.Checked = WhiteSpaced
  frmDoc.rt.DisplayWhitespace = WhiteSpaced
  HighLight = GetString(HKEY_CLASSES_ROOT, "cEdit\options\", "hlline", False)
  frmMain.hlline.Checked = HighLight
End Sub
