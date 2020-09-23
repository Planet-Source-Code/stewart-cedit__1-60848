Attribute VB_Name = "modApp"
Option Explicit
'+------------------------------------------------------------+
'| Application module. Contains functions specificly involved |
'| with file creation and editing, opening, etc.              |
'| Coder: Ackbar                                              |
'| You may use this as long as you leave this section in.     |
'+------------------------------------------------------------+
Public Const PROJECT_EXTENSIONS = "*.vbp;*.vbg;*.vbproj;*.cep"

Public closeall As Boolean 'This is gonna be used to check if the
                           'the the closealldoc function was called
                           'so if it is we don't update the enabled/disabled
                           'state (flickering issues)
Public Type lang
  Name As String
  Keywords As String
  Operators As String
  SingleLineComment As String
  MultiLineComment1 As String
  MultiLineComment2 As String
  ScopeKeywords1 As String
  ScopeKeywords2 As String
  EscapeChar As String
  StringDelims As String
  Style As Integer
  TagAttributeNames As String
  TagElementNames As String
  TagEntities As String
  TerminatorCharacter As String
  FileAssociation As String
  CaseSensative As Integer
End Type

Public Type Recent
  Recent1 As String
  Recent2 As String
  Recent3 As String
  Recent4 As String
  Recent5 As String
  Recent6 As String
End Type

Public Type FTP
  Name As String
  URL As String
  UserName As String
  Password As String
  Anonymous As Integer
  PortNum As Integer
  LastDir As String
End Type

Public HighLight As Boolean 'This stores whether or not to highlight the
                            'selected line :)
Public Type FormState
    Deleted As Boolean
    dirty As Boolean
    Type As Integer
    Color As Long
End Type
Public FState() As FormState
Public fIndex As Integer
Public Document() As New frmDoc
Public dnum As Integer
Public Recnt As Recent
Public WhiteSpaced As Boolean

                            
Public Sub doNew(str As String)
  On Error Resume Next
  fIndex = FindFreeIndex()
  If fIndex = 0 Then
    fIndex = 1
    ReDim Document(1 To 1)
    ReDim FState(1 To 1)
  End If
  Document(fIndex).Changed = False
  Document(fIndex).Tag = fIndex
  Document(fIndex).Caption = "Untitled " & Document(fIndex).Tag
  Document(fIndex).Move 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
  Document(fIndex).WindowState = vbMaximized
  Document(fIndex).rt.Language = str
  Document(fIndex).Visible = True
End Sub

Function FindFreeIndex() As Integer
    On Error GoTo errHandler
    Dim I As Integer
    Dim ArrayCount As Integer
    ArrayCount = UBound(Document)
    For I = 1 To ArrayCount
        If FState(I).Deleted Then
            FindFreeIndex = I
            FState(I).Deleted = False
            Exit Function
        End If
    Next
    ReDim Preserve Document(1 To ArrayCount + 1)
    ReDim Preserve Document(1 To ArrayCount + 1)
    ReDim Preserve FState(1 To ArrayCount + 1)
    FindFreeIndex = UBound(Document)
    Exit Function
errHandler:
    FindFreeIndex = 0
End Function

Public Function StripPath(t As String) As String
Dim X As Integer
Dim ct As Integer
    StripPath = t
    X = InStr(t, "\")
    Do While X
        ct = X
        X = InStr(ct + 1, t, "\")
    Loop
    If ct > 0 Then StripPath = Mid(t, ct + 1)
End Function


Public Sub doSave()
  On Error Resume Next
  If Document(dnum).IsFile = False Then
    doSaveAs
    Document(dnum).DoAct
  Else
    Document(dnum).rt.SaveFile Document(dnum).Caption, False
    Document(dnum).Changed = False
    Document(dnum).DoAct
  End If
End Sub


Public Sub doSaveAs()
  On Error GoTo errHandler
  Dim msg As VbMsgBoxResult
  frmMain.cd.CancelError = True
  frmMain.cd.filename = ""
  frmMain.cd.DialogTitle = "Save document... " & Document(dnum).Caption
  frmMain.cd.Filter = AllSupport & FilterB  '"All Files|*.*|Text Files|*.txt|Html Files|*.html;*.htm|Style Sheets|*.css|Java Scripting|*.js|C Files|*.c|C++ Files|*.cpp|C/C++ Header Files|*.h|Perl Files|*.pl|CGI/Perl Files|*.cgi|XML Files|*.xml|Pascal Files|*.pas|Basic Module Files|*.bas|Basic Form Files|*.frm|Basic Project Files|*.vbp|Basic Class Modules|*.cls"
  frmMain.cd.ShowSave
  'If frmMain.cd.filename = "" Then Exit Function
  If FileExists(frmMain.cd.filename) = True Then
     msg = MsgBox(frmMain.cd.filename & " Already exists." & Chr(10) & "Do you want to replace it?", vbYesNo + vbQuestion, "Overwrite")
    If msg = vbYes Then
      'continue
    ElseIf msg = vbNo Then
      doSaveAs
    End If
  End If
  Document(dnum).IsFile = True
  Document(dnum).rt.SaveFile frmMain.cd.filename, True
  Document(dnum).Caption = frmMain.cd.filename
  Document(dnum).rt.Language = SetSyntax(frmMain.cd.filename)
  Document(fIndex).SetLangWords
  Document(dnum).Changed = False
  AddRecent frmMain.cd.filename
errHandler:
  If Err.Number = 32755 Or Err.Number = 0 Then
    Exit Sub
  Else
    MsgBox "Error: " & Err.Number & Chr(10) & Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
  End If
  Exit Sub
End Sub
Public Sub RegisterAll()
  Dim s As String, path As String
  path = CheckPath(App.path & "\lang\")
  s = Dir(path, vbNormal)
  Do While LenB(s) <> 0
    If Right$(LCase$(s), 3) = "lng" Then
      RegLang s
    End If
    s = Dir
  Loop
End Sub

Public Sub RegLang(fle As String)
  Dim fFile As Integer
  Dim X As Integer, UA() As String
  Dim LangD As CodeSenseCtl.Language
  Set LangD = New CodeSenseCtl.Language
  Dim globals As CodeSenseCtl.globals
  Set globals = New CodeSenseCtl.globals
  Dim Langf As lang
  fFile = FreeFile()
  Open App.path & "\lang\" & fle For Binary Access Read As #fFile
    Get #fFile, , Langf
  Close #1
  With LangD
    .Keywords = Langf.Keywords
    .Operators = Langf.Operators
    .SingleLineComments = Langf.SingleLineComment
    .MultiLineComments1 = Langf.MultiLineComment1
    .MultiLineComments2 = Langf.MultiLineComment2
    .ScopeKeywords1 = Langf.ScopeKeywords1
    .ScopeKeywords2 = Langf.ScopeKeywords2
    .EscapeChar = Langf.EscapeChar
    .Style = Langf.Style
    .StringDelims = Langf.StringDelims
    .TagAttributeNames = Langf.TagAttributeNames
    .TagElementNames = Langf.TagElementNames
    .TagEntities = Langf.TagEntities
    .TerminatorChar = Langf.TerminatorCharacter
    .CaseSensitive = Langf.CaseSensative
    AllSupport = AllSupport & ";" & Langf.FileAssociation
    FilterB = FilterB & "|" & Langf.Name & " Files|*" & Langf.FileAssociation
    UA() = Split(Langf.FileAssociation, " ")
    For X = 0 To UBound(UA)
      ClrString = ClrString & " " & UA(X) & ":" & Langf.Name
    Next
    Erase UA()
    'frmMain.lang(frmMain.lang.Count + 1).Caption = Langf.Name
    Call globals.RegisterLanguage(Langf.Name, LangD)
    Langs = Langs & Langf.Name & Chr(10)
    'frmmain.lang
  End With
End Sub

'+--------------------------------------------------------------------+

'+--------------------------------------------------------------------+
'| CheckPath is a simple function that will insert the needed \ on the|
'| end of a path if it's not there. Thats all :)                      |
'+--------------------------------------------------------------------+
Public Function CheckPath(ByVal path As String) As String
If Right$(path, 1) <> "\" Then
  CheckPath = path & "\"
Else
  CheckPath = path
End If
End Function

'+--------------------------------------------------------------------+
'| This is a revised version of the setsyntax code. It's far improved |
'| in that it now easily supports external languages. Also it is far  |
'| less code :)                                                       |
'+--------------------------------------------------------------------+
Public Function SetSyntax(file As String) As String
  Dim Extension As String, UA() As String, ClrExt As String, X As Long
  If InStr(1, ClrString, " ") <> 0 Then
    UA = Split(ClrString, " ")
  End If
  Extension = LCase$(Mid$(file, InStrRev(file, ".") + 1, Len(file) - InStrRev(file, ".")))
  For X = 0 To UBound(UA)
    ClrExt = LCase$(Mid$(UA(X), 1, InStr(1, UA(X), ":") - 1))
    If LCase$(ClrExt) = LCase$(Extension) Then Exit For
  Next
  If LCase$(ClrExt) <> LCase$(Extension) Then Exit Function
  SetSyntax = LCase$(Mid$(UA(X), InStrRev(UA(X), ":") + 1, Len(UA(X)) - InStrRev(file, ":")))
  If LenB(SetSyntax) = 0 Then SetSyntax = "Text"
  Erase UA
End Function
Public Sub DoOpen(path As String)
  On Error Resume Next
  Dim X As Long
  
  If Dir(path) = "" Then
    If MsgBox("The file: " & path & Chr(10) & "does not exist. Do you wish to create it?", vbYesNo + vbQuestion, "Create File") = vbNo Then Exit Sub
  End If
  
  If IsProject(path) Then
    LoadProject path
    'LoadVBProject path
    Exit Sub
  End If
  
  'If there's 0 open docs no need to do the loop to verify the file's not open.
  'but if there are files we wanna make sure that none are the one were about
  'to open (whats the point of opening the same file twice) and if this is the
  'case then we will just setfocus :)
  If fIndex > 0 Then
    'First lets check and find out of this file is open or not
    For X = 1 To UBound(Document)
      If FState(X).Deleted = False Then
        If Document(X).IsFile = True And Document(X).filename = path Then
          Document(X).SetFocus
          Exit Sub
        End If
      End If
    Next
  End If
  fIndex = FindFreeIndex()
  If fIndex = 0 Then
    fIndex = 1
    ReDim Document(1 To 1)
    ReDim FState(1 To 1)
  End If
  Document(fIndex).Changed = False
  Document(fIndex).Tag = fIndex
  Document(fIndex).Caption = path
  Document(fIndex).IsFile = True
  Document(fIndex).filename = path
  Document(fIndex).rt.OpenFile path
  Document(fIndex).rt.Language = SetSyntax(path)
  Document(fIndex).cboLanguage.Text = Document(fIndex).rt.Language
  Document(fIndex).SetLangWords
  Document(fIndex).Show
End Sub


Public Function FileExists(FullFileName As String) As Boolean
    On Error Resume Next
    If Dir(FullFileName) = "" Then
      FileExists = False
    Else
      FileExists = True
    End If
End Function




Public Sub openftp(str As String, path As String, ftpDir As String, FTPAccount As String)
  On Error Resume Next
  fIndex = FindFreeIndex()
  Document(fIndex).Changed = False
  Document(fIndex).Tag = fIndex
  Document(fIndex).Caption = path
  Document(fIndex).filename = path
  Document(fIndex).rt.Text = str
  Document(fIndex).rt.Language = SetSyntax(path)
  Document(fIndex).SetLangWords
  Document(fIndex).FTP = True
  Document(fIndex).FTPAccount = FTPAccount
  Document(fIndex).ftpDir = ftpDir
  Document(fIndex).Show
End Sub

Public Sub InsertString(rt As CodeSense, str As String)
      Dim r As CodeSenseCtl.range
      Set r = New CodeSenseCtl.range
      rt.SelText = str
      Set r = rt.GetSel(False)
      rt.SetCaretPos r.EndLineNo, r.StartColNo + Len(str)
      rt.SetFocus
End Sub

Public Sub AddRecent(str As String)
  Dim FreeFileNum As Integer
  With Recnt
    .Recent6 = .Recent5
    .Recent5 = .Recent4
    .Recent4 = .Recent3
    .Recent3 = .Recent2
    .Recent2 = .Recent1
    .Recent1 = str
  End With
  FreeFileNum = FreeFile()
  Open App.path & "\temp\recent.rct" For Binary Access Write As #FreeFileNum
    Put #FreeFileNum, , Recnt
  Close #FreeFileNum
  With frmMain
    If Recnt.Recent1 <> "" Then
      .mnuRec(0).Caption = Recnt.Recent1
      .mnuRec(0).Visible = True
    End If
    If Recnt.Recent2 <> "" Then
      .mnuRec(1).Caption = Recnt.Recent2
      .mnuRec(1).Visible = True
    End If
    If Recnt.Recent3 <> "" Then
      .mnuRec(2).Caption = Recnt.Recent3
      .mnuRec(2).Visible = True
    End If
    If Recnt.Recent4 <> "" Then
      .mnuRec(3).Caption = Recnt.Recent4
      .mnuRec(3).Visible = True
    End If
    If Recnt.Recent5 <> "" Then
      .mnuRec(4).Caption = Recnt.Recent5
      .mnuRec(4).Visible = True
    End If
    If Recnt.Recent6 <> "" Then
      .mnuRec(5).Caption = Recnt.Recent6
      .mnuRec(5).Visible = True
    End If
  End With

End Sub

Public Sub LoadRecent()
  Dim FreeFileNum As Integer
  FreeFileNum = FreeFile()
  Open App.path & "\temp\recent.rct" For Binary Access Read As #FreeFileNum
    Get #FreeFileNum, , Recnt
  Close #FreeFileNum
  With frmMain
    If Recnt.Recent1 <> "" Then
      .mnuRec(0).Caption = Recnt.Recent1
      .mnuRec(0).Visible = True
    End If
    If Recnt.Recent2 <> "" Then
      .mnuRec(1).Caption = Recnt.Recent2
      .mnuRec(1).Visible = True
    End If
    If Recnt.Recent3 <> "" Then
      .mnuRec(2).Caption = Recnt.Recent3
      .mnuRec(2).Visible = True
    End If
    If Recnt.Recent4 <> "" Then
      .mnuRec(3).Caption = Recnt.Recent4
      .mnuRec(3).Visible = True
    End If
    If Recnt.Recent5 <> "" Then
      .mnuRec(4).Caption = Recnt.Recent5
      .mnuRec(4).Visible = True
    End If
    If Recnt.Recent6 <> "" Then
      .mnuRec(5).Caption = Recnt.Recent6
      .mnuRec(5).Visible = True
    End If
  End With
End Sub

Public Function GetExtension(sFileName As String) As String
    Dim lPos As Long
    lPos = InStrRev(sFileName, ".")
    If lPos = 0 Then
        GetExtension = " "
    Else
        GetExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
    End If
End Function


Public Function IsProject(sFile As String) As Boolean
Dim sExtension As String
    sExtension = GetExtension(sFile)
    If InStr(1, PROJECT_EXTENSIONS & ";", "." & sExtension & ";") Then IsProject = True
End Function

Private Sub LoadProject(path As String)
  'determine what type of file we are dealing with.
  Dim ext As String
  ext = GetExtension(path)
  Select Case ext
    Case "vbp"
      LoadVBProject path
    Case "vbg"
      LoadVBGroup path
  End Select
End Sub

Public Function ShowSite(URL As String)
  If frmBrowse.Visible = False Then frmBrowse.Visible = True
  frmBrowse.SetFocus
  frmBrowse.www.Navigate URL
End Function
