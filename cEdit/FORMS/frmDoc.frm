VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDoc 
   Caption         =   "Untitled"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4275
   Icon            =   "frmDoc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   4275
   WindowState     =   2  'Maximized
   Begin VB.Frame fmLang 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   -70
      Width           =   4335
      Begin VB.ComboBox cboLanguage 
         Height          =   315
         ItemData        =   "frmDoc.frx":1042
         Left            =   480
         List            =   "frmDoc.frx":1044
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   200
         Width           =   2025
      End
      Begin VB.ComboBox cboProcedures 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmDoc.frx":1046
         Left            =   2880
         List            =   "frmDoc.frx":1048
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   200
         Width           =   1995
      End
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   0
         Left            =   2520
         Picture         =   "frmDoc.frx":104A
         ToolTipText     =   "Jump to..."
         Top             =   195
         Width           =   240
      End
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "frmDoc.frx":1194
         ToolTipText     =   "Language"
         Top             =   200
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imgCode 
      Left            =   3360
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   11
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoc.frx":12DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoc.frx":1540
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoc.frx":17A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CodeSenseCtl.CodeSense rt 
      Height          =   2055
      Left            =   840
      OleObjectBlob   =   "frmDoc.frx":1A04
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TabNum As Long
Public Changed As Boolean
Public FileName As String
Public r As CodeSenseCtl.range
Public FTP As Boolean
Public IsFile As Boolean
Public FTPAccount As String
Public ftpDir As String
Private LastLine As New Collection
Private LineIndex As Long
Dim Keywords() As String, Elements() As String, Attributes() As String

Private Sub cboLanguage_Click()
  On Error Resume Next
  rt.Language = cboLanguage.Text
  LangKeywords cboLanguage.Text
  SetLangWords
End Sub

Private Sub cboProcedures_Click()
  InsertString rt, cboProcedures.Text
End Sub

Private Sub Form_Activate()
  dnum = Me.Tag
    
  frmMain.tb.Tabs("key" & Me.Tag).Caption = StripPath(Me.Caption)
  frmMain.tb.Tabs("key" & Me.Tag).Selected = True
  frmMain.tb.Tabs("key" & Me.Tag).ToolTipText = Me.Caption
  If Changed = False Then
    frmMain.tb.Tabs("key" & Me.Tag).Image = 1
  Else
    frmMain.tb.Tabs("key" & Me.Tag).Image = 2
  End If
  EnableMac
  EnableMenu
  OpenAble
  ShouldEnable
  If rt.Language = "" Then
    frmMain.stBar.Panels(3).Text = "Text"
  Else
    frmMain.stBar.Panels(3).Text = rt.Language
  End If
End Sub

'+-----------------------------------------------------------------------+
'| Build language list. I descided for some reason to store the languages|
'| on the document form itself this time. I think it will improve        |
'| performance a bit. On top of that I think it will make it easier to   |
'| use. You will also notice a language keyword list next to the language|
'| list. It is updated to display all keywords of the language as a      |
'| quick refrence. Highlighting a keyword will insert it.                |
'+-----------------------------------------------------------------------+
Private Sub AddLang()
  Dim UA() As String, LngCnt As Long
  cboLanguage.Clear
  cboProcedures.Clear
  cboLanguage.AddItem "Text"
  cboLanguage.AddItem "C/C++"
  cboLanguage.AddItem "Basic"
  cboLanguage.AddItem "Java"
  cboLanguage.AddItem "Pascal"
  cboLanguage.AddItem "SQL"
  cboLanguage.AddItem "HTML"
  cboLanguage.AddItem "XML"
  cboLanguage.Text = "Text"
  UA = Split(Langs, Chr$(10))
  For LngCnt = 0 To UBound(UA) - 1: cboLanguage.AddItem UA(LngCnt): Next
  Erase UA
End Sub

Private Sub Form_Load()
  On Error Resume Next
  AddLang
  frmMain.tb.Tabs.Add fIndex, ("key" & Format(fIndex)), Me.Caption ' "me.tag: " & Me.Tag    'Me.Caption  AddLang
  ReadOptions rt
  ReadInput
  Clear
  LastLine.Add 0
  LineIndex = 1
  If WhiteSpaced = True Then
    rt.DisplayWhitespace = True
  Else
    rt.DisplayWhitespace = False
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim msgRes As VbMsgBoxResult
  If rt.Modified = True Then
    msgRes = MsgBox("Document: " & Me.Caption & Chr(10) & "Do you wish to save?", vbYesNoCancel + vbQuestion, "Save")
    If msgRes = vbYes Then
      doSave
    ElseIf msgRes = vbNo Then
      'do nothing

    ElseIf msgRes = vbCancel Then
      'Cancel
      Cancel = 1
      StopClose = True
      rt.SetFocus
    End If

  End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  fmLang.Move 0, -100, Me.ScaleWidth
  rt.Move 0, fmLang.Top + fmLang.Height, Me.ScaleWidth, Me.ScaleHeight - fmLang.Height - fmLang.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  'Me.Visible = False
  DisableMac
  DisableMenu
  CloseAble
  FState(Me.Tag).Deleted = True
  frmMain.tb.Tabs.Remove ("key" & Format(Me.Tag))
  'Unload Me
  dnum = 0
End Sub




Private Sub rt_Change(ByVal Control As CodeSenseCtl.ICodeSense)
  Changed = rt.Modified
  If Changed = True Then
    frmMain.tb.Tabs("key" & Me.Tag).Image = 2
  Else
    frmMain.tb.Tabs("key" & Me.Tag).Image = 1
  End If
End Sub

Private Function rt_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
'  On Error Resume Next
  Dim i As Integer
  ListCtrl.hImageList = imgCode.hImageList
  ListCtrl.EnableHotTracking True
  For i = 0 To UBound(Keywords) - 1
    ListCtrl.AddItem Keywords(i), 0
  Next
  For i = 0 To UBound(Elements) - 1
    ListCtrl.AddItem Elements(i), 1
  Next
  For i = 0 To UBound(Attributes) - 1
    ListCtrl.AddItem Attributes(i), 2
  Next
  
  rt_CodeList = True
End Function

Private Function rt_CodeListSelChange(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As String
  on error resume next
  rt_CodeListSelChange = ListCtrl.GetItemText(lItem)
    Dim R As CodeSenseCtl.IRange
    Dim colorh As Long
   
    Set R = Control.GetSel(True)
       
    colorh = Control.GetColor(cmClrHighlightedLine)
    Call Control.SetColor(cmClrHighlightedLine, 
Control.GetColor(cmClrWindow))
    Control.HighlightedLine = R.StartLineNo
    DoEvents
    Call Control.SetColor(cmClrHighlightedLine, colorh)
  
    Set R = Nothing
   
    Err = 0
             
End Function

Private Function rt_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    Dim strItem As String
    Dim range As New CodeSenseCtl.range

    ' Determine which item was selected in the list
    strItem = ListCtrl.GetItemText(ListCtrl.SelectedItem)

    ' Replace current selection
    rt.ReplaceSel (strItem)

    ' Get new selection
    Set range = rt.GetSel(True)

    ' Update range to end of newly inserted text
    range.StartColNo = range.StartColNo + Len(strItem)
    range.EndColNo = range.StartColNo
    range.EndLineNo = range.StartLineNo

    ' Move cursor
    rt.SetSel range, True

    ' Clear any text left in the status bar

    ' Don't prevent list view control from being hidden
    rt_CodeListSelMade = False

End Function

Private Function rt_CodeListSelWord(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As Boolean
  rt_CodeListSelWord = True
End Function

Private Sub rt_GotFocus()
  Dim X As Integer
  SetLang
  For X = 1 To frmMain.tb.Tabs.Count
    If frmMain.tb.Tabs(X).Tag = Me.TabNum Then
      frmMain.tb.Tabs(X).Selected = True
      Exit For
    End If
  Next
End Sub



Private Function rt_MouseDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
  If Button = 2 Then PopupMenu frmMain.edit
End Function


Private Sub rt_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
  On Error Resume Next
  Set r = rt.GetSel(True)

  frmMain.stBar.Panels(2).Text = "Ln " & r.EndLineNo + 1 & ", Col. " & r.EndColNo & ", Lines " & rt.LineCount
  If HighLight = True Then
    rt.HighlightedLine = r.EndLineNo
  Else
    rt.HighlightedLine = -1
  End If
  If LastLine.Count <> 0 Then
    If LastLine(LineIndex) <> r.EndLineNo Then
      LastLine.Add r.EndLineNo
      LineIndex = LastLine.Count
    End If
  Else
    LastLine.Add r.EndLineNo
    LineIndex = LastLine.Count
  End If
  If LineIndex < LastLine.Count Then
    frmMain.tBar.Buttons(39).Enabled = True
  Else
    frmMain.tBar.Buttons(39).Enabled = False
  End If
  
  If LineIndex > 1 Then
    frmMain.tBar.Buttons(38).Enabled = True
  Else
    frmMain.tBar.Buttons(38).Enabled = False
  End If
  ShouldEnable
End Sub

Private Sub SetLang()
  setChk
End Sub

Private Sub setChk()
  Dim X As Integer
End Sub

Public Sub NextLine()
  On Error Resume Next
  If LineIndex + 1 > LastLine.Count Then Exit Sub 'this won't work we are at the end of lines
  LineIndex = LineIndex + 1
  'rt.ExecuteCmd cmCmdGotoLine, LastLine(LineIndex)
  rt.SetCaretPos LastLine(LineIndex), 1
  If LineIndex < LastLine.Count Then
    frmMain.tbProgramming.Buttons(13).Enabled = True
  Else
    frmMain.tbProgramming.Buttons(13).Enabled = True
  End If
  
  If LineIndex > 1 Then
    frmMain.tbProgramming.Buttons(12).Enabled = True
  Else
    frmMain.tbProgramming.Buttons(12).Enabled = False
  End If
End Sub

Public Sub PrevLine()
  On Error Resume Next
  If LineIndex - 1 < 1 Then Exit Sub 'this won't work we are at the end of lines
  LineIndex = LineIndex - 1
  rt.SetCaretPos LastLine(LineIndex), 1
  If LineIndex < LastLine.Count Then
    frmMain.tbProgramming.Buttons(13).Enabled = True
  Else
    frmMain.tbProgramming.Buttons(13).Enabled = False
  End If
  
  If LineIndex > 1 Then
    frmMain.tbProgramming.Buttons(12).Enabled = True
  Else
    frmMain.tbProgramming.Buttons(12).Enabled = False
  End If
End Sub

Public Sub CommentBlock()
  If rt.Language = "" Then Exit Sub
  Dim i As Integer
  Dim p As CodeSenseCtl.Language
  Dim s As CodeSenseCtl.globals
  Set s = New CodeSenseCtl.globals
  Dim UA() As String
  Dim cmstring As String
  Dim NewStr As String
  Set p = s.GetLanguageDef(rt.Language)
  cmstring = p.SingleLineComments
  Set p = Nothing
  Set s = Nothing
  NewStr = ""
  UA = Split(rt.SelText, Chr$(10))
  For i = 0 To UBound(UA)
    NewStr = NewStr & cmstring & UA(i)
  Next
  Erase UA
  InsertString rt, NewStr
  NewStr = ""
  cmstring = ""
End Sub

Public Sub UncommentBlock()
  If rt.Language = "" Then Exit Sub
  Dim p As CodeSenseCtl.Language
  Dim s As CodeSenseCtl.globals, cmstring As String
  Set s = New CodeSenseCtl.globals
  Set p = s.GetLanguageDef(rt.Language)
  cmstring = p.SingleLineComments
  Set p = Nothing
  Set s = Nothing
  InsertString rt, Replace(rt.SelText, cmstring, "")
End Sub

Private Sub Clear()
  Dim X As Integer
  For X = LastLine.Count To 1 Step -1
    LastLine.Remove (X)
  Next
End Sub

Public Sub DoAct()
  frmMain.tb.Tabs("key" & Me.Tag).Caption = StripPath(Me.Caption)
  frmMain.tb.Tabs("key" & Me.Tag).Selected = True
  frmMain.tb.Tabs("key" & Me.Tag).ToolTipText = Me.Caption
  If Changed = False Then
    frmMain.tb.Tabs("key" & Me.Tag).Image = 1
  Else
    frmMain.tb.Tabs("key" & Me.Tag).Image = 2
  End If
End Sub

Private Sub EnableMac()
  Dim X As Integer
  If (frmMain.tbMacro.Buttons("mac1").Enabled) = True Then Exit Sub 'They are already enabled
  For X = 1 To 10
    frmMain.tbMacro.Buttons("mac" & X).Enabled = True
  Next
  frmMain.tbMacro.Buttons("cmac").Enabled = True
End Sub

Private Sub DisableMac()
  Dim X As Integer
  If frmMain.tbMacro.Buttons("mac1").Enabled = False Then Exit Sub
  If UBound(Document()) > 0 Then Exit Sub
  If closeall = True Then Exit Sub
  For X = 1 To 10
    frmMain.tbMacro.Buttons("mac" & X).Enabled = False
  Next
  frmMain.tbMacro.Buttons("cmac").Enabled = False
End Sub

Private Sub DisableMenu()
  Dim X As Integer
  If frmMain.close.Enabled = False Then Exit Sub  'Just make sure it's not already done to save the effort
  If UBound(Document()) > 0 Then Exit Sub
  frmMain.close.Enabled = False
  frmMain.save.Enabled = False
  frmMain.saveas.Enabled = False
  frmMain.saveall.Enabled = False
  frmMain.saveto.Enabled = False
  frmMain.prints.Enabled = False
  frmMain.printsetup.Enabled = False
  frmMain.properties.Enabled = False
  frmMain.undo.Enabled = False
  frmMain.redo.Enabled = False
  frmMain.cut.Enabled = False
  frmMain.copy.Enabled = False
  frmMain.paste.Enabled = False
  frmMain.delete.Enabled = False
  frmMain.mnuComment.Enabled = False
  frmMain.mnuUncomment.Enabled = False
  frmMain.selectall.Enabled = False
  frmMain.selectline.Enabled = False
  frmMain.datetime.Enabled = False
  frmMain.find.Enabled = False
  frmMain.findnext.Enabled = False
  frmMain.findprev.Enabled = False
  frmMain.mnuReplace.Enabled = False
  frmMain.goto.Enabled = False
  frmMain.mnuToggle.Enabled = False
  frmMain.mnuNext.Enabled = False
  frmMain.mnuPrev.Enabled = False
  frmMain.mnuClear.Enabled = False
  frmMain.mnuNLine.Enabled = False
  frmMain.mnuLPrev.Enabled = False
  frmMain.countall.Enabled = False
  frmMain.mnuCompile.Enabled = False
  frmMain.mnuBuildConfig.Enabled = False
  For X = 1 To 10
    frmMain.mac(X).Enabled = False
  Next
  frmMain.mnuSave.Enabled = False
  frmMain.mnuCreate.Enabled = False
  frmMain.tilehor.Enabled = False
  frmMain.tilever.Enabled = False
  frmMain.arrangeicons.Enabled = False
  frmMain.cascade.Enabled = False
  frmMain.closeall.Enabled = False
  frmMain.inbrowser.Enabled = False
  frmMain.wnlist.Enabled = False
End Sub

Private Sub EnableMenu()
  Dim X As Integer
  If frmMain.close.Enabled = True Then Exit Sub
  If closeall = True Then Exit Sub
  frmMain.close.Enabled = True
  frmMain.save.Enabled = True
  frmMain.saveas.Enabled = True
  frmMain.saveall.Enabled = True
  frmMain.saveto.Enabled = True
  frmMain.prints.Enabled = True
  frmMain.printsetup.Enabled = True
  frmMain.properties.Enabled = True
  frmMain.undo.Enabled = True
  frmMain.redo.Enabled = True
  frmMain.cut.Enabled = True
  frmMain.copy.Enabled = True
  frmMain.paste.Enabled = True
  frmMain.delete.Enabled = True
  frmMain.mnuComment.Enabled = True
  frmMain.mnuUncomment.Enabled = True
  frmMain.selectall.Enabled = True
  frmMain.selectline.Enabled = True
  frmMain.datetime.Enabled = True
  frmMain.find.Enabled = True
  frmMain.findnext.Enabled = True
  frmMain.findprev.Enabled = True
  frmMain.mnuReplace.Enabled = True
  frmMain.goto.Enabled = True
  frmMain.mnuToggle.Enabled = True
  frmMain.mnuNext.Enabled = True
  frmMain.mnuPrev.Enabled = True
  frmMain.mnuClear.Enabled = True
  frmMain.mnuNLine.Enabled = True
  frmMain.mnuLPrev.Enabled = True
  frmMain.countall.Enabled = True
  frmMain.mnuCompile.Enabled = True
  frmMain.mnuBuildConfig.Enabled = True
  For X = 1 To 10
    frmMain.mac(X).Enabled = True
  Next
  frmMain.mnuSave.Enabled = True
  frmMain.mnuCreate.Enabled = True
  frmMain.tilehor.Enabled = True
  frmMain.tilever.Enabled = True
  frmMain.arrangeicons.Enabled = True
  frmMain.cascade.Enabled = True
  frmMain.closeall.Enabled = True
  frmMain.inbrowser.Enabled = True
  frmMain.wnlist.Enabled = True
End Sub

Private Sub CloseAble()
  ' This is a simple function which will
  ' disable all the unnecisary toolbar buttons
  ' on the main toolbar when the form unloads.
  
  ' First off lets check a button which is always
  ' enabled if a doc is open but never enabled if
  ' it's not and check (just to avoid doing this if
  ' their already disabled
  If (frmMain.tBar.Buttons("close").Enabled = False) Then Exit Sub 'Exit sub if it's already disabled
  If UBound(Document) >= 1 Then Exit Sub
  With frmMain.tBar
    .Buttons("close").Enabled = False
    .Buttons("save").Enabled = False
    .Buttons("saveas").Enabled = False
    .Buttons("saveall").Enabled = False
    .Buttons("reload").Enabled = False
    .Buttons("print").Enabled = False
    .Buttons("undo").Enabled = False
    .Buttons("redo").Enabled = False
    .Buttons("cut").Enabled = False
    .Buttons("copy").Enabled = False
    .Buttons("paste").Enabled = False
    .Buttons("delete").Enabled = False
    .Buttons("find").Enabled = False
    .Buttons("findnext").Enabled = False
    .Buttons("findprev").Enabled = False
    .Buttons("tilehor").Enabled = False
    .Buttons("tilever").Enabled = False
    .Buttons("cascade").Enabled = False
    .Buttons("tabl").Enabled = False
    .Buttons("tabr").Enabled = False
    .Buttons("cblock").Enabled = False
    .Buttons("ublock").Enabled = False
    .Buttons("tbmark").Enabled = False
    .Buttons("nbmark").Enabled = False
    .Buttons("pbmark").Enabled = False
    .Buttons("cbmark").Enabled = False
    .Buttons("nline").Enabled = False
    .Buttons("pline").Enabled = False
    .Buttons("nline").Enabled = False
    .Buttons("ctag").Enabled = False
  End With
End Sub

Private Sub OpenAble()
  ' This function will enabled all the neccisary buttons on the primary toolbar
  ' Please note not all the toolbar buttons will automaticly be enabled
  ' as some are dependant on certain factors.
  
  ' Check to verify their not already enabled cause if they are
  ' we don't need to waste the processor time on this
  If (frmMain.tBar.Buttons("close").Enabled = True) Then Exit Sub ' If it's enabled already just exit this sub
  
  With frmMain.tBar
    .Buttons("close").Enabled = True
    .Buttons("save").Enabled = True
    .Buttons("saveas").Enabled = True
    .Buttons("saveall").Enabled = True
    .Buttons("reload").Enabled = True
    .Buttons("delete").Enabled = True
    .Buttons("print").Enabled = True
    .Buttons("tilehor").Enabled = True
    .Buttons("tilever").Enabled = True
    .Buttons("cascade").Enabled = True
  End With
  With frmMain.tbSearch
    .Buttons("find").Enabled = True
    .Buttons("findnext").Enabled = True
    .Buttons("findprev").Enabled = True
  End With
  With frmMain.tbProgramming
    .Buttons("tabl").Enabled = True
    .Buttons("tabr").Enabled = True
    .Buttons("cblock").Enabled = True
    .Buttons("ublock").Enabled = True
    .Buttons("tbmark").Enabled = True
    .Buttons("nbmark").Enabled = True
    .Buttons("pbmark").Enabled = True
    .Buttons("cbmark").Enabled = True
    .Buttons("nline").Enabled = True
    .Buttons("pline").Enabled = True
    .Buttons("ctag").Enabled = True
  End With
End Sub

Private Sub ShouldEnable()
  ' This simple function will simply check if the
  ' several different things can take place and then
  ' depending on that enable or disable certain buttons
  ' on the toolbar.
  
  With frmMain.tBar
    If rt.CanUndo Then
      .Buttons("undo").Enabled = True
    Else
      .Buttons("undo").Enabled = False
    End If
    If rt.CanRedo Then
      .Buttons("redo").Enabled = True
    Else
      .Buttons("redo").Enabled = False
    End If
    If rt.CanCut Then
      .Buttons("cut").Enabled = True
    Else
      .Buttons("cut").Enabled = False
    End If
    If rt.CanCopy Then
      .Buttons("copy").Enabled = True
    Else
      .Buttons("copy").Enabled = False
    End If
    If rt.CanPaste Then
      .Buttons("paste").Enabled = True
    Else
      .Buttons("paste").Enabled = False
    End If
  End With
End Sub


Public Sub SetLangWords()
  Dim lng As CodeSenseCtl.Language, X As Long
  Dim glb As CodeSenseCtl.globals
  Set lng = New CodeSenseCtl.Language
  Set glb = New CodeSenseCtl.globals
  If rt.Language <> "" Then
    Set lng = glb.GetLanguageDef(rt.Language)
  End If
  Erase Keywords
  Erase Attributes
  Erase Elements
  Keywords = Split(lng.Keywords, Chr$(10))
  Attributes = Split(lng.TagAttributeNames, Chr$(10))
  Elements = Split(lng.TagElementNames, Chr$(10))
End Sub

Private Sub LangKeywords(lang As String)
  On Error Resume Next
  Dim lng As CodeSenseCtl.Language, UA() As String, X As Long
  Dim glb As CodeSenseCtl.globals
  cboProcedures.Clear
  Set lng = New CodeSenseCtl.Language
  Set glb = New CodeSenseCtl.globals
  Set lng = glb.GetLanguageDef(lang)
  UA = Split(lng.Keywords, Chr$(10))
  
  For X = 0 To UBound(UA) - 1
   cboProcedures.AddItem UA(X)
  Next
  Erase UA
End Sub

