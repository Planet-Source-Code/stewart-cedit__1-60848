Attribute VB_Name = "modGeneral"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Public Result As String
Public Sub FlatBorder(ByVal hwnd As Long)
  Dim TFlat As Long
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Function InputStr(Optional Question As String, Optional WinTitle As String, Optional default As String, Optional Start As Integer, Optional IconFile As String) As String
  If Question <> "" Then
    frmInput.lblInfo.Caption = Question
  End If
  If WinTitle <> "" Then
    frmInput.Caption = WinTitle
  Else
    frmInput.Caption = App.Title
  End If
  If default <> "" Then
    frmInput.txtInput.Text = default
  End If
  If IconFile <> "" Then
    frmInput.picIcon.Picture = LoadPicture(IconFile)
  End If
  If Start <> 0 Then
    frmInput.txtInput.SelStart = Start
  End If
  frmInput.Show vbModal
  InputStr = Result
End Function


Public Sub GetAccounts(cbo As ComboBox)
  Dim s As String
  cbo.Clear
  s = Dir(App.path & "\accounts\")
  Do While s <> ""
    If Right(s, 3) = "ftp" Then
      cbo.AddItem Left(s, Len(s) - 4)
    End If
    s = Dir
  Loop
End Sub

Public Function StrWrap(str As String) As String
  StrWrap = """" & str & """"
End Function

Public Function SplitStr(str As String, ReturnStr As String) As String
  ' This will split a set of words in quotes
  ' Used in the ftp portion
  Dim FindQ As Long, FindQ2 As Long
  FindQ = InStr(1, str, """")
  If FindQ = 0 Then
    SplitStr = ""
    Exit Function
  End If
  FindQ2 = InStr(FindQ + 1, str, """")
  If FindQ2 = 0 Then
    SplitStr = ""
    Exit Function
  End If
  SplitStr = Mid(str, FindQ + 1, FindQ2 - 2)
  ReturnStr = Mid(str, FindQ2 + 1, Len(str) - FindQ2)
End Function


Public Sub Flatten(ByVal frm As Form)
  Dim CTL As Control
  For Each CTL In frm.Controls
    Select Case TypeName(CTL)
      Case "CommandButton", "TextBox", "ListBox", "FileTree", "TreeView", "ProgressBar", "PictureBox"
        FlatBorder CTL.hwnd
    End Select
  Next
End Sub

