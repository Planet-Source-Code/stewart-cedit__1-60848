VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Associations"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Cool"
      Default         =   -1  'True
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   300
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Forget It"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Help"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   1500
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   300
      Width           =   2295
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "HTML Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Text Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Basic Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CGI Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pascal Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "XML Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "C/C++ Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Java Script Files"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select File Associations:"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   3495
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ChkRemoved(7) As Boolean
Dim WasChked(7) As Boolean

Private Sub chk_Click(Index As Integer)
  If chk(Index).Value = 0 And WasChked(Index) = True Then
    ChkRemoved(Index) = True
  Else
    ChkRemoved(Index) = False
  End If
End Sub

Private Sub Command1_Click()
Dim ms As String, x As Integer, icons As String, exepat As String, todo As String
If chk(0).Value = 1 And WasChked(0) = False Then

  MakeFileType "htm", "cEdit.icoument", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
  MakeFileType "html", "cEdit.icoument", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If
If chk(1).Value = 1 And WasChked(1) = False Then
  MakeFileType "txt", "Text Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If
If chk(2).Value = 1 And WasChked(2) = False Then
  MakeFileType "bas", "Basic Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
  MakeFileType "frm", "Basic Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
  MakeFileType "vbp", "Basic Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If
If chk(3).Value = 1 And WasChked(3) = False Then
  MakeFileType "pl", "CGI Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
  MakeFileType "cgi", "CGI Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If
If chk(4).Value = 1 And WasChked(4) = False Then
  MakeFileType "xml", "XML Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If
If chk(5).Value = 1 And WasChked(5) = False Then
  MakeFileType "pas", "Pascal Files", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If
If chk(6).Value = 1 And WasChked(6) = False Then
  MakeFileType "c", "C Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
  MakeFileType "cpp", "C++ Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
  MakeFileType "h", "Header Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If
If chk(7).Value = 1 And WasChked(7) = False Then
    MakeFileType "js", "Java Script Document", App.path & "\grx\cEdit.ico", "open", App.path & "\cEdit.exe %1", False, True, True
End If

For x = 0 To 7
  If ChkRemoved(x) = True Then
      If x = 0 Then
        icons = ReadINI("Reg", "htmico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "htmexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "htmact", App.path & "\reg.ini")
        MakeFileType "htm", "cEdit.icoument", icons, todo, exepat, False, True, False
        
        icons = ReadINI("Reg", "htmlico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "htmlexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "htmlact", App.path & "\reg.ini")
        MakeFileType "html", "cEdit.icoument", icons, todo, exepat, False, True, False
        
     ElseIf x = 1 Then
        icons = ReadINI("Reg", "txtico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "txtexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "txtact", App.path & "\reg.ini")
        MakeFileType "txt", "Text Document", icons, todo, exepat, False, True, False
        
     ElseIf x = 2 Then
        icons = ReadINI("Reg", "basico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "basexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "basact", App.path & "\reg.ini")
        MakeFileType "bas", "Basic Document", icons, todo, exepat, False, True, False
                
        icons = ReadINI("Reg", "frmico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "frmexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "frmact", App.path & "\reg.ini")
        MakeFileType "frm", "Basic Document", icons, todo, exepat, False, True, False
                
        icons = ReadINI("vbp", "vbpico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "vbpexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "vbpact", App.path & "\reg.ini")
        MakeFileType "vbp", "Basic Document", icons, todo, exepat, False, True, False
        
     ElseIf x = 3 Then
        icons = ReadINI("Reg", "plico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "plexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "plact", App.path & "\reg.ini")
        MakeFileType "pl", "CGI Document", icons, todo, exepat, False, True, False
       
        icons = ReadINI("Reg", "cgiico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "cgiexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "cgiact", App.path & "\reg.ini")
        MakeFileType "cgi", "CGI Document", icons, todo, exepat, False, True, False
       
     ElseIf x = 4 Then
        icons = ReadINI("Reg", "xmlico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "xmlexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "xmlact", App.path & "\reg.ini")
        MakeFileType "xml", "XML Document", icons, todo, exepat, False, True, False
    
     ElseIf x = 5 Then
        icons = ReadINI("Reg", "pasico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "pasexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "pasact", App.path & "\reg.ini")
        MakeFileType "pas", "Pascal Document", icons, todo, exepat, False, True, False
    
     ElseIf x = 6 Then
        icons = ReadINI("Reg", "cico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "cexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "cact", App.path & "\reg.ini")
        MakeFileType "c", "C/C++ Document", icons, todo, exepat, False, True, False
                
        icons = ReadINI("Reg", "cppico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "cppexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "cppact", App.path & "\reg.ini")
        MakeFileType "c", "C/C++ Document", icons, todo, exepat, False, True, False
                
        icons = ReadINI("Reg", "hico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "hexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "hact", App.path & "\reg.ini")
        MakeFileType "h", "Header Document", icons, todo, exepat, False, True, False
    
     ElseIf x = 7 Then
        icons = ReadINI("Reg", "jsico", App.path & "\reg.ini")
        exepat = ReadINI("Reg", "jsexe", App.path & "\reg.ini")
        todo = ReadINI("Reg", "jact", App.path & "\reg.ini")
        MakeFileType "js", "Java Script Document", icons, todo, exepat, False, True, False
    End If
  End If
Next
For x = 0 To 7
  ms = ReplaceChars(chk(x).Caption, " ", "_")
  writeini "regdata", ms, chk(x).Value, App.path & "\reg.ini"
Next
Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  MsgBox "With this you can choose which files will be opened by this program." & Chr(10) & Chr(10) & "Warning..." & Chr(10) & "Do not use this if you plan on deleting" & Chr(10) & "this application soon.", vbOKOnly + vbQuestion, "Help.."
End Sub

Private Sub Form_Load()
Dim x As Integer, ms As String
LoadFormData Me
For x = 0 To 7
  ChkRemoved(x) = False
  ms = ReplaceChars(chk(x).Caption, " ", "_")
  chk(x).Value = ReadINI("regdata", ms, App.path & "\reg.ini")
  If chk(x).Value = 1 Then
    WasChked(x) = True
  Else
    WasChked(x) = False
  End If
  Me.Move frmMain.Left + ((frmMain.Width - Me.Width) \ 2), frmMain.Top + ((frmMain.Height - Me.Height) \ 2)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub
