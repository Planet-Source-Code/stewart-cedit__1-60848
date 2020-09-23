VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmTemplate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Template Editor"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmTemplate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CodeSenseCtl.CodeSense rt 
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "frmTemplate.frx":1042
      TabIndex        =   7
      Top             =   1560
      Width           =   6495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboLang 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   5295
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "Language:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Template Name:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Data"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboLang_Click()
  On Error Resume Next
  If cboLang.ListIndex = 0 Then rt.Language = "" Else rt.Language = cboLang.Text
  rt.SetFocus
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim fFile As Integer, str As String, file As String
  str = cboLang.Text & vbCrLf & rt.Text
  If Left(str, 1) = """" Then str = Right(str, Len(str) - 1)
  If Right(str, 1) = """" Then str = Left(str, Len(str) - 1)
  fFile = FreeFile()
  file = Replace(txtName.Text, "/", "")
  file = Replace(file, "\", "")
  file = App.path & "\templates\" & file & ".tmp"
  rt.Text = cboLang.Text & vbCrLf & rt.Text
  rt.SaveFile file, False
  rt.Text = ""
  txtName.Text = ""
End Sub

Private Sub Form_Load()
  On Error Resume Next
  LoadFormData Me
  AddLang
  ReadOptions rt
End Sub

Private Sub AddLang()
  Dim UA() As String, LngCnt As Long
  cboLang.Clear
  cboLang.AddItem "Text"
  cboLang.AddItem "C/C++"
  cboLang.AddItem "Basic"
  cboLang.AddItem "Java"
  cboLang.AddItem "Pascal"
  cboLang.AddItem "SQL"
  cboLang.AddItem "HTML"
  cboLang.AddItem "XML"
  cboLang.Text = "Text"
  UA = Split(Langs, Chr$(10))
  For LngCnt = 0 To UBound(UA) - 1: cboLang.AddItem UA(LngCnt): Next
  Erase UA
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub
