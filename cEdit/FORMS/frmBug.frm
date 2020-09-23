VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1558013-91A7-11D4-AA5B-00A0CC334D72}#2.0#0"; "WWTabs.ocx"
Begin VB.Form frmBug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Bug Log"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgMain 
      Left            =   2085
      Top             =   1485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   9
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBug.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBug.frx":04E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBug.frx":09C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Height          =   1335
      Left            =   600
      ScaleHeight     =   1275
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin cEdit.ctlFrame cmdTBar 
         Height          =   1335
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   2355
         Begin MSComctlLib.Toolbar tbBug 
            Height          =   810
            Left            =   60
            TabIndex        =   4
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   1429
            ButtonWidth     =   423
            ButtonHeight    =   476
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "imgMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtMain 
         ForeColor       =   &H00C0C0C0&
         Height          =   975
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComctlLib.ListView lstTask 
         Height          =   780
         Left            =   840
         TabIndex        =   1
         Top             =   195
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Task ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Completed"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
      Begin cEdit.VSFileSearch vs 
         Height          =   1935
         Left            =   480
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
      End
   End
   Begin WWTabs.WTabs wt 
      Height          =   300
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionTips     =   "|"
      Captions        =   "&Todo List|Results"
   End
End
Attribute VB_Name = "frmBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_Load()
  FlatBorder picMain.hwnd
  LoadTasks
End Sub

Private Sub LoadTasks()
  Dim iFile As Integer, iPos As Integer
  Dim strStore As String, strTemp As String
  Dim strHold(2) As String
  
  
  iFile = FreeFile
  Open App.path + "\data\tasks.dat" For Input As #iFile
  Do Until EOF(iFile)
    Input #iFile, strStore
    iPos = InStr(1, strStore, "|")
    strTemp = Left(strStore, iPos - 1)
    strHold(0) = strTemp
    strTemp = Mid(strStore, iPos + 1)
    iPos = InStr(1, strTemp, "|")
    strHold(1) = Left(strTemp, iPos - 1)
    strHold(2) = Mid(strTemp, iPos + 1)
    With lstTask.ListItems.Add
      .SubItems(1) = strHold(0)
      .SubItems(2) = strHold(1)
      .SubItems(3) = strHold(2)
    End With
  Loop
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  picMain.Move frmMain.fDock.DockedFormCaptionOffsetLeft("frmBug"), frmMain.fDock.DockedFormCaptionOffsetTop("frmBug"), Me.ScaleWidth - frmMain.fDock.DockedFormCaptionOffsetLeft("frmBug"), Me.ScaleHeight - frmMain.fDock.DockedFormCaptionOffsetTop("frmBug") - wt.Height
  wt.Move picMain.Left, picMain.Top + picMain.Height, picMain.Width

End Sub

Private Sub picMain_Resize()
  On Error Resume Next
  cmdTBar.Move 20, 20, cmdTBar.Width, picMain.ScaleHeight - 45
  tbBug.Move 75, 75
'  tbMain.Move cmdTbar.Width - 10, 0, picMain.ScaleWidth + 10 - cmdTbar.Width, picMain.ScaleHeight + 10
  lstTask.Move cmdTBar.Width + 10, 10, picMain.ScaleWidth - cmdTBar.Width - 30, picMain.ScaleHeight - 45
  lstTask.ColumnHeaders(4).Width = lstTask.Width - lstTask.ColumnHeaders(2).Width - lstTask.ColumnHeaders(3).Width - 80
  'rbrMain.Move 60, 60, rbrMain.Width, picMain.ScaleHeight - 130
  txtMain.Move lstTask.Left + 15, lstTask.Top + 15, lstTask.Width - 30, lstTask.Height - 30
  vs.Move cmdTBar.Width + 10, 0, picMain.ScaleWidth - cmdTBar.Width - 10, picMain.ScaleHeight
'  fSearch.Move lstTask.Left + 15, lstTask.Top + 15, lstTask.Width - 30, lstTask.Height - 30
  cmdTBar.Move 20, txtMain.Top, cmdTBar.Width, txtMain.Height
  tbBug.Move 60, 60
  
End Sub



Private Sub tbBug_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Dim iFile, iVal As Integer
  Select Case Button.Index
    Case 1
      frmTask.bEdit = False
      frmTask.txtName.Text = ""
      frmTask.txtDesc.Text = ""
      frmTask.hPer.Value = 0
      frmTask.Show vbModal, frmMain
    Case 2
      frmTask.bEdit = True
      frmTask.iItemNum = lstTask.SelectedItem.Index
      frmTask.txtName.Text = lstTask.SelectedItem.SubItems(1)
      frmTask.txtDesc.Text = lstTask.SelectedItem.SubItems(3)
      iVal = InStr(1, lstTask.SelectedItem.SubItems(2), "%")
      frmTask.hPer.Value = (Left(lstTask.SelectedItem.SubItems(2), iVal - 1) \ 10)
      frmTask.Show vbModal, frmMain
    Case 3
      lstTask.ListItems.Remove (lstTask.SelectedItem.Index)
  End Select
  iFile = FreeFile
  Open App.path + "\data\tasks.dat" For Output As #iFile
  For iVal = 0 To lstTask.ListItems.Count
      Print #iFile, lstTask.ListItems(iVal).SubItems(1) + "|" + lstTask.ListItems(iVal).SubItems(2) + "|" + lstTask.ListItems(iVal).SubItems(3)
  Next
  Close #iFile
End Sub

Private Sub WTabs1_Click(ByVal ActualClick As Boolean)
End Sub

Private Sub vs_DblClick(SelectedFile As String, LineNumber As String)
  On Error Resume Next
  Dim i As Long
  DoOpen SelectedFile
  Document(dnum).rt.SetCaretPos LineNumber - 1, 0
End Sub

Private Sub wt_Click(ByVal ActualClick As Boolean)
  txtMain.Visible = False
  lstTask.Visible = False
  vs.Visible = False
  Select Case wt.ActiveTab
    Case 0
      lstTask.Visible = True
    Case 1
      vs.Visible = True
    Case Else
      txtMain.Visible = True
  End Select

End Sub
