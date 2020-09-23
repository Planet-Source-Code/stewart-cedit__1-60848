VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNav 
   Caption         =   "Quick Nav"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSnippet 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   1080
      ScaleHeight     =   2655
      ScaleWidth      =   2055
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      Begin MSComctlLib.ListView lstSnippet 
         Height          =   1815
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   3201
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         _Version        =   393217
         Icons           =   "images"
         SmallIcons      =   "images"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4440
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2850
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   270
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   2640
      ScaleHeight     =   3555
      ScaleWidth      =   2505
      TabIndex        =   2
      Top             =   -600
      Width           =   2505
      Begin VB.PictureBox picSize 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   240
         ScaleHeight     =   45
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   300
         TabIndex        =   4
         Top             =   0
         Width           =   2235
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   3
         Top             =   420
         Width           =   2220
      End
      Begin MSComctlLib.ListView File1 
         Height          =   1710
         Left            =   480
         TabIndex        =   5
         Top             =   2040
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   3016
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgSize 
         Height          =   50
         Left            =   360
         MouseIcon       =   "frmNav.frx":0000
         MousePointer    =   99  'Custom
         Top             =   1920
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   2160
      ScaleHeight     =   2595
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   2265
      Begin MSComctlLib.ImageList images 
         Left            =   1500
         Top             =   390
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNav.frx":0152
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNav.frx":06A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TagsD 
         Height          =   1530
         Left            =   1725
         TabIndex        =   1
         Top             =   90
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   2699
         _Version        =   393217
         Indentation     =   5
         LineStyle       =   1
         Style           =   7
         ImageList       =   "images"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.TabStrip Tbs 
      Height          =   4335
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   7646
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tags"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Snippets"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483644
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483644
      _Version        =   393216
   End
End
Attribute VB_Name = "frmNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal I&, ByVal hDCDest&, _
    ByVal x&, ByVal Y&, ByVal FLAGS&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

Sub FillFile1WithFiles(ByVal path As String)
'-------------------------------------------
'Scan the selected folder for files
'and add then to the listview
'-------------------------------------------
Dim Item As ListItem
Dim s As String

path = CheckPath(path)    'Add '\' to end if not present
s = Dir(path, vbNormal)
Do While s <> ""
  Set Item = File1.ListItems.Add(, , s)
  Item.Key = path & s
  'Item.SmallIcon = "Folder"
  Item.Text = s
  Item.SubItems(1) = path
  s = Dir
Loop

End Sub
Private Sub Form_Load()
  pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
  pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
  pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
  pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY
  imgSize.Top = 1920
  Dir1_Change
  AddSnippets
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Tbs.Move 30 + frmMain.fDock.DockedFormCaptionOffsetLeft("frmNav"), 30 + frmMain.fDock.DockedFormCaptionOffsetTop("frmNav"), Me.ScaleWidth - 60 - frmMain.fDock.DockedFormCaptionOffsetLeft("frmNav") - 60, Me.ScaleHeight - 60 - frmMain.fDock.DockedFormCaptionOffsetTop("frmNav") - 60
  Picture4.Move Tbs.ClientLeft, Tbs.ClientTop, Tbs.ClientWidth, Tbs.ClientHeight
  Picture5.Move Tbs.ClientLeft, Tbs.ClientTop, Tbs.ClientWidth, Tbs.ClientHeight
  picSnippet.Move Tbs.ClientLeft, Tbs.ClientTop, Tbs.ClientWidth, Tbs.ClientHeight
  TagsD.Move 0, 30, Picture5.ScaleWidth, Picture5.ScaleHeight - 30
End Sub

Private Sub Dir1_Change()
Dim path As String

Initialise
path = Dir1.path
FillFile1WithFiles path
GetAllIcons
ShowIcons
End Sub

Private Sub Drive1_Change()
  Dir1.path = Drive1.Drive
End Sub


Private Sub Initialise()
'-----------------------------------------------
'Initialise the controls
'-----------------------------------------------
On Local Error Resume Next

'Break the link to iml lists
File1.ListItems.Clear
File1.icons = Nothing
File1.SmallIcons = Nothing

'Clear the image lists
iml32.ListImages.Clear
iml16.ListImages.Clear

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim filename As String

On Local Error Resume Next
For Each Item In File1.ListItems
  filename = Item.SubItems(1) & Item.Text
  GetIcon filename, Item.Index
Next

End Sub

Private Function GetIcon(filename As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection



'Get a handle to the small icon
hSIcon = SHGetFileInfo(filename, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(filename, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    ImageList_Draw hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    ImageList_Draw hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function
Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the File1
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With File1
  '.ListItems.Clear
  .icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub


Private Sub File1_DblClick()
  DoOpen Dir1.path & "\" & File1.SelectedItem.Text
End Sub

Private Sub imgSize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picSize.Visible = True
End Sub

Private Sub imgSize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim nxtY As Long
  If Button = 1 Then
    nxtY = (imgSize.Top + Y)
    If nxtY < 800 Then nxtY = 800
    If nxtY > (Picture4.ScaleHeight - 800) Then nxtY = Picture4.Height - 800
    picSize.Top = nxtY
    imgSize.Move picSize.Left, picSize.Top, picSize.Width, picSize.Height
  End If
End Sub

Private Sub imgSize_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picSize.Visible = False
  Resize
End Sub

Private Sub Resize()
  On Error Resume Next
  imgSize.Left = 0
  imgSize.Width = Picture4.ScaleWidth
  picSize.Move 0, imgSize.Top, imgSize.Width, imgSize.Height
  Drive1.Move 0, 30, Picture4.ScaleWidth
  Dir1.Move 0, Drive1.Top + Drive1.Height + 30, Picture4.ScaleWidth, imgSize.Top - Dir1.Top
  If Dir1.Height > (Picture4.ScaleHeight - 1500) Then Dir1.Height = Picture4.ScaleHeight - 1500
  imgSize.Move 0, Dir1.Top + Dir1.Height, Picture4.ScaleWidth
  File1.Move 0, imgSize.Top + imgSize.Height, Picture4.ScaleWidth, Picture4.Height - (imgSize.Top + imgSize.Height)
End Sub
 
Private Sub Picture1_Click()

End Sub

Private Sub Picture1_Resize()

End Sub

Private Sub lstSnippet_DblClick()
  On Error Resume Next
  Dim fFile As Integer, str As String
  fFile = FreeFile()
  Open App.path & "\snippets\" & lstSnippet.SelectedItem.Text & ".snippet" For Input As #fFile
    str = Input(LOF(fFile), fFile)
  Close #fFile
  Call InsertString(Document(dnum).rt, str)
End Sub

Private Sub lstSnippet_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
  On Error Resume Next
    Dim OLEFilename As String, ext As String, file2 As String
    Dim I As Integer
    For I = 1 To Data.Files.Count
        If Data.GetFormat(vbCFFiles) Then
            OLEFilename = Data.Files(I)
        End If
        On Error GoTo errexit
       ext = GetExtension(OLEFilename)
       
       ext = Left(OLEFilename, Len(OLEFilename) - (Len(ext) + 1))
       file2 = StripPath(ext)
       CopyFile OLEFilename, App.path & "\snippets\" & file2 & ".snippet", False
    Next I
    AddSnippets
errexit:
    Exit Sub
End Sub

Private Sub lstSnippet_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
  On Error Resume Next
    If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
End Sub

Private Sub picSnippet_Resize()
  lstSnippet.Move 0, 0, picSnippet.ScaleWidth, picSnippet.ScaleHeight
End Sub

Private Sub Picture4_Resize()
  Resize
End Sub

Private Sub tagsd_DblClick()
  Dim timedate As String
  On Error Resume Next
  Dim r As CodeSenseCtl.range
  Set r = New CodeSenseCtl.range
  timedate = TagsD.SelectedItem.Text
  Document(dnum).rt.SelText = timedate
  Set r = Document(dnum).rt.GetSel(False)
  Document(dnum).rt.SetCaretPos r.StartLineNo + 1, r.StartColNo + Len(timedate)
  Document(dnum).rt.SetFocus
End Sub

Private Sub tbs_Click()
  Picture4.Visible = False
  Picture5.Visible = False
  picSnippet.Visible = False
  If Tbs.SelectedItem.Index = 1 Then
    Picture4.Visible = True
  ElseIf Tbs.SelectedItem.Index = 2 Then
    Picture5.Visible = True
  ElseIf Tbs.SelectedItem.Index = 3 Then
    picSnippet.Visible = True
  End If
End Sub

Private Sub AddSnippets()
  Dim s As String
  s = Dir(App.path & "\snippets\")
  lstSnippet.ListItems.Clear
  Do Until s = ""
    If Right(s, 7) = "snippet" Then
      lstSnippet.ListItems.Add , , Left(s, Len(s) - 8), 2, 2
    End If
    s = Dir
  Loop
End Sub
