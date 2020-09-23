VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About cEdit"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   3780
      TabIndex        =   1
      Top             =   0
      Width           =   3780
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         Caption         =   "About cEdit"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "frmAbout.frx":1042
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 4.6.1"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Line LineSep 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   60
      X2              =   3700
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line LineSep 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   60
      X2              =   3700
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Label lblInfo 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "http://cedit.sourceforge.net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   2835
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  FlatBorder cmdOk.hwnd
  LoadFormData Me
  Label2.ForeColor = vbBlue
  lblData(1) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  Me.Move frmMain.Left + ((frmMain.Width - Me.Width) \ 2), frmMain.Top + ((frmMain.Height - Me.Height) \ 2)
  lblInfo.Caption = "cEdit Version " & App.Major & "." & App.Minor & "." & App.Revision & Chr(10) & "Copyright (c)1999-2004 cEdit Software" & Chr(10) & Chr(10) & "Open Source, Freeware Code Editor"
  lblInfo.Caption = lblInfo.Caption & Chr(10) & "Developers: Stewart, Junchi"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormData Me
End Sub

Private Sub Label2_Click()
  OpenURL "http://cedit.sourceforge.net", Me.hwnd
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then Label2.ForeColor = vbRed
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label2.ForeColor = vbBlue
End Sub
