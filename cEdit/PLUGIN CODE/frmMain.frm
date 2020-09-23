VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Plugin"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "&MsgBox"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Reverse Text"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Display Text"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Close"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Clear Text"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Insert String"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Increase Hosts Size"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Decrease Hosts Size"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Host"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide Host"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Host's Caption"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objHost As Object

Private Sub Command1_Click()
objHost.Visible = False
End Sub

Private Sub Command10_Click()
  Dim p As String, x As Long, newstr As String
  p = objHost.ActiveForm.rt.Text
  For x = Len(p) To 1 Step -1
    newstr = newstr & Mid(p, x, 1)
  Next
  objHost.ActiveForm.rt.Text = newstr
End Sub

Private Sub Command11_Click()
  objHost.MessageBox "Hello", vbYesNo + vbCritical, "This is a test"
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command2_Click()
objHost.Visible = True
End Sub

Private Sub Command3_Click()
objHost.Caption = Text1
End Sub

Private Sub Command4_Click()
objHost.Width = objHost.Width - 500
End Sub

Private Sub Command5_Click()
objHost.Width = objHost.Width + 500
End Sub

Private Sub Command6_Click()
  objHost.addtext "Inserts Text"
End Sub

Private Sub Command7_Click()
  objHost.ActiveForm.rt.Text = ""
End Sub

Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub Command9_Click()
  MsgBox objHost.ActiveForm.rt.Text
End Sub

Private Sub Label1_Click()

End Sub


Private Sub Form_Unload(Cancel As Integer)
  objHost.fdock.FormUndock Me.Name
End Sub
