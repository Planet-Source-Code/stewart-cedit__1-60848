VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug - stdout"
   ClientHeight    =   3255
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Menu MEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu MCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu MSave 
         Caption         =   "&Save to file..."
         Enabled         =   0   'False
      End
      Begin VB.Menu MNone1 
         Caption         =   "-"
      End
      Begin VB.Menu MClear 
         Caption         =   "C&lear"
      End
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
  On Error Resume Next
  txtOut.Move 0 + frmMain.fDock.DockedFormCaptionOffsetLeft("frmOutput"), 0 + frmMain.fDock.DockedFormCaptionOffsetTop("frmOutput"), Me.ScaleWidth - frmMain.fDock.DockedFormCaptionOffsetLeft("frmOutput") - 80, Me.ScaleHeight - frmMain.fDock.DockedFormCaptionOffsetTop("frmOutput") - 80

End Sub

Private Sub MClear_Click()
txtOut.Text = ""
End Sub

Private Sub MCopy_Click()
Clipboard.SetText txtOut.Text

End Sub

Private Sub txtOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu MEdit
End If
End Sub
