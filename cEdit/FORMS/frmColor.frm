VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   45
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox picBak 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   300
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.PictureBox picClr 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   45
      Picture         =   "frmColor.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   45
      Width           =   750
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
Dim RGBColorHex As String

Private Sub picBak_Click()
  On Error Resume Next
  RGBColorHex = LNGtoHEX(picBak.BackColor)
  InsertString Document(dnum).rt, RGBColorHex
End Sub

Private Sub picClr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  If Button = 1 Then picFore.BackColor = picClr.Point(X, Y)
  If Button = 2 Then picBak.BackColor = picClr.Point(X, Y)
End Sub

Private Sub picClr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  If Button = 1 Then picFore.BackColor = picClr.Point(X, Y)
  If Button = 2 Then picBak.BackColor = picClr.Point(X, Y)
End Sub

Private Sub picFore_Click()
  On Error Resume Next
  RGBColorHex = LNGtoHEX(picFore.BackColor)
  InsertString Document(dnum).rt, RGBColorHex
End Sub

Private Function LNGtoHEX(ByVal lColor As Long) As String

    Dim b(2) As Byte
    
    CopyMemory b(0), lColor, 3

    ' You can't just Hex$ a long to get a web-ready hex triplet color string,
    ' it'll be rerversed (i.e. ff6034 instead of 3460ff).
    LNGtoHEX = "#" & Right$("00000" & LCase$(Hex$(RGB(b(2), b(1), b(0)))), 6)
    
End Function
