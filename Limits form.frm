VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Limits Drawer"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4545
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose Picture"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2705
      Left            =   120
      ScaleHeight     =   2640
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldX As Single
Dim oldy As Single

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If oldX And oldy <> 0 Then

Dim Xtt As Single
Dim Ytt As Single
Dim Oxt As Single
Dim Oyt As Single
Xtt = x / 100
Ytt = y / 100
Oxt = oldX / 100
Oyt = oldy / 100

m = (Ytt - Oyt) / (Xtt - Oxt)
c = Ytt - m * Xtt

Dim xtag As Single

For xtag = Xtt To Oxt

Picture1.PSet (xtag * 100, (xtag * m + c) * 100), vbRed
Next

End If

oldX = x
oldy = y

Form5.Caption = Oxt

End Sub
