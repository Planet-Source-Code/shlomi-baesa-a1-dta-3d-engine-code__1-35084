VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fog Properties"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3450
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Enter"
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2160
      Picture         =   "Fog Properties.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   2760
      Picture         =   "Fog Properties.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "(R:G:b)"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   ":"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   ":"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim thick As Integer
Dim R As Single
Dim G As Single
Dim B As Single

Private Sub Command1_Click()
Fognum = Fognum - 5
Call Fog(Fognum, Fogr, Fogg, Fogb)
End Sub

Private Sub Command2_Click()
Fognum = Fognum + 5
Call Fog(Fognum, Fogr, Fogg, Fogb)
End Sub


Private Sub Command3_Click()
On Error GoTo num
res = InputBox("Enter a number between 0 to 1000", "Fog")
Fognum = res
Call Fog(Fognum, Fogr, Fogg, Fogb)
Exit Sub
num:
res = MsgBox("The value must be numeric", vbCritical, "Error")
Fognum = 1
End Sub

Private Sub Form_Click()
Form3.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Enabled = True
Form3.Show
End Sub

Private Sub Text1_LostFocus()
On Error GoTo num
Fogr = Text1.Text

Exit Sub
num:
res = MsgBox("The value must be numeric", vbCritical, "Error")
Text1.Text = 0
Fogr = Text1.Text
End Sub


Private Sub Text2_LostFocus()
On Error GoTo num
Fogg = Text2.Text

Exit Sub
num:
res = MsgBox("The value must be numeric", vbCritical, "Error")
Text2.Text = 0
Fogg = Text2.Text
End Sub

Private Sub Text3_LostFocus()
On Error GoTo num
Fogb = Text3.Text

Exit Sub
num:
res = MsgBox("The value must be numeric", vbCritical, "Error")
Text3.Text = 0
Fogb = Text3.Text
End Sub
