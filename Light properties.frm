VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Light Point Properties"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3675
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command1"
      Height          =   135
      Left            =   3120
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command1"
      Height          =   135
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   135
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   135
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   135
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   135
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "0"
      Top             =   1080
      Width           =   975
   End
   Begin ComctlLib.ProgressBar pb1 
      Height          =   135
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
      Max             =   10
   End
   Begin ComctlLib.ProgressBar pb2 
      Height          =   135
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
      Max             =   10
   End
   Begin ComctlLib.ProgressBar pb3 
      Height          =   135
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
      Max             =   10
   End
   Begin VB.Label Label4 
      Caption         =   "B"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "G"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "R"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Light range"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If pb1.Value <> 0 Then pb1.Value = pb1.Value - 1
Light.x = pb1.Value

End Sub

Private Sub Command2_Click()
If pb2.Value <> 0 Then pb2.Value = pb2.Value - 1
Light.y = pb2.Value
End Sub

Private Sub Command3_Click()
If pb3.Value <> 0 Then pb3.Value = pb3.Value - 1
Light.z = pb3.Value
End Sub

Private Sub Command4_Click()
If pb1.Value <> 10 Then pb1.Value = pb1.Value + 1
Light.x = pb1.Value
End Sub

Private Sub Command5_Click()
If pb2.Value <> 10 Then pb2.Value = pb2.Value + 1
Light.y = pb2.Value
End Sub

Private Sub Command6_Click()
If pb3.Value <> 10 Then pb3.Value = pb3.Value + 1
Light.z = pb3.Value
End Sub

Private Sub Text1_Change()

On Error GoTo bug

Light.y1 = Text1.Text

Exit Sub
bug:
res = MsgBox("The value must be numeric", vbCritical, "Error")
Light.y1 = 0
End Sub
