VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mesh Builder"
   ClientHeight    =   2625
   ClientLeft      =   10905
   ClientTop       =   1050
   ClientWidth     =   8655
   Icon            =   "Mesh builder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   2280
      Left            =   6240
      ScaleHeight     =   2220
      ScaleMode       =   0  'User
      ScaleWidth      =   2505.909
      TabIndex        =   13
      Top             =   240
      Width           =   2280
   End
   Begin VB.PictureBox Picture1 
      Height          =   2280
      Left            =   3840
      ScaleHeight     =   2220
      ScaleMode       =   0  'User
      ScaleWidth      =   2505.909
      TabIndex        =   12
      Top             =   240
      Width           =   2280
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Reset"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "place"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   327682
      Min             =   1
      Max             =   20
      SelStart        =   1
      Value           =   1
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "Mesh builder.frx":0442
      Left            =   2400
      List            =   "Mesh builder.frx":0449
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open texture"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Text            =   "10"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Text            =   "10"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "10"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "10"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "10"
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape4 
      Height          =   975
      Left            =   240
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      Height          =   1215
      Left            =   240
      Top             =   240
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   2280
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sizex As Single
Dim sizey As Single
Dim sizez As Single
Dim sizel As Single
Dim sizer As Single
Dim x As Single
Dim y As Single
Dim z As Single
Dim gen_tex As Integer



Sub Vwall(x, y, z, leng, DEC, sx, sy)
SetV x + leng, y + sy, z + DEC, gen_tex, 0
SetV x - leng, y + sy, z + DEC, 0, 0
SetV x - sx, y - sy, z, 0, gen_tex
SetV x + sx, y - sy, z, gen_tex, gen_tex
End Sub

Sub Fwall(x, y, z, leng, DEC, sx, sy)
SetV x + leng, y, z + sy, gen_tex, 0
SetV x - leng, y, z + sy, 0, 0
SetV x - sx, y, z - sy, 0, gen_tex
SetV x + sx, y, z - sy, gen_tex, gen_tex
End Sub

Sub Hwall(x, y, z, leng, DEC, sz, sy)
SetV x + DEC, y + sy, z + leng, gen_tex, 0
SetV x + DEC, y + sy, z - leng, 0, 0
SetV x, y - sy, z - sz, 0, gen_tex
SetV x, y - sy, z + sz, gen_tex, gen_tex
End Sub



Sub cube(x, y, z, sx, sy, sz, leng1, leng2)
DEC2 = (sx + leng1)
dec1 = (sz + leng2)



Vwall x, y, z + sz, -leng1, -dec1, sx, sy
Hwall x + sx, y, z, -leng2, -DEC2, sz, sy

Vwall x, y, z - sz, -leng1, dec1, sx, sy
Hwall x - sx, y, z, -leng2, DEC2, sz, sy

Fwall x, y + sy, z, -leng1, -dec1, -leng1, -leng2
Fwall x, y - sy, z, sx, sy, sx, sz

End Sub

Function SetV(x, y, z, tx, ty)
DX.CreateD3DVertex x, y, z, 0, 0, 0, tx, ty, Obj_ver(Obj_num)
Obj_num = Obj_num + 1
End Function

Private Sub Command1_Click()
On Error GoTo bug

Form3.comd.FileName = ""
Form3.comd.Filter = "BMP (*.bmp)|*.bmp"
Form3.comd.ShowOpen
Image1.Picture = LoadPicture(Form3.comd.FileName)
List1.List(tex_num) = FiEnd(Form3.comd.FileName)
Obj_tex_FILE(tex_num) = Form3.comd.FileName
Set tex(tex_num) = Bmap(Form3.comd.FileName, True)
tex_num = tex_num + 1
usedt = True
List1.ListIndex = 0
Exit Sub
bug:
End Sub

Private Sub Command2_Click()
VerNum2 = VerNum2 + 24
For i = 0 To 24
ver_num(i + VerNum2 - 24) = List1.ListIndex
Vert(i + VerNum2 - 24).x = Obj_ver(i).x
Vert(i + VerNum2 - 24).y = Obj_ver(i).y
Vert(i + VerNum2 - 24).z = Obj_ver(i).z
Vert(i + VerNum2 - 24).tv = Obj_ver(i).tv
Vert(i + VerNum2 - 24).tu = Obj_ver(i).tu

Next

End Sub

Private Sub Command4_Click()
On Error GoTo bug
Form3.comd.FileName = ""
Form3.comd.Filter = "MDL (*.mdl)|*.mdl"
Form3.comd.ShowSave
SAVE_MDL (Form3.comd.FileName)
Exit Sub
bug:
res = MsgBox("Unable to save", vbCritical, "ERROR")
End Sub

Private Sub Command5_Click()
Rexyz x, y, z
Form3.ry x, y, z
End Sub

Private Sub Form_Load()
sizex = 10
sizey = 10
sizez = 10
sizel = -10
sizer = -10

End Sub




Private Sub List1_Click()
Image1.Picture = LoadPicture(Obj_tex_FILE(List1.ListIndex))
End Sub

Private Sub Slider1_Click()
gen_tex = Slider1.Value
End Sub

Private Sub Text1_Change()
On Error GoTo ops
Obj_num = 0
sizex = Text1.Text
cube x, y, z, sizex, sizey, sizez, sizel, sizer
Exit Sub
ops:
sizex = 1
Text1.Text = "1"
End Sub



Private Sub Text2_Change()
On Error GoTo ops
Obj_num = 0
sizey = Text2.Text
cube x, y, z, sizex, sizey, sizez, sizel, sizer
Exit Sub
ops:
sizex = 1
Text2.Text = "1"
End Sub



Private Sub Text3_change()
On Error GoTo ops
Obj_num = 0
sizez = Text3.Text
cube x, y, z, sizex, sizey, sizez, sizel, sizer
Exit Sub

ops:
sizex = 1
Text3.Text = "1"
End Sub

Private Sub Text4_change()
On Error GoTo ops
Obj_num = 0
sizel = -Text4.Text
cube x, y, z, sizex, sizey, sizez, sizel, sizer
Exit Sub

ops:
sizex = -1
Text4.Text = "1"
End Sub

Private Sub Text5_change()
On Error GoTo ops
Obj_num = 0
sizer = -Text5.Text
cube x, y, z, sizex, sizey, sizez, sizel, sizer
Exit Sub

ops:
sizex = -1
Text5.Text = "1"
End Sub

Function setxyz(xx, yy, zz)
x = xx
y = yy
z = zz
End Function

Private Sub Timer1_Timer()
Obj_num = 0

cube x, y, z, sizex, sizey, sizez, sizel, sizer
End Sub
