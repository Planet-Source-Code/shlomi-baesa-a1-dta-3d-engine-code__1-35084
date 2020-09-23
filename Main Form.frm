VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "DTA -  3D  Engine"
   ClientHeight    =   7980
   ClientLeft      =   435
   ClientTop       =   1785
   ClientWidth     =   10185
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7980
   ScaleWidth      =   10185
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command9 
      Caption         =   "Light"
      Height          =   375
      Left            =   5640
      Picture         =   "Main Form.frx":0442
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Fog"
      Height          =   375
      Left            =   4920
      Picture         =   "Main Form.frx":0884
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Background"
      Height          =   375
      Left            =   3720
      Picture         =   "Main Form.frx":0CC6
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Map"
      Height          =   375
      Left            =   3000
      Picture         =   "Main Form.frx":1108
      TabIndex        =   1
      ToolTipText     =   "Open 3DM map"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Place"
      Height          =   375
      Left            =   2280
      Picture         =   "Main Form.frx":154A
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin ComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   7725
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Model"
      Height          =   375
      Left            =   1560
      Picture         =   "Main Form.frx":198C
      TabIndex        =   2
      ToolTipText     =   "Add a model"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open"
      Height          =   375
      Left            =   840
      Picture         =   "Main Form.frx":1DCE
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      Picture         =   "Main Form.frx":2210
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   0
      Top             =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   5160
   End
   Begin MSComDlg.CommonDialog comd 
      Left            =   0
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   7035
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   600
      Width           =   10125
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   -100
      Top             =   0
      Width           =   10215
   End
   Begin VB.Menu fi 
      Caption         =   "File"
      Begin VB.Menu saveing 
         Caption         =   "Save"
      End
      Begin VB.Menu opening 
         Caption         =   "open"
      End
      Begin VB.Menu ending 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu opt 
      Caption         =   "Options"
      Begin VB.Menu adding_model 
         Caption         =   "Add model"
      End
      Begin VB.Menu drawl 
         Caption         =   "Draw limits"
         Checked         =   -1  'True
      End
      Begin VB.Menu sharp 
         Caption         =   "Sharpness"
      End
      Begin VB.Menu fogy 
         Caption         =   "Fog"
         Begin VB.Menu fog_prop 
            Caption         =   "Properties"
         End
      End
      Begin VB.Menu add_ligh 
         Caption         =   "Source light"
         Begin VB.Menu light_prop 
            Caption         =   "Properties"
         End
         Begin VB.Menu rmove_ligh 
            Caption         =   "Remove light"
         End
      End
      Begin VB.Menu mdl_sz 
         Caption         =   "Model size (%)"
      End
      Begin VB.Menu ve 
         Caption         =   "View"
         Begin VB.Menu solidp 
            Caption         =   "Solid"
         End
         Begin VB.Menu wireframe 
            Caption         =   "Wireframe"
         End
      End
      Begin VB.Menu ambie 
         Caption         =   "Ambiet light"
      End
      Begin VB.Menu map_sizer 
         Caption         =   "Map size (%)"
      End
      Begin VB.Menu set_en 
         Caption         =   "Set Enterance"
      End
   End
   Begin VB.Menu tol 
      Caption         =   "Tools"
      Begin VB.Menu bmpto 
         Caption         =   "BMP to 3DM"
      End
      Begin VB.Menu mesh 
         Caption         =   "Mesh Builder"
      End
      Begin VB.Menu plac 
         Caption         =   "Place Model "
      End
   End
   Begin VB.Menu texture 
      Caption         =   "Texture"
      Begin VB.Menu anima 
         Caption         =   "Animation"
         Begin VB.Menu low 
            Caption         =   "Lower"
            Checked         =   -1  'True
         End
         Begin VB.Menu high 
            Caption         =   "Higher"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu he 
      Caption         =   "Help"
      Begin VB.Menu ab 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim obj_x As Single
Dim obj_y As Single
Dim obj_z As Single


Private Sub ab_Click()
frmAbout.Show
End Sub

Private Sub ambie_Click()
On Error GoTo num
res = InputBox("Enter a number between 0 to 1", "Ambiet light")
Light_Power = res
Exit Sub
num:
res = MsgBox("The value must be numeric", vbCritical, "Error")
Light_Power = 1
End Sub

Private Sub bmp_to_3dm_Click()

End Sub

Private Sub bmpto_Click()
Form2.Show
Form3.Enabled = False
End Sub

Private Sub Command1_Click()



    comd.Filter = "3D MAP FILE (*.3DM)|*.3DM"
    comd.ShowOpen
    MapFile = comd.FileName
       
Read_3DM (MapFile)
        
Exit Sub
bug:
        
End Sub

Private Sub Command10_Click()
place_mdl
End Sub

Private Sub Command2_Click()
Form5.Show
End Sub

Private Sub Command3_Click()
    

    
    comd.Filter = "3D MODEL FILE (*.mdl)|*.mdl"
    comd.ShowOpen
        
        Mode(Model_Num).Name = FiEnd(comd.FileName)
        READ_MDL (comd.FileName)
        
        
Exit Sub
bug:
        
End Sub

Private Sub Command4_Click()
On Error GoTo bug
comd.Filter = "3D Environment (*.3DT)|*.3DT"
comd.ShowOpen

Open comd.FileName For Input As #7
    
    Line Input #7, stri
    BackFile = stri
    Back_Pict (BackFile)
    
    Line Input #7, stri
    F_Pos.x = stri
    Line Input #7, stri
    F_Pos.y = stri
    Line Input #7, stri
    F_Pos.z = stri
    
    Line Input #7, stri
    Fognum = stri
    Line Input #7, stri
    Fogr = stri
    Line Input #7, stri
    Fogg = stri
    Line Input #7, stri
    Fogb = stri
    Line Input #7, mapf
    Line Input #7, stri
    Light_Power = stri
    Line Input #7, stri
    shap = stri

Line Input #7, num1

For i = 0 To num1 - 1
Line Input #7, File
Set tex(i + tex_num) = Bmap(File, True)
Obj_tex_FILE(i + tex_num) = File
Form1.List1.List(i + tex_num) = "MDL  " & i
Next
tex_num = tex_num + i

Line Input #7, num2
Line Input #7, num2

For i = 0 To num2
Line Input #7, num3
Obj_ver(i).x = num3

Line Input #7, num3
Obj_ver(i).y = num3

Line Input #7, num3
Obj_ver(i).z = num3

Line Input #7, num3
Obj_ver(i).tu = num3

Line Input #7, num3
Obj_ver(i).tv = num3

Line Input #7, num3
Tex_obj(i) = num3 + tex_num - num1

Next
Obj_num = Obj_num + i

Close #7
Call Fog(Fognum, Fogr, Fogg, Fogb)
Get_Pos
'----------------------------------------'

Read_3DM (FiSta(comd.FileName) & mapf)
MapFile = (FiSta(comd.FileName) & mapf)
Exit Sub
bug:

End Sub

Private Sub Command5_Click()
On Error GoTo bug

comd.Filter = "3D Environment (*.3DT)|*.3DT"
res = InputBox("Enter the project name", "Save")
FName = res
file2 = "c:\" & FName
File = FiEnd(MapFile)
MkDir (file2)
FileCopy MapFile, file2 & "\" & FiEnd(MapFile) & ".3DM"


Open file2 & "\" & FName & ".3DT" For Output As #7
    BackgroundF = FiEnd(BackFile) & ".bmp"
    Print #7, BackgroundF
    FileCopy BackFile, file2 & "\" & BackgroundF
    Print #7, F_Pos.x
    Print #7, F_Pos.y
    Print #7, F_Pos.z
    Print #7, Fognum
    Print #7, Fogr
    Print #7, Fogg
    Print #7, Fogb
    Print #7, File & ".3DM"
    Print #7, Light_Power
    Print #7, shap
    
   Print #7, tex_num
For i = 0 To tex_num
Print #7, (Obj_tex_FILE(i))
Next

Print #7, Str(VerNum2 - 1)
For i = 0 To VerNum2 - 1
Print #7, Str(Obj_ver(i).x)
Print #7, Str(Obj_ver(i).y)
Print #7, Str(Obj_ver(i).z)
Print #7, Str(Obj_ver(i).tu)
Print #7, Str(Obj_ver(i).tv)
Print #7, Str(Tex_obj(i))
Next
    
    
Close #7

Exit Sub
bug:
res = MsgBox("Unable to save", vbCritical, "ERROR")
End Sub

Private Sub Command6_Click()
Form4.Show
Form3.Enabled = False
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()

'Change's background
On Error GoTo bug

comd.FileName = ""
comd.Filter = "BMP (*.bmp)|*.bmp"
comd.ShowOpen
Back_Pict (comd.FileName)
BackFile = comd.FileName
Exit Sub
bug:

End Sub

Private Sub Command9_Click()

'Add's a point light to the general light system that will be posed at the camera pos
On Error GoTo bug

Rexyz x, y, z

Add_Light x, y, z, Light.x \ 10, Light.y \ 10, Light.z \ 10, Int(Light.y1), 1
Light_Number = Light_Number + 1
Exit Sub
bug:
res = MsgBox("Sorry system error!")
End Sub

Private Sub ending_Click()

'End's program
End

End Sub

Private Sub fog_prop_Click()

'Show's the fog properries form
Form4.Show
Form3.Enabled = False

End Sub

Private Sub Form_Load()
map_size = 3
'Load's surfaces , math , Z-buffer  , 3DDevice etc..
Call Load_All(Form3.Picture1)
mdlsize = 1
'Light's brightness (begining)
Light_Power = 1

'Fog's thickness
Fognum = 200

'Map sharpness
shap = 0.8

'View as solid
Solid = True

BackFile = App.Path + "\sky.bmp"

End Sub


Private Sub Form_Resize()
Form3.Picture1.Height = Form3.Height - Shape1.Height - 70
Form3.Picture1.Width = Form3.Width
Shape1.Width = Form3.Width
End Sub

Private Sub light_prop_Click()

'open's light properties form
Form6.Show

End Sub

Private Sub map_sizer_Click()
res = InputBox("Enter a number", "Map sharpness")
map_size = res

Exit Sub
num:

res = MsgBox("The value must be numeric", vbCritical, "Error")
map_size = 1
End Sub

Private Sub mdl_sz_Click()
res = InputBox("Enter a number", "Map sharpness")
mdlsize = res

Exit Sub
num:

res = MsgBox("The value must be numeric", vbCritical, "Error")
mdlsize = 1
End Sub

Private Sub mesh_Click()
Form1.Show
Form1.Timer1.Enabled = True
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

'set's key to keycode state
Key = KeyCode
If Shift = 1 Then Extra_Speed = True

Select Case KeyCode
Case vbKeyUp
AKEY(0) = True
Case vbKeyDown
AKEY(1) = True
Case vbKeyQ
AKEY(4) = True
Case vbKeyA
AKEY(5) = True
Case vbKeyRight
AKEY(2) = True
Case vbKeyLeft
AKEY(3) = True
Case vbKeyHome
obj_y = obj_y + 1
Case vbKeyEnd
obj_y = obj_y - 1
Case vbKeyDelete
obj_x = obj_x + 1
Case vbKeyPageDown
obj_x = obj_x - 1
Case vbKey1
obj_z = obj_z + 1
Case vbKey2
obj_z = obj_z - 1
Case vbKeyI
move_mdl 1, 0, 0
Case vbKeyK
move_mdl -1, 0, 0
Case vbKeyJ
move_mdl 0, 0, 1
Case vbKeyL
move_mdl 0, 0, -1
Case vbKeyH
move_mdl 0, -1, 0
Case vbKeyY
move_mdl 0, 1, 0
End Select
Form1.setxyz obj_x, obj_y, obj_z
End Sub


Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp
AKEY(0) = 0
Case vbKeyDown
AKEY(1) = 0
Case vbKeyRight
AKEY(2) = 0
Case vbKeyLeft
AKEY(3) = 0
Case vbKeyQ
AKEY(4) = 0
Case vbKeyA
AKEY(5) = 0
End Select
End Sub

Private Sub plac_Click()
place_mdl
End Sub

Private Sub set_en_Click()

'Set first position to be written at the save file
Set_Pos

End Sub

Private Sub sharp_Click()

'Sets map's sharpness
res = InputBox("Enter a number", "Map sharpness")
shap = res
Read_3DM (MapFile)
Exit Sub
num:

res = MsgBox("The value must be numeric", vbCritical, "Error")
shap = 1

End Sub

Private Sub solidp_Click()

'Set's view to solid
Solid = True

End Sub

Private Sub Timer1_Timer()

'Call's fog
Call Fog(Fognum, Fogr, Fogg, Fogb)

'Call's main loop
Main_Loop (Form3.Picture1)

'Count fps


'Count's vertex ( there is a limited number)

End Sub

Private Sub Timer2_Timer()
RESET_FPS = True
'Print's fps



End Sub

Private Sub wireframe_Click()

'Set's view to WireFrame
Solid = False

End Sub

Function ry(xl, yl, zl)
obj_x = xl
obj_y = yl
obj_z = zl
End Function

