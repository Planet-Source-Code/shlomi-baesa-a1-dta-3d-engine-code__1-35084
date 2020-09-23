VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BMP to 3DM Tool"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   327682
      Max             =   5
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open BMP"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   1050
      Left            =   120
      Picture         =   "Bmp To 3DM.frx":0000
      ScaleHeight     =   990
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Accuracy"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsize As Integer
Private Sub Command1_Click()

On Error GoTo bug
Size = gsize
Dim verp(0 To 30000) As ver3
Dim Vec4 As D3DVECTOR
Dim NUMI As Integer

x = -1
z = -1
HEG = 8
For i = 1 To Picture2.Height \ (15 * Size)
For t = 1 To Picture2.Width \ (15 * Size)

        col = GetPixel(Picture2.hdc, x + Size, z + Size) \ 1000000
        With Vec4
        .x = x \ Size
        .z = z \ Size
        .y = col * HEG
        End With

        col = GetPixel(Picture2.hdc, x, z + Size) \ 1000000
        y1 = col

        col = GetPixel(Picture2.hdc, x, z) \ 1000000
        y2 = col

        col = GetPixel(Picture2.hdc, x + Size, z) \ 1000000
        y3 = col

    verp(NUMI).x = Vec4.x
    verp(NUMI).y = Vec4.y
    verp(NUMI).z = Vec4.z
    verp(NUMI).y1 = y1 * HEG
    verp(NUMI).y2 = y2 * HEG
    verp(NUMI).y3 = y3 * HEG
    NUMI = NUMI + 1
    x = x + Size

Next

    x = 0
    z = z + Size
    
Next


    Form3.comd.Filter = "3D MAP FILE (*.3DM)|*.3DM"
    Form3.comd.FileName = ""
    Form3.comd.ShowSave
    Open Form3.comd.FileName For Output As #6

For i = 0 To NUMI

Print #6, verp(i).x
Print #6, verp(i).y
Print #6, verp(i).z
Print #6, verp(i).y1
Print #6, verp(i).y2
Print #6, verp(i).y3

Next

Close #6

Exit Sub
bug:

End Sub

Private Sub Command2_Click()
On Error GoTo NOFILE
Form3.comd.Filter = "Bitmap (*.bmp)|*.bmp"
Form3.comd.ShowOpen
Picture2.Picture = LoadPicture(Form3.comd.FileName)
NOFILE:
End Sub

Private Sub Form_Load()
gsize = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Enabled = True
End Sub

Private Sub Slider1_Click()
gsize = Slider1.Value + 1
End Sub
