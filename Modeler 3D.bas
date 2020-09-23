Attribute VB_Name = "Module2"
'API function retrive's a pixel from a picture
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'---------------------------------------------------------------------'
'Model's info verexes etx...
Public Obj_tex_FILE(0 To 8000) As String
Public Obj_ver(0 To 8000) As D3DVERTEX
Public Obj_num As Integer
Public Tex_obj(0 To 8000) As Integer
Public tex(0 To 800) As DirectDrawSurface7
Public Type Model
    Models(1 To 2000) As D3DVERTEX
    Name As String
End Type
Public Mode(0 To 100) As Model
Public Model_Num As Integer
Public tex_num As Integer
'------------------------------'
Public MDL_ver(0 To 800) As D3DVERTEX
Public MDL_num As Integer
Public MDL_tex As Integer
Public MDL_text(0 To 800) As Integer
Public MDL_tex_file(0 To 25) As String

'A type used to all sort of things
Type ver3
x As Single
y As Single
z As Single
y1 As Single
y2 As Single
y3 As Single
End Type
'------------------------------'
'Map sharpness
Public shap As Single
'--------------------'
'Fog's info
Public Fogbol As Boolean
Public Fogg As Single
Public Fogr As Single
Public Fogb As Single
Public Fognum As Integer
'---------------------------'
'Save's the first position of the camera
Public F_Pos As ver3
'---------------------------'
'Saves light point properties
Public Light As ver3
Public MapFile As String
Public mdlsize As Single


Sub Ground(vec As D3DVECTOR, y1, y2, y3)
Size = 2 * map_size

ver3 Size + vec.x, vec.y - 1, vec.z + Size, 1, 0
ver3 vec.x - Size, y1 - 1, vec.z + Size, 0, 0
ver3 -Size + vec.x, y2 - 1, vec.z - Size, 0, 1
ver3 vec.x + Size, y3 - 1, vec.z - Size, 1, 1

End Sub


Function vec2(x, y, z) As D3DVECTOR

With vec2
.x = x
.y = y
.z = z
End With

End Function

Sub ver3(x, y, z, xt, yt)

DX.CreateD3DLVertex x, y, z, 1, 1, xt, yt, HMap(VerNum3)
VerNum3 = VerNum3 + 1

End Sub

Sub SAVE_MDL(FileName As String)

'Number of textures used
'Texture files dir
'Number of vertexes
'x
'y
'z
'Texture x
'Texture y
'Texture used with current vertex

Open FileName For Output As #5
Print #5, tex_num
For i = 0 To tex_num
Print #5, (Obj_tex_FILE(i))
Next

Print #5, Str(VerNum2 - 1)
For i = 0 To VerNum2 - 1
Print #5, Str(Vert(i).x)
Print #5, Str(Vert(i).y)
Print #5, Str(Vert(i).z)
Print #5, Str(Vert(i).tu)
Print #5, Str(Vert(i).tv)
Print #5, Str(ver_num(i))
Next
Close #5

End Sub

Sub READ_MDL(FileName As String)
MDL_num = 0

Open FileName For Input As #5
Line Input #5, num1
MDL_tex = num1 - 1

For i = 0 To num1 - 1
Line Input #5, File
MDL_tex_file(i) = File
Next

Line Input #5, num2
Line Input #5, num2

For i = 0 To num2
Line Input #5, num3
MDL_ver(i).x = num3 * mdlsize

Line Input #5, num3
MDL_ver(i).y = num3 * mdlsize

Line Input #5, num3
MDL_ver(i).z = num3 * mdlsize

Line Input #5, num3
MDL_ver(i).tu = num3

Line Input #5, num3
MDL_ver(i).tv = num3

Line Input #5, num3
MDL_text(i) = num3

Next
MDL_num = MDL_num + i
Close #5

End Sub

Sub place_mdl()
NUMI = VerNum2

For i = 0 To MDL_tex
Set tex(i + tex_num) = Bmap(MDL_tex_file(i), True)
Obj_tex_FILE(i + tex_num) = MDL_tex_file(i)

Next
tex_num = tex_num + MDL_tex + 1


For i = 0 To MDL_num - 1
Vert(NUMI + i).x = MDL_ver(i).x
Vert(NUMI + i).y = MDL_ver(i).y
Vert(NUMI + i).z = MDL_ver(i).z
Vert(NUMI + i).tu = MDL_ver(i).tu
Vert(NUMI + i).tv = MDL_ver(i).tv
ver_num(i + NUMI) = MDL_text(i) + tex_num - MDL_tex - 1

Next
VerNum2 = VerNum2 + i
End Sub

Sub move_mdl(x, y, z)
For i = 0 To MDL_num - 1
MDL_ver(i).x = MDL_ver(i).x + x
MDL_ver(i).y = MDL_ver(i).y + y
MDL_ver(i).z = MDL_ver(i).z + z
Next
End Sub

Function zogi(num1) As Boolean
a = num1 / 2
B = num1 / 2
If a = Int(B) Then zogi = True Else zogi = False
End Function

Function FiEnd(FileName As String) As String
    For i = 1 To Len(FileName) - 1
       If Mid(FileName, i, 1) = "\" Then FiEnd = Mid(FileName, i + 1, Len(FileName))
    Next
        LENGTH = Len(FiEnd)
        FiEnd = Mid(FiEnd, 1, LENGTH - 4)
End Function

Function FiSta(FileName As String) As String
    For i = 1 To Len(FileName) - 1
       If Mid(FileName, i, 1) = "\" Then FiSta = Left(FileName, Len(FileName) - Len(FiEnd(FileName)) - 4)
    Next
        
End Function


