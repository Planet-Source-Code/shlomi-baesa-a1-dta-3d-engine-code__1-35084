Attribute VB_Name = "Module3"
Public BackFile As String
Public map_size As Integer

Type Vmap
 v1 As D3DVECTOR
 y1 As Single
 y2 As Single
 y3 As Single
End Type

Public vm(0 To 8000) As Vmap







Sub Read_3DM(File As String)
       
Dim Numo As Integer
       
Form3.Caption = Form3.Caption + "  - Please wait while loading..."

Dim Vec4 As D3DVECTOR
VerNum3 = 0
Close #1
Open File For Input As #1

For i = 1 To (1050 \ 15 * 1)
For t = 1 To (1125 \ 15 * 1)

Line Input #1, xv
Line Input #1, yv
Line Input #1, zv

Vec4.x = xv
Vec4.y = yv
Vec4.z = zv

Line Input #1, y1
Line Input #1, y2
Line Input #1, y3

Vec4.x = Vec4.x * 4 * map_size
Vec4.z = Vec4.z * 4 * map_size

If Vec4.y = 0 And y1 = 0 And y2 = 0 And y3 = 0 Then Tex_Map(NUMI) = 1

'---------------------------------------------------------------'
'---------------------------------------------------------------'
Dim Ms As Single

m = (Vec4.y - y2) \ (-4 * map_size)
c = Vec4.y - m * (Vec4.z - map_size * 2)

m2 = (y1 - y3) \ (-4 * map_size)
c2 = y1 - m2 * (Vec4.z - map_size * 2)

Dim Q As Integer
Dim S As Integer

For Q = (Vec4.z - map_size * 2) To (Vec4.z + map_size * 2)
Dim Yt As Single
Dim Yt2 As Single
Dim Mt As Single
Dim Ct As Single

Yt2 = m * Q + c
Yt = m2 * Q + c2

Mt = (Yt - Yt2) \ (-map_size * 4)
Ct = Yt - Mt * (Vec4.x - map_size * 2)

For S = (Vec4.x - map_size * 2) To (Vec4.x + map_size * 2)

HeMap(S, Q) = (Mt * S + Ct)
'If (Mt * S + Ct) <> 0 Then Form3.Text1.Text = Form3.Text1.Text & "  " & Q & "  " & "  " & S & (Mt * S + Ct)
Next

Next

'---------------------------------------------------------------'
'---------------------------------------------------------------'

NUMI = NUMI + 1
Diff = shap
Vec4.y = Vec4.y * shap
Ground Vec4, y1 * shap, y2 * shap, y3 * shap

Next
Next

Close #1
Form3.Caption = "DTA -  3D  Engine"
Exit Sub
bug:

End Sub


