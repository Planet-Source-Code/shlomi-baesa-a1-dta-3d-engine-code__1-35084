Attribute VB_Name = "Module1"
'-----------------------------------------------------------------------------'
'The dx7 object
Public DX As New DirectX7
'-------------------------'
'The DirectDraw object
Public DDraw As DirectDraw7
'-------------------------'
'Surfaces and thier properties
Dim DDS As DDSURFACEDESC2
'The main surface
Dim Prim As DirectDrawSurface7
'The  background picture
Dim Background As DirectDrawSurface7
Dim gun As DirectDrawSurface7

'The backbuffer surface
Dim Back As DirectDrawSurface7
'--------------------------'
'The Direct3D object
Dim D3 As Direct3D7
'--------------------------'
'The Direct3DDevice object
Public D3D As Direct3DDevice7
'--------------------------'
Dim View(0) As D3DRECT
'--------------------------'
Public Dind As DirectInputDevice
'--------------------------'
'Get's the key state
Public Key As Byte
'-----------------'
'camera position and angle
Dim SideX As Single
Dim SideY As Single
Dim x As Single
Dim z As Single
Dim y As Single
Dim CamY As Single
Dim speed As Single
'---------------'
Public T_GROUND As DirectDrawSurface7
Public T_GROUND2 As DirectDrawSurface7
'--------------------------------------'
'Map information
Public HeMap(-2000 To 2000, -2000 To 2000) As Single
Public VerNum3 As Long
Public HMap(0 To 80000) As D3DLVERTEX
Public Tex_Map(0 To 80000) As Integer
'--------------------------------------'
'Models that were placed
Public VerNum2 As Integer
Public Vert(0 To 80000) As D3DVERTEX
Public Obj_Tex(0 To 80000) As DirectDrawSurface7
'------------------------'
'Camera's extra speed
Public Extra_Speed As Boolean
'------------------------'
'light's sources counter
Public Light_Number As Integer
'-----------------------'
'Amnient light power (x,x,x)
Public Light_Power As Single
'-----------------------'
'controls the view properties (WireFrame or solid)
Public Solid As Boolean
'-----------------------'
Dim wa(0 To 4) As DirectDrawSurface7
Dim watern As Integer
Dim FPS As Integer
Public RESET_FPS As Boolean

Dim Sky_Box(0 To 16) As D3DVERTEX
Dim Sky_Tex As DirectDrawSurface7
Dim Sky_Rot As Single

Public ver_tex(0 To 8000) As DirectDrawSurface7
Public ver_num(0 To 8000) As Integer
Public AKEY(0 To 5) As Boolean
Dim LastY As Single
Dim upp As Boolean
Public usedt As Boolean


Sub Start_DDraw(Your_Form, FILEBACK As String)

'Setting the DirectDraw object
Set DDraw = DX.DirectDrawCreate("")
DDraw.SetCooperativeLevel Your_Form.hWnd, DDSCL_NORMAL

'Setting the primary surface
With DDS
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_VIDEOMEMORY Or DDSCAPS_PRIMARYSURFACE
End With
Set Prim = DDraw.CreateSurface(DDS)

'Setting the backbuffer surface
With DDS
    .lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_VIDEOMEMORY Or DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    .lHeight = 500
    .lWidth = 500
End With
Set Back = DDraw.CreateSurface(DDS)

'Setting the background surface
With DDS
    .lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_VIDEOMEMORY Or DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    .lHeight = 500
    .lWidth = 100 * size0
End With
Set Background = DDraw.CreateSurfaceFromFile(FILEBACK, DDS)






End Sub

Sub Start_D3D()

Dim VIEWT As D3DVIEWPORT7
Dim Mat As D3DMATERIAL7

'Set's the Direct3DDevice
Set D3D = D3.CreateDevice("IID_IDirect3DHALDevice", Back)

'Set's the viewing properties
With VIEWT
    .lHeight = 500
    .lWidth = 500
    .maxz = 1
    .minz = 0
End With

D3D.SetViewport VIEWT

With View(0)
    .X1 = 0
    .y1 = 0
    .y2 = 500
    .X2 = 500
End With

'Set's the material
With Mat
    .Ambient.a = 1
    .Ambient.G = 1
    .Ambient.R = 1
    .Ambient.B = 1
End With
D3D.SetMaterial Mat

End Sub

Sub Do_Math()
'All of it is math for camera world etc...

Dim vec(0 To 3) As D3DVECTOR
Dim MW As D3DMATRIX
Dim MV As D3DMATRIX
Dim MP As D3DMATRIX
Dim Ms As D3DMATRIX

'--------------------'
DX.IdentityMatrix MW
D3D.SetTransform D3DTRANSFORMSTATE_WORLD, MW

'--------------------------------------------'
vec(0).x = 0
vec(0).y = 0
vec(0).z = 6

vec(1).x = 0
vec(1).y = 0
vec(1).z = 0

vec(2).x = 0
vec(2).y = 5
vec(2).z = 0

DX.IdentityMatrix MV
DX.ViewMatrix MV, vec(0), vec(1), vec(2), 0
D3D.SetTransform D3DTRANSFORMSTATE_VIEW, MV

'------------------------------------------------'
DX.IdentityMatrix MP
DX.ProjectionMatrix MP, 1, 32000, 3.14 \ 2
D3D.SetTransform D3DTRANSFORMSTATE_PROJECTION, MP

End Sub


Sub Depth_Buffer()

'The z_buffer surface
Dim S_ZBUFFER As DirectDrawSurface7

Dim pixf As DDPIXELFORMAT
Dim pixf2 As Direct3DEnumPixelFormats
Dim ddsu2 As DDSURFACEDESC2

'Set's the 3D object
Set D3 = DDraw.GetDirect3D

'The z-buffer
Set pixf2 = D3.GetEnumZBufferFormats("IID_IDirect3DHALDevice")

Dim i As Long

For i = 1 To pixf2.GetCount()
    pixf2.GetItem i, pixf
    If pixf.lFlags = DDPF_ZBUFFER Then
    Exit For
    End If
Next i

With ddsu2
    .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
    .ddsCaps.lCaps = DDSCAPS_ZBUFFER
    .lWidth = 500
    .lHeight = 500
    .ddpfPixelFormat = pixf
    .ddsCaps.lCaps = ddsu2.ddsCaps.lCaps Or DDSCAPS_VIDEOMEMORY
End With

Set S_ZBUFFER = DDraw.CreateSurface(ddsu2)

'Attach Z-buffer to other surfaces
Back.AddAttachedSurface S_ZBUFFER

End Sub

Sub Render_State()

'Rendering stats (light , fog , shade etc...)

'Ambiet light
D3D.SetRenderState D3DRENDERSTATE_AMBIENT, DX.CreateColorRGB(Light_Power, Light_Power, Light_Power)

'Fog properties
D3D.SetRenderState D3DRENDERSTATE_FOGVERTEXMODE, 2

'culling mode
D3D.SetRenderState D3DRENDERSTATE_CULLMODE, 0

'Shade mode
D3D.SetRenderState D3DRENDERSTATE_SHADEMODE, 0

'Colorkeying enabling


'The fog call
Call Fog(Fognum, Fogr, Fogg, Fogb)

End Sub

Sub Load_All(Your_Form)
upp = False
'Load's surfaces , math , Z-buffer  , 3DDevice etc..
Start_DDraw Your_Form, App.Path + "\sky.bmp"
Depth_Buffer
Start_D3D
Render_State
Do_Math
SkyBox "dsf"

Set T_GROUND = Bmap(App.Path & "\" & "moss05.BMP", False)
Set T_GROUND2 = Bmap(App.Path & "\" & "moss05.BMP", True)

Dim CLIPPER As DirectDrawClipper
Dim DRect As RECT

'The clipper uses to keep the erndering inside our frame
DX.GetWindowRect Your_Form.hWnd, DRect
Set CLIPPER = DDraw.CreateClipper(0)
CLIPPER.SetHWnd Your_Form.hWnd
Prim.SetClipper CLIPPER


End Sub


Sub Main_Loop(Your_Form)
Do While 0 = 0


Dim DRect As RECT
Dim sRect As RECT

FPS = FPS + 1
If RESET_FPS = True Then
Form3.sb1.SimpleText = "  The FPS Clock -  " & FPS & "  |  " _
& "  Nmuber of Vertex -  " & VerNum2 & "  |  " & "  Textures Number -  " & tex_num
FPS = 0
RESET_FPS = False
End If

'Clear's the surfaces
D3D.Clear 1, View(), D3DCLEAR_ZBUFFER Or D3DCLEAR_TARGET, vbBlack, 1, 0

'Call's the render state
Render_State

'Call's camera math
Call camera_move(Key)

'Call's the rendering functions
Render_Secne

'Blt the the backbuffer to the main surface
DX.GetWindowRect Form3.Picture1.hWnd, DRect

Prim.Blt DRect, Back, sRect, DDBLT_WAIT

DoEvents

Loop
End Sub

Function Bmap(FileName, EnableCKey As Boolean) As DirectDrawSurface7

'This function set's a texture to a surface with a colorkey
'The file loaded to the surface must be a squere bitmap ( for example : 256 * 256 pixel bitmap)

Dim dsd As DDSURFACEDESC2
Dim pict As DirectDrawSurface7
Dim kind As Direct3DEnumPixelFormats
Dim CKey As DDCOLORKEY


Set kind = D3D.GetTextureFormatsEnum
kind.GetItem 1, dsd.ddpfPixelFormat

With dsd
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_PIXELFORMAT Or DDSD_TEXTURESTAGE Or DDSD_CKSRCBLT
    .ddsCaps.lCaps = DDSCAPS_TEXTURE
    .ddsCaps.lCaps2 = DDSCAPS2_TEXTUREMANAGE
    .lTextureStage = 50

End With

Set pict = DDraw.CreateSurfaceFromFile(FileName, dsd)

'The colorkey to be used
With CKey
.high = vbWhite
.low = vbWhite
End With

'Color keying the texture
If EnableCKey = True Then pict.SetColorKey DDCKEY_SRCBLT, CKey

Set Bmap = pict

End Function

Sub Fog(Thickness As Integer, R, G, B)

'Enables fog

D3D.SetRenderState D3DRENDERSTATE_FOGENABLE, True

'Fog color
D3D.SetRenderState D3DRENDERSTATE_FOGCOLOR, DX.CreateColorRGB(R, G, B)

D3D.SetRenderState D3DRENDERSTATE_FOGVERTEXMODE, D3DFOG_LINEAR

D3D.SetRenderStateSingle D3DRENDERSTATE_FOGSTART, 0

D3D.SetRenderStateSingle D3DRENDERSTATE_FOGEND, Thickness

End Sub

Sub camera_move(Key As Byte)

'Camera moverment rotation around the Y and X axis

Dim mside As D3DMATRIX
Dim mxside As D3DMATRIX
Dim myside As D3DMATRIX
Dim mloc As D3DMATRIX
Dim mvie As D3DMATRIX

DX.IdentityMatrix mxside
DX.IdentityMatrix myside
DX.IdentityMatrix mside
DX.IdentityMatrix mvie
DX.IdentityMatrix mloc


If Extra_Speed = True Then
speed = speed + 0.3
Extra_Speed = False
End If
If AKEY(2) = True Then SideX = SideX + 0.05

If AKEY(3) = True Then SideX = SideX - 0.05

If AKEY(4) = True Then SideY = SideY + 0.05

If AKEY(5) = True Then SideY = SideY - 0.05

Begin:
'---------------------------------------------------------'

If AKEY(0) = True Then
    z = z - Cos(SideX) * (0.5 + speed)
    x = x - Sin(SideX) * (0.5 + speed)
End If
'------------------------------------------------------------'
If AKEY(1) = True Then
    z = z + Cos(SideX) * (0.5 + speed)
    x = x + Sin(SideX) * (0.5 + speed)
End If
'---------------------------------------------'

DX.RotateYMatrix mxside, SideX
DX.MatrixMultiply mside, mxside, myside
DX.RotateXMatrix myside, SideY
DX.MatrixMultiply mside, mxside, myside


'Get's height of current point on map

xm = (Int(x \ (map_size * 4))) * map_size * 4
zm = (Int(z \ (map_size * 4))) * map_size * 4
'--------------------------------------------------------------'
Dim Ms As Single

y1 = HeMap(-xm + map_size * 2, -zm + map_size * 2)
y = HeMap(-xm - map_size * 2, -zm + map_size * 2)
y3 = HeMap(-xm + map_size * 2, -zm - map_size * 2)
y2 = HeMap(-xm - map_size * 2, -zm - map_size * 2)

m = (y - y2) \ (-4 * map_size)
c = y - m * (zm - map_size * 2)

m2 = (y1 - y3) \ (-4 * map_size)
c2 = y1 - m2 * (zm - map_size * 2)

Dim Q As Single
Dim S As Single

Dim Yt As Single
Dim Yt2 As Single
Dim Mt As Single
Dim Ct As Single

Yt2 = m * z + c
Yt = m2 * z + c2

Mt = (Yt - Yt2) \ (-map_size * 4)
Ct = Yt - Mt * (xm - map_size * 2)

y = (Mt * x + Ct)


'-------------------------------------------------------------------'



'Set's the X Y Z position
mloc.rc41 = x
mloc.rc42 = -y * shap - 15 - CamY
mloc.rc43 = z

DX.MatrixMultiply mvie, mloc, mside
D3D.SetTransform D3DTRANSFORMSTATE_VIEW, mvie

'Reset's the key state
Key = 0


End Sub


Sub Load_Midi_Music(FileName As String, Your_Form As Form)

'Load's a midi file and play's it

Dim musicl As DirectMusicLoader
Dim musicp As DirectMusicPerformance

Set musicl = DX.DirectMusicLoaderCreate
Set musicp = DX.DirectMusicPerformanceCreate
Call musicp.Init(Nothing, Form.hWnd)
Call musicp.SetPort(-1, 1)
Dim musnam As DirectMusicSegment
Set musnam = musicl.LoadSegment(FileName)
Call musnam.SetStandardMidiFile
musicp.SetMasterAutoDownload (True)
Call musnam.Download(musicp)
Call musicp.PlaySegment(musnam, 0, 0)

End Sub

Sub Add_Light(x, y, z, R, G, B, Light_Range, ADD As Boolean)

On Error GoTo bug

'Add's a point light

Dim Light As D3DLIGHT7
Dim color As D3DCOLORVALUE

'Light's color
With color
    .B = B
    .G = G
    .R = R
    .a = 1
End With

'Light's attenuation , range  ,pos etc...
With Light
    .dltType = D3DLIGHT_POINT
    .position.x = x
    .position.y = y
    .position.z = z
    .specular.a = 1
    .specular.R = 1
    .specular.B = 1
    .attenuation1 = 0.1
    .attenuation0 = 1
    .attenuation2 = 0
    .diffuse = color
    .specular = color
    .Ambient = color
    .range = Light_Range
End With

D3D.SetLight Light_Number, Light
D3D.LightEnable Light_Number, True

'Add light to general light system
If ADD = True Then Light_Number = Light_Number + 1

Exit Sub
bug:
res = MsgBox("The problem was in the Add_light function", vbCritical, "Error")
End

End Sub

Function Vec1(value1, value2, value3) As D3DVECTOR

Vec1.x = value1
Vec1.y = value2
Vec1.z = value3

End Function

Sub Render_Secne()

'render's secene
Dim Mat  As D3DMATERIAL7

D3D.BeginScene

'draw's secene as WireFrame or Solid
If Solid = True Then
    'Rotate the sky box
    Dim MW As D3DMATRIX
    Dim MPo As D3DMATRIX
    Dim MG As D3DMATRIX
    
    DX.IdentityMatrix MG
    DX.IdentityMatrix MPo
    DX.IdentityMatrix MW
    
    Sky_Rot = Sky_Rot + 0.0009
    MPo.rc41 = 30
    MPo.rc43 = 30
    
    DX.RotateYMatrix MW, Sky_Rot
    DX.MatrixMultiply MG, MW, MPo
    D3D.SetTransform D3DTRANSFORMSTATE_WORLD, MG
    
    With Mat
        .Ambient.a = 1
        .Ambient.G = 1
        .Ambient.R = 1
        .Ambient.B = 1
    End With
    D3D.SetMaterial Mat
    
    D3D.SetRenderState D3DRENDERSTATE_FOGENABLE, False
    D3D.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, 0
    
    'Draw the sky box
    D3D.SetTexture 0, Sky_Tex
    
    For i = 0 To 3
    D3D.DrawPrimitive 6, D3DFVF_VERTEX, Sky_Box(NUMI), 4, D3DDP_WAIT
    NUMI = NUMI + 4
    Next
    NUMI = 0
    
    D3D.SetRenderState D3DRENDERSTATE_FOGENABLE, 1

    'Return world to normal after rotating sky box
    DX.RotateYMatrix MW, 0
    D3D.SetTransform D3DTRANSFORMSTATE_WORLD, MW
    ran = 6
    'Render map
    For i = 0 To (VerNum3 / 4)
    
        If Tex_Map(i) <> 1 Then
        D3D.SetTexture 0, T_GROUND
        Else
        D3D.SetTexture 0, T_GROUND2
        End If
        
      
    If ran = 6 Then
    ran = 8
    GoTo drawl3
    End If
    If ran = 8 Then ran = 6
drawl3:
    
        With Mat
            .Ambient.a = 1
            .Ambient.G = ran
            .Ambient.R = ran
            .Ambient.B = ran
        End With

        D3D.SetMaterial Mat

        D3D.DrawPrimitive 6, D3DFVF_VERTEX, HMap(NUMI), 4, D3DDP_WAIT
        NUMI = NUMI + 4
    
    Next
        
        NUMI = 0

ran = 0.5
    
    D3D.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, 1
    'Render unplaced models
    For i = 0 To 24
    
    If usedt = True Then D3D.SetTexture 0, tex(Form1.List1.ListIndex)
    D3D.DrawPrimitive 6, D3DFVF_VERTEX, Obj_ver(NUMI), 4, 0
    NUMI = NUMI + 4
    
    Next
    
    NUMI = 0

    For i = 0 To (MDL_num / 4)
    
    D3D.DrawPrimitive 3, D3DFVF_VERTEX, MDL_ver(NUMI), 4, 0
    NUMI = NUMI + 4
    
    Next
    
    NUMI = 0

ran = 0.4

    For i = 0 To (VerNum2 / 4)
    D3D.SetTexture 0, tex(ver_num(NUMI))
    
    If ran = 0.4 Then
    ran = 1.2
    GoTo drawl
    End If
    If ran = 1.2 Then ran = 0.4
drawl:
    
    With Mat
        .Ambient.a = 1
        .Ambient.G = ran
        .Ambient.R = ran
        .Ambient.B = ran
    End With
    D3D.SetMaterial Mat
        
    D3D.DrawPrimitive 6, D3DFVF_VERTEX, Vert(NUMI), 4, 0
    NUMI = NUMI + 4
    
    Next
    
    NUMI = 0
'----------------------------------------------------------------------'
'----------------------------------------------------------------------'
Else
    
    D3D.SetTexture 0, Sky_Tex
    For i = 0 To 4
    D3D.DrawPrimitive 3, D3DFVF_VERTEX, Sky_Box(NUMI), 4, D3DDP_WAIT
    NUMI = NUMI + 4
    Next
    NUMI = 0
    
    For i = 0 To (VerNum3 / 4)
    
    If Tex_Map(NUMI) = 1 Then
    D3D.SetTexture 0, T_GROUND
    Else
    D3D.SetTexture 0, T_GROUND2
    End If

    D3D.DrawPrimitive 3, D3DFVF_VERTEX, HMap(NUMI), 4, 0
    NUMI = NUMI + 4
    
    Next
    
    NUMI = 0


    For i = 0 To (VerNum2 / 4) + (Obj_num / 4)
    D3D.SetTexture 0, tex(Tex_obj(NUMI))
    D3D.DrawPrimitive 3, D3DFVF_VERTEX, Obj_ver(NUMI), 4, 0
    NUMI = NUMI + 4
    
    Next
    
    NUMI = 0

End If

D3D.EndScene

End Sub

Sub ver1(x, y, z, tx, ty)

DX.CreateD3DVertex x, y, z, 0, 0, 0, tx, ty, Vert(VerNum2)
VerNum2 = VerNum2 + 1

End Sub



Sub play_wave()

'play's wave file
sound_buff.Play (DSBPLAY_DEFAULT)

End Sub


Sub Load_Midi(FileName As String, Your_Form As Form)

'Load's midi music

Dim musicl As DirectMusicLoader
Dim musicp As DirectMusicPerformance

Set musicl = DX.DirectMusicLoaderCreate
Set musicp = DX.DirectMusicPerformanceCreate
Call musicp.Init(Nothing, Form.hWnd)
Call musicp.SetPort(-1, 1)
Dim musnam As DirectMusicSegment
Set musnam = musicl.LoadSegment(FileName)
Call musnam.SetStandardMidiFile
musicp.SetMasterAutoDownload (True)
Call musnam.Download(musicp)
Call musicp.PlaySegment(musnam, 0, 0)

End Sub

Sub Set_Pos()
F_Pos.x = x
F_Pos.y = y
F_Pos.z = z
End Sub

Sub Get_Pos()
x = F_Pos.x
y = F_Pos.y
z = F_Pos.z
End Sub


Function Rexyz(cX, cY, cZ)
cX = -x
cZ = -z
cY = -y

End Function

Sub Back_Pict(File As String)

Set Sky_Tex = Bmap(File, False)

End Sub

Sub water()
Dim srec As RECT
Dim drec As RECT

watern = watern + 1
T_GROUND2.Blt drec, wa(watern), srec, DDBLT_WAIT

If watern = 4 Then watern = 0

End Sub


Sub SkyBox(File)
Size = 100
Set Sky_Tex = Bmap(App.Path + "\sky.bmp", False)

DX.CreateD3DVertex 100 * Size, 100 * Size, 100 * Size, 0, 0, 0, 1, 1, Sky_Box(0)
DX.CreateD3DVertex -100 * Size, 100 * Size, 100 * Size, 0, 0, 0, 0, 1, Sky_Box(1)
DX.CreateD3DVertex -100 * Size, -100, 100 * Size, 0, 0, 0, 0, 0, Sky_Box(2)
DX.CreateD3DVertex 100 * Size, -100, 100 * Size, 0, 0, 0, 1, 0, Sky_Box(3)

DX.CreateD3DVertex 100 * Size, 100 * Size, -100 * Size, 0, 0, 0, 1, 1, Sky_Box(4)
DX.CreateD3DVertex -100 * Size, 100 * Size, -100 * Size, 0, 0, 0, 0, 1, Sky_Box(5)
DX.CreateD3DVertex -100 * Size, -100, -100 * Size, 0, 0, 0, 0, 0, Sky_Box(6)
DX.CreateD3DVertex 100 * Size, -100, -100 * Size, 0, 0, 0, 1, 0, Sky_Box(7)

DX.CreateD3DVertex 100 * Size, 100 * Size, 100 * Size, 0, 0, 0, 1, 1, Sky_Box(8)
DX.CreateD3DVertex 100 * Size, 100 * Size, -100 * Size, 0, 0, 0, 0, 1, Sky_Box(9)
DX.CreateD3DVertex 100 * Size, -100, -100 * Size, 0, 0, 0, 0, 0, Sky_Box(10)
DX.CreateD3DVertex 100 * Size, -100, 100 * Size, 0, 0, 0, 1, 0, Sky_Box(11)

DX.CreateD3DVertex -100 * Size, 100 * Size, 100 * Size, 0, 0, 0, 1, 1, Sky_Box(12)
DX.CreateD3DVertex -100 * Size, 100 * Size, -100 * Size, 0, 0, 0, 0, 1, Sky_Box(13)
DX.CreateD3DVertex -100 * Size, -100, -100 * Size, 0, 0, 0, 0, 0, Sky_Box(14)
DX.CreateD3DVertex -100 * Size, -100, 100 * Size, 0, 0, 0, 1, 0, Sky_Box(15)

End Sub

