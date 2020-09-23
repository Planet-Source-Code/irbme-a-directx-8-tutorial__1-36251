Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


'Custom vertex type
Public Type UNLITVERTEX
    X    As Single          'X coordinate
    Y    As Single          'Y Coordinate
    Z    As Single          'Z Coordinate
    nx   As Single          'Normal vector X Coordinate
    ny   As Single          'Normal vector Y Coordinate
    nz   As Single          'Normal vector Z Coordinate
    tu   As Single          'Texture X Coordinate
    tv   As Single          'Texture Y Coordinate
End Type


'Flexible vertex format descriptor
Public Const UNLIT_FVF = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

Public Const PI As Single = 3.14159265358979 'PI
Public Const Rad = PI / 180                  'Degrees to Radians convertor


Public Dx           As DirectX8               'DirectX
Public D3D          As Direct3D8              'Direct3D
Public D3DDevice    As Direct3DDevice8        'Hardware device
Public D3DX         As D3DX8                  'Helper library
Public VBuffer      As Direct3DVertexBuffer8  'Vertex buffer to store geometry


Public MainFont     As D3DXFont     'Main font object
Public MainFontDesc As IFont        'Main font descriptor
Public TextRect     As RECT         'Text rect
Public fnt          As New StdFont  'Font object to hold descriptor


Public matWorld     As D3DMATRIX    'World matrix
Public matView      As D3DMATRIX    'View matrix
Public matProj      As D3DMATRIX    'Projection matrix


Public Textures(0 To 4) As Direct3DTexture8  'Textures
Public CurrentTexture As Integer

Public FPS_LastCheck   As Long      'Last FPS count
Public FPS_Count       As Long      'FPS
Public FPS_Current     As Integer   'FPS Counter


Public Lights(0 To 2)       As D3DLIGHT8 '3 Lights
Public Light1On             As Boolean
Public Light2On             As Boolean
Public Light3On             As Boolean


Public bRunning   As Boolean      'Loop terminator
Public Cube(35)   As UNLITVERTEX  'Vertex array


Public RotateAngleX As Single      'X rotation angle
Public RotateAngleY As Single      'Y rotation angle
Public ZoomFactor   As Single      'Zoom Factor




Public Function Initialise(Adapter As Long, Device As CONST_D3DDEVTYPE, Width As Long, Height As Long, Format As CONST_D3DFORMAT) As Boolean

  Dim DispMode As D3DDISPLAYMODE
  Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    DispMode.Format = Format
    DispMode.Width = Width
    DispMode.Height = Height
    
    D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP       'Use a backbuffer
    D3DWindow.BackBufferCount = 1                   '1 backbuffer
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.BackBufferWidth = DispMode.Width
    D3DWindow.BackBufferHeight = DispMode.Height
    D3DWindow.hDeviceWindow = fMain.hWnd
    D3DWindow.EnableAutoDepthStencil = 1            'Enable auto depth stencil
    
    'Check for 16 bit depth stencil
    If D3D.CheckDeviceFormat(Adapter, Device, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
    Else
        MsgBox "The 16 bit auto depth stencil is not supported by your graphics card. Unable to continue..."
        GoTo ErrHandler
    End If
    
    'Create hardware device
    On Error Resume Next
      Err.Number = 0
      Set D3DDevice = D3D.CreateDevice(Adapter, Device, fMain.hWnd, D3DCREATE_PUREDEVICE, D3DWindow)
      If Err.Number <> 0 Then
          Err.Number = 0
          Set D3DDevice = D3D.CreateDevice(Adapter, Device, fMain.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)
          If Err.Number <> 0 Then
              Err.Number = 0
              Set D3DDevice = D3D.CreateDevice(Adapter, Device, fMain.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, D3DWindow)
              If Err.Number <> 0 Then
                  Err.Number = 0
                  Set D3DDevice = D3D.CreateDevice(Adapter, Device, fMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
              End If
          End If
      End If
      
    On Error GoTo ErrHandler
      
    'Set device states
    D3DDevice.SetVertexShader UNLIT_FVF
    D3DDevice.SetRenderState D3DRS_LIGHTING, 1
    D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    D3DDevice.SetRenderState D3DRS_AMBIENT, &H202020
    
    'Setup world matrix
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    'Setup view matrix
    D3DXMatrixLookAtLH matView, MakeVector(0, 5, 9), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
    D3DDevice.SetTransform D3DTS_VIEW, matView
    
    'Setup projection matrix
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 1, 0.1, 500
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    
    'Create font
    fnt.Name = "Verdana"
    fnt.Size = 8
    Set MainFontDesc = fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    
    'Load textures
    Set Textures(0) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Texture0.bmp", 128, 128, D3DX_DEFAULT, 0, DispMode.Format, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Set Textures(1) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Texture1.bmp", 128, 128, D3DX_DEFAULT, 0, DispMode.Format, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Set Textures(2) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Texture2.bmp", 128, 128, D3DX_DEFAULT, 0, DispMode.Format, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Set Textures(3) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Texture3.bmp", 128, 128, D3DX_DEFAULT, 0, DispMode.Format, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Set Textures(4) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Texture4.bmp", 128, 128, D3DX_DEFAULT, 0, DispMode.Format, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    
    If Not InitialiseGeometry Then
        GoTo ErrHandler
    End If
    
    If Not SetupLights Then
        GoTo ErrHandler
    End If
    
    Initialise = True
    
Exit Function
ErrHandler:
    Initialise = False
    
End Function





Private Function CreateVertex(X As Single, Y As Single, Z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As UNLITVERTEX
    
    With CreateVertex
        .X = X: .Y = Y: .Z = Z
        .nx = nx: .ny = ny: .nz = nz
        .tu = tu: .tv = tv:
    End With
    
End Function


Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z
End Function


Private Function GenerateTriangleNormals(p0 As UNLITVERTEX, p1 As UNLITVERTEX, p2 As UNLITVERTEX) As D3DVECTOR

    Dim v01 As D3DVECTOR   'Vector from points 0 to 1
    Dim v02 As D3DVECTOR   'Vector from points 0 to 2
    Dim vNorm As D3DVECTOR 'Final vector

    'Create vectors from points 0 to 1 and 0 to 2
    D3DXVec3Subtract v01, MakeVector(p1.X, p1.Y, p1.Z), MakeVector(p0.X, p0.Y, p0.Z)
    D3DXVec3Subtract v02, MakeVector(p2.X, p2.Y, p2.Z), MakeVector(p0.X, p0.Y, p0.Z)

    'Get cross product
    D3DXVec3Cross vNorm, v01, v02

    'Normalize vector
    D3DXVec3Normalize vNorm, vNorm

    'Return the value
    GenerateTriangleNormals.X = vNorm.X
    GenerateTriangleNormals.Y = vNorm.Y
    GenerateTriangleNormals.Z = vNorm.Z
    
End Function



Private Function InitialiseGeometry() As Boolean
    
  Dim vN As D3DVECTOR       'Triangle normal
   
    On Error GoTo ErrHandler

    'Front
    Cube(0) = CreateVertex(-1, -1, 1, 0, 0, 0, 0, 0)
    Cube(1) = CreateVertex(1, 1, 1, 0, 0, 0, 1, 1)
    Cube(2) = CreateVertex(-1, 1, 1, 0, 0, 0, 0, 1)
        
    vN = GenerateTriangleNormals(Cube(0), Cube(1), Cube(2))
        
    Cube(0).nx = vN.X: Cube(0).ny = vN.Y: Cube(0).nz = vN.Z
    Cube(1).nx = vN.X: Cube(1).ny = vN.Y: Cube(1).nz = vN.Z
    Cube(2).nx = vN.X: Cube(2).ny = vN.Y: Cube(2).nz = vN.Z
        
        
    Cube(3) = CreateVertex(1, 1, 1, 0, 0, 0, 1, 1)
    Cube(4) = CreateVertex(-1, -1, 1, 0, 0, 0, 0, 0)
    Cube(5) = CreateVertex(1, -1, 1, 0, 0, 0, 1, 0)
        
    vN = GenerateTriangleNormals(Cube(3), Cube(4), Cube(5))
        
    Cube(3).nx = vN.X: Cube(3).ny = vN.Y: Cube(3).nz = vN.Z
    Cube(4).nx = vN.X: Cube(4).ny = vN.Y: Cube(4).nz = vN.Z
    Cube(5).nx = vN.X: Cube(5).ny = vN.Y: Cube(5).nz = vN.Z
        
    'Back
    Cube(6) = CreateVertex(-1, 1, -1, 0, 0, 0, 0, 1)
    Cube(7) = CreateVertex(1, 1, -1, 0, 0, 0, 1, 1)
    Cube(8) = CreateVertex(-1, -1, -1, 0, 0, 0, 0, 0)
        
    vN = GenerateTriangleNormals(Cube(6), Cube(7), Cube(8))
            
    Cube(6).nx = vN.X: Cube(6).ny = vN.Y: Cube(6).nz = vN.Z
    Cube(7).nx = vN.X: Cube(7).ny = vN.Y: Cube(7).nz = vN.Z
    Cube(8).nx = vN.X: Cube(8).ny = vN.Y: Cube(8).nz = vN.Z
        
        
    Cube(9) = CreateVertex(1, -1, -1, 0, 0, 0, 1, 0)
    Cube(10) = CreateVertex(-1, -1, -1, 0, 0, 0, 0, 0)
    Cube(11) = CreateVertex(1, 1, -1, 0, 0, 0, 1, 1)
        
    vN = GenerateTriangleNormals(Cube(9), Cube(10), Cube(11))
            
    Cube(9).nx = vN.X: Cube(9).ny = vN.Y: Cube(9).nz = vN.Z
    Cube(10).nx = vN.X: Cube(10).ny = vN.Y: Cube(10).nz = vN.Z
    Cube(11).nx = vN.X: Cube(11).ny = vN.Y: Cube(11).nz = vN.Z
        
    'Right
    Cube(12) = CreateVertex(-1, -1, -1, 0, 0, 0, 0, 0)
    Cube(13) = CreateVertex(-1, 1, 1, 0, 0, 0, 1, 1)
    Cube(14) = CreateVertex(-1, 1, -1, 0, 0, 0, 1, 0)
        
    vN = GenerateTriangleNormals(Cube(12), Cube(13), Cube(14))
            
    Cube(12).nx = vN.X: Cube(12).ny = vN.Y: Cube(12).nz = vN.Z
    Cube(13).nx = vN.X: Cube(13).ny = vN.Y: Cube(13).nz = vN.Z
    Cube(14).nx = vN.X: Cube(14).ny = vN.Y: Cube(14).nz = vN.Z
        
        
    Cube(15) = CreateVertex(-1, 1, 1, 0, 0, 0, 1, 1)
    Cube(16) = CreateVertex(-1, -1, -1, 0, 0, 0, 0, 0)
    Cube(17) = CreateVertex(-1, -1, 1, 0, 0, 0, 0, 1)
        
    vN = GenerateTriangleNormals(Cube(15), Cube(16), Cube(17))
        
    Cube(15).nx = vN.X: Cube(15).ny = vN.Y: Cube(15).nz = vN.Z
    Cube(16).nx = vN.X: Cube(16).ny = vN.Y: Cube(16).nz = vN.Z
    Cube(17).nx = vN.X: Cube(17).ny = vN.Y: Cube(17).nz = vN.Z
        
    'Left
    Cube(18) = CreateVertex(1, 1, -1, 0, 0, 0, 1, 0)
    Cube(19) = CreateVertex(1, 1, 1, 0, 0, 0, 1, 1)
    Cube(20) = CreateVertex(1, -1, -1, 0, 0, 0, 0, 0)
        
    vN = GenerateTriangleNormals(Cube(18), Cube(19), Cube(20))
            
    Cube(18).nx = vN.X: Cube(18).ny = vN.Y: Cube(18).nz = vN.Z
    Cube(19).nx = vN.X: Cube(19).ny = vN.Y: Cube(19).nz = vN.Z
    Cube(20).nx = vN.X: Cube(20).ny = vN.Y: Cube(20).nz = vN.Z
        
        
    Cube(21) = CreateVertex(1, -1, 1, 0, 0, 0, 0, 1)
    Cube(22) = CreateVertex(1, -1, -1, 0, 0, 0, 0, 0)
    Cube(23) = CreateVertex(1, 1, 1, 0, 0, 0, 1, 1)
        
    vN = GenerateTriangleNormals(Cube(21), Cube(22), Cube(23))
            
    Cube(21).nx = vN.X: Cube(21).ny = vN.Y: Cube(21).nz = vN.Z
    Cube(22).nx = vN.X: Cube(22).ny = vN.Y: Cube(22).nz = vN.Z
    Cube(23).nx = vN.X: Cube(23).ny = vN.Y: Cube(23).nz = vN.Z
        
        
    'Top
    Cube(24) = CreateVertex(-1, 1, 1, 0, 0, 0, 0, 1)
    Cube(25) = CreateVertex(1, 1, 1, 0, 0, 0, 1, 1)
    Cube(26) = CreateVertex(-1, 1, -1, 0, 0, 0, 0, 0)
        
    vN = GenerateTriangleNormals(Cube(24), Cube(25), Cube(26))
            
    Cube(24).nx = vN.X: Cube(24).ny = vN.Y: Cube(24).nz = vN.Z
    Cube(25).nx = vN.X: Cube(25).ny = vN.Y: Cube(25).nz = vN.Z
    Cube(26).nx = vN.X: Cube(26).ny = vN.Y: Cube(26).nz = vN.Z
        
        
    Cube(27) = CreateVertex(1, 1, -1, 0, 0, 0, 1, 0)
    Cube(28) = CreateVertex(-1, 1, -1, 0, 0, 0, 0, 0)
    Cube(29) = CreateVertex(1, 1, 1, 0, 0, 0, 1, 1)
        
    vN = GenerateTriangleNormals(Cube(27), Cube(28), Cube(29))
            
    Cube(27).nx = vN.X: Cube(27).ny = vN.Y: Cube(27).nz = vN.Z
    Cube(28).nx = vN.X: Cube(28).ny = vN.Y: Cube(28).nz = vN.Z
    Cube(29).nx = vN.X: Cube(29).ny = vN.Y: Cube(29).nz = vN.Z
        
    'Bottom
    Cube(30) = CreateVertex(-1, -1, -1, 0, 0, 0, 0, 0)
    Cube(31) = CreateVertex(1, -1, 1, 0, 0, 0, 1, 1)
    Cube(32) = CreateVertex(-1, -1, 1, 0, 0, 0, 0, 1)
        
    vN = GenerateTriangleNormals(Cube(30), Cube(31), Cube(32))
            
    Cube(30).nx = vN.X: Cube(30).ny = vN.Y: Cube(30).nz = vN.Z
    Cube(31).nx = vN.X: Cube(31).ny = vN.Y: Cube(31).nz = vN.Z
    Cube(32).nx = vN.X: Cube(32).ny = vN.Y: Cube(32).nz = vN.Z
        
    Cube(33) = CreateVertex(1, -1, 1, 0, 0, 0, 1, 1)
    Cube(34) = CreateVertex(-1, -1, -1, 0, 0, 0, 0, 0)
    Cube(35) = CreateVertex(1, -1, -1, 0, 0, 0, 1, 0)
        
    vN = GenerateTriangleNormals(Cube(33), Cube(34), Cube(35))
            
    Cube(33).nx = vN.X: Cube(33).ny = vN.Y: Cube(33).nz = vN.Z
    Cube(34).nx = vN.X: Cube(34).ny = vN.Y: Cube(34).nz = vN.Z
    Cube(35).nx = vN.X: Cube(35).ny = vN.Y: Cube(35).nz = vN.Z

    'Create vertex buffer
    Set VBuffer = D3DDevice.CreateVertexBuffer(Len(Cube(0)) * 36, 0, UNLIT_FVF, D3DPOOL_DEFAULT)
        
    If VBuffer Is Nothing Then
        MsgBox "Unable to create a vertext buffer. Unable to continue..."
        GoTo ErrHandler
    End If
    
    'Fill buffer with geometry
    D3DVertexBuffer8SetData VBuffer, 0, Len(Cube(0)) * 36, 0, Cube(0)

    InitialiseGeometry = True
        
Exit Function
ErrHandler:
    InitialiseGeometry = False
    
End Function


Private Function SetupLights() As Boolean

  Dim Mtrl As D3DMATERIAL8 'Material
  Dim Col As D3DCOLORVALUE 'Color

    On Error GoTo ErrHandler:
    
    'Create color
    Col.a = 1
    Col.r = 1
    Col.g = 1
    Col.b = 1
    
    'Apply material
    Mtrl.Ambient = Col
    Mtrl.diffuse = Col
    D3DDevice.SetMaterial Mtrl
    
    'Create directional light
    Lights(0).Type = D3DLIGHT_DIRECTIONAL
    Lights(0).diffuse.r = 1
    Lights(0).diffuse.g = 1
    Lights(0).diffuse.b = 1
    Lights(0).Direction = MakeVector(0, -1, 0)
    
    'Create point light
    Lights(1).Type = D3DLIGHT_POINT
    Lights(1).Position = MakeVector(5, 0, 2)
    Lights(1).diffuse.b = 1
    Lights(1).Range = 100
    Lights(1).Attenuation1 = 0.05
    
    'Create spotlight
    Lights(2).Type = D3DLIGHT_SPOT
    Lights(2).Position = MakeVector(-4, 0, 0)
    Lights(2).Range = 100
    Lights(2).Direction = MakeVector(1, 0, 0)
    Lights(2).Theta = 30 * Rad
    Lights(2).Phi = 50 * Rad
    Lights(2).diffuse.g = 1
    Lights(2).Attenuation1 = 0.05
    
    'Apply lights to device
    D3DDevice.SetLight 0, Lights(0)
    D3DDevice.SetLight 1, Lights(1)
    D3DDevice.SetLight 2, Lights(2)
    
    'Turn on lights
    Light1On = True
    Light2On = True
    Light3On = True
    
    SetupLights = True
    
Exit Function
ErrHandler:
    MsgBox "An error occured while setting up the scene lighting. Unable to continue..."
    SetupLights = False
    
End Function


Public Sub Render()

    'Clear screen
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0   '//Clear the screen black
    'Start renderer
    D3DDevice.BeginScene
        
        'Apply texture
        D3DDevice.SetTexture 0, Textures(CurrentTexture)
        'Set buffer stream
        D3DDevice.SetStreamSource 0, VBuffer, Len(Cube(0))
        'Draw vertices
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
        
        'Draw the frame rate
        TextRect.Top = 0
        TextRect.bottom = 20
        TextRect.Right = 75
        D3DX.DrawText MainFont, &HFFFFCC00, CStr(FPS_Current) & "fps", TextRect, DT_TOP Or DT_LEFT
        
        'Draw instructions
        TextRect.Top = 20
        TextRect.bottom = 35
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "Keys 1 - 3 to toggle lights on/Off.", TextRect, DT_TOP Or DT_LEFT
        
        TextRect.Top = 35
        TextRect.bottom = 50
        TextRect.Right = 250
        
        D3DX.DrawText MainFont, &HFFFFCC00, "       Light1: " & IIf(Light1On, " on", " off"), TextRect, DT_TOP Or DT_LEFT
        TextRect.Top = 50
        TextRect.bottom = 65
        TextRect.Right = 250
        
        D3DX.DrawText MainFont, &HFFFFCC00, "       Light2: " & IIf(Light2On, " on", " off"), TextRect, DT_TOP Or DT_LEFT
        TextRect.Top = 65
        TextRect.bottom = 80
        TextRect.Right = 250
        
        D3DX.DrawText MainFont, &HFFFFCC00, "       Light3: " & IIf(Light3On, " on", " off"), TextRect, DT_TOP Or DT_LEFT
        
        TextRect.Top = 80
        TextRect.bottom = 95
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "UP/Down Keys to rotate cube on X axis", TextRect, DT_TOP Or DT_LEFT

        TextRect.Top = 95
        TextRect.bottom = 110
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "Left/Right Keys to rotate cube on Y axis", TextRect, DT_TOP Or DT_LEFT

        TextRect.Top = 110
        TextRect.bottom = 125
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "       X Angle: " & RotateAngleX & " degrees", TextRect, DT_TOP Or DT_LEFT
        
        TextRect.Top = 125
        TextRect.bottom = 140
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "       Y Angle: " & RotateAngleY & " degrees", TextRect, DT_TOP Or DT_LEFT

        TextRect.Top = 140
        TextRect.bottom = 155
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "A/Z Keys to zoom in/out", TextRect, DT_TOP Or DT_LEFT
        
        TextRect.Top = 155
        TextRect.bottom = 1770
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "       Zoom Factor " & ZoomFactor, TextRect, DT_TOP Or DT_LEFT

        TextRect.Top = 170
        TextRect.bottom = 185
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "T key to toggle through textures", TextRect, DT_TOP Or DT_LEFT

        TextRect.Top = 185
        TextRect.bottom = 200
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "       Current Texture: " & CurrentTexture, TextRect, DT_TOP Or DT_LEFT

        TextRect.Top = 210
        TextRect.bottom = 225
        TextRect.Right = 250
        D3DX.DrawText MainFont, &HFFFFCC00, "Escape key or Click to exit", TextRect, DT_TOP Or DT_LEFT

    'Close renderer
    D3DDevice.EndScene
    'Flip
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    
End Sub


Public Sub Cleanup()

    'Loop has been terminated so terminate objects
    If ObjPtr(D3DX) Then Set D3DX = Nothing
    If ObjPtr(D3DDevice) Then Set D3DDevice = Nothing
    If ObjPtr(D3D) Then Set D3D = Nothing
    If ObjPtr(Dx) Then Set Dx = Nothing
    
    'Unload form and terminate program
    Unload fEnum
    Unload fMain
    End

End Sub
