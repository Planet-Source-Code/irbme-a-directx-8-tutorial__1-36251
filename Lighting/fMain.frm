VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Click()
    bRunning = False   'Terminate
End Sub


Public Sub Start(Adapter As Long, Device As CONST_D3DDEVTYPE, DispModeWidth As Long, DispModeHeight As Long, DispModeFormat As CONST_D3DFORMAT)
  
  Dim matTemp      As D3DMATRIX   'Temporary matrix
  
  'Without these variables, the user can end up
  'toggling something about 10 times with only one key press
  Dim Released1Key As Boolean
  Dim Released2Key As Boolean
  Dim Released3Key As Boolean
  Dim ReleasedTKey As Boolean
  
  'Cursor visibility is not True/False but is incremented/Decremented
  'i.e. If visibility is at 2 and ShowCursor(0) is called, the cursor is
  'still visible.
  Dim CursorVisibility As Long
    
    'Set default values
    CurrentTexture = 0
    ZoomFactor = 9
    Released1Key = True
    Released2Key = True
    Released3Key = True
    
    Me.Show                 'Show form
    bRunning = Initialise(Adapter, Device, DispModeWidth, DispModeHeight, DispModeFormat) 'Initialise
    
    'Hide cursor
    While CursorVisibility >= 0
        CursorVisibility = ShowCursor(0)
    Wend
    
    While bRunning
        
        'Terminate
        If GetAsyncKeyState(vbKeyEscape) Then
            bRunning = False
        End If
        
        'Toggle through textures
        If GetAsyncKeyState(vbKeyT) Then
            ReleasedTKey = False
        ElseIf Not ReleasedTKey Then
            ReleasedTKey = True
            CurrentTexture = CurrentTexture + 1
            If CurrentTexture > 4 Then CurrentTexture = 0
        End If
        
        
        'Update lights
        
        If GetAsyncKeyState(vbKey1) Then
            Released1Key = False
        ElseIf Not Released1Key Then
            Released1Key = True
            Light1On = Not Light1On
        End If
        
        If GetAsyncKeyState(vbKey2) Then
            Released2Key = False
        ElseIf Not Released2Key Then
            Released2Key = True
            Light2On = Not Light2On
        End If
        
        If GetAsyncKeyState(vbKey3) Then
            Released3Key = False
        ElseIf Not Released3Key Then
            Released3Key = True
            Light3On = Not Light3On
        End If
        
        'Update X rotation
        If GetAsyncKeyState(vbKeyUp) Then
            RotateAngleX = RotateAngleX - 2
            If RotateAngleX < 0 Then RotateAngleX = 360
        ElseIf GetAsyncKeyState(vbKeyDown) Then
            RotateAngleX = RotateAngleX + 2
            If RotateAngleX >= 360 Then RotateAngleX = 0
        End If
        
        'Update Y rotation
        If GetAsyncKeyState(vbKeyLeft) Then
            RotateAngleY = RotateAngleY + 2
            If RotateAngleY >= 360 Then RotateAngleY = 0
        ElseIf GetAsyncKeyState(vbKeyRight) Then
            RotateAngleY = RotateAngleY - 2
            If RotateAngleY < 0 Then RotateAngleY = 360
        End If
        
        'Update Zoom
        If GetAsyncKeyState(vbKeyA) Then
            ZoomFactor = ZoomFactor - 1
            If ZoomFactor < 1 Then ZoomFactor = 1
        ElseIf GetAsyncKeyState(vbKeyZ) Then
            ZoomFactor = ZoomFactor + 1
            If ZoomFactor >= 40 Then ZoomFactor = 40
        End If
        
        
        'Clear world matrix
        D3DXMatrixIdentity matWorld
        
        'Clear temporary matrix
        D3DXMatrixIdentity matTemp
        'Rotate world on X axis
        D3DXMatrixRotationX matTemp, RotateAngleX * (PI / 180)
        'Multiply matrices
        D3DXMatrixMultiply matWorld, matWorld, matTemp
        
        'Clear temporary matrix
        D3DXMatrixIdentity matTemp
        'Rotate world on Y axis
        D3DXMatrixRotationY matTemp, RotateAngleY * (PI / 180)
        'Multiply matrices
        D3DXMatrixMultiply matWorld, matWorld, matTemp
        
        
        'Apply newly multiplied world matrix
        D3DDevice.SetTransform D3DTS_WORLD, matWorld
        
        
        
        'Clear view matrix
        D3DXMatrixIdentity matView
        'Apply zoom
        D3DXMatrixLookAtLH matView, MakeVector(0, 5, ZoomFactor), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
        
        
        'Apply new view matrix
        D3DDevice.SetTransform D3DTS_VIEW, matView
        
        
        'Turn On/Off lights
        D3DDevice.LightEnable 0, Light1On
        D3DDevice.LightEnable 1, Light2On
        D3DDevice.LightEnable 2, Light3On

        'Render scene
        Render
        
        'Calculate frame rate
        If GetTickCount() - FPS_LastCheck >= 1000 Then
            FPS_Current = FPS_Count
            FPS_Count = 0
            FPS_LastCheck = GetTickCount()
        End If
        
        FPS_Count = FPS_Count + 1
        
        DoEvents 'Refresh OS
        
    Wend 'Nect frame
    
    'Show cursor
    While CursorVisibility <= 0
        CursorVisibility = ShowCursor(1)
    Wend
    
    Cleanup
    
End Sub
