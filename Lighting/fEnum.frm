VERSION 5.00
Begin VB.Form fEnum 
   Caption         =   "Lighting Tutorial"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbRes 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1770
      Width           =   4455
   End
   Begin VB.ComboBox cmbDevice 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1050
      Width           =   4455
   End
   Begin VB.ComboBox cmbAdapters 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   330
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2250
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2250
      Width           =   855
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolutions available:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1530
      Width           =   1875
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rendering Devices Available:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   810
      Width           =   2520
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hardware Adapters Available:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   2565
   End
End
Attribute VB_Name = "fEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private AdapterInfo As D3DADAPTER_IDENTIFIER8
Private Device As CONST_D3DDEVTYPE
Private DispMode As D3DDISPLAYMODE
Private Adapter As Long

Private Sub cmbAdapters_Click()

  'Enumerate adapters

    Adapter = cmbAdapters.ListIndex
    cmbAdapters.ListIndex = 0

    'Enumerate devices
    EnumDevices
    cmbDevice.ListIndex = 0

    'Enumerate resolutions
    If InStr(cmbDevice.List(0), "HAL") Then
        Device = D3DDEVTYPE_HAL
      Else
        Device = D3DDEVTYPE_REF
    End If

    EnumDispModes
    cmbRes.ListIndex = cmbRes.ListCount - 1

End Sub

Private Sub cmbDevice_Click()

  'Load the device into the Device variable

    If InStr(cmbDevice.List(cmbDevice.ListIndex), "HAL") Then
        Device = D3DDEVTYPE_HAL
      Else
        Device = D3DDEVTYPE_REF
    End If

    EnumDispModes
    cmbRes.ListIndex = cmbRes.ListCount - 1

End Sub

Private Sub cmbRes_Click()

  'Load resolution into display mode

    D3D.EnumAdapterModes Adapter, cmbRes.ListIndex, DispMode

End Sub

Private Sub cmdCancel_Click()

    Cleanup

End Sub

Private Sub cmdOK_Click()

    fMain.Show          'Show main form
    fMain.Start cmbAdapters.ListIndex, Device, DispMode.Width, DispMode.Height, DispMode.Format
    Me.Hide             'Unload this form
    Unload Me

End Sub

Private Sub EnumAdapters()

  Dim i As Integer, j As Integer, sTemp As String

    cmbAdapters.Clear

    'Loop through all adapters
    For i = 0 To D3D.GetAdapterCount - 1
        'Get adapter info
        D3D.GetAdapterIdentifier i, 0, AdapterInfo

        'Get adapter description
        sTemp = ""
        For j = 0 To 511
            sTemp = sTemp & Chr$(AdapterInfo.Description(j))
        Next j

        sTemp = Replace(sTemp, Chr$(0), "")

        'Add to list
        cmbAdapters.AddItem sTemp
    Next i

End Sub

Private Sub EnumDevices()

    On Local Error Resume Next
    Dim Caps As D3DCAPS8

      cmbDevice.Clear

      D3D.GetDeviceCaps Adapter, D3DDEVTYPE_HAL, Caps

      'Crude way of checking for HAL avaliability
      'But it works.
      If Err.Number = D3DERR_NOTAVAILABLE Then
          cmbDevice.AddItem "Reference Rasterizer (REF)"
        Else
          cmbDevice.AddItem "Hardware Acceleration (HAL)"
          cmbDevice.AddItem "Reference Rasterizer (REF)"
      End If

End Sub

Private Sub EnumDispModes()

  Dim i As Integer

    cmbRes.Clear

    'Loop through all disp modes
    For i = 0 To D3D.GetAdapterModeCount(Adapter) - 1
        D3D.EnumAdapterModes Adapter, i, DispMode
        '32 bit pixel format
        If DispMode.Format = D3DFMT_R8G8B8 Or DispMode.Format = D3DFMT_X8R8G8B8 Or DispMode.Format = D3DFMT_A8R8G8B8 Then
            If D3D.CheckDeviceType(Adapter, Device, DispMode.Format, DispMode.Format, False) >= 0 Then
                cmbRes.AddItem DispMode.Width & "x" & DispMode.Height & "    [32 BIT]"
            End If
            '16 bit pixel format
          Else
            If D3D.CheckDeviceType(Adapter, Device, DispMode.Format, DispMode.Format, False) >= 0 Then
                cmbRes.AddItem DispMode.Width & "x" & DispMode.Height & "    [16 BIT]"
            End If
        End If
    Next i

End Sub

Private Sub Form_Load()

    Me.Show
    Me.Refresh

    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate
    Set D3DX = New D3DX8

    EnumAdapters
    cmbAdapters_Click

End Sub

