VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectX9_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Direct93D As IDirect3D9
Private Direct9Device As IDirect3DDevice9

Function Create() As IDirect3D9
    Set Create = Direct3DCreate9()
End Function

Function D9CreateDevice( _
    TargetFormhWnd As Long, _
    Optional Adapter As Long = D3DADAPTER_DEFAULT, _
    Optional DeviceType As D3DDEVTYPE = D3DDEVTYPE_HAL, _
    Optional BehaviourFlags As D3DCREATEFLAGS = D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
    Optional ppBackBufferCount As Long = 1, _
    Optional ppWindowed As Long = 1, _
    Optional ppBackBufferFormat As D3DFORMAT = D3DFMT_A8R8G8B8, _
    Optional ppSwapEffect As D3DSWAPEFFECT = D3DSWAPEFFECT_DISCARD, _
    Optional ppEnableAutoDepthStencil As Long = 1, _
    Optional ppAutoDepthStencilFormat As D3DFORMAT = D3DFMT_D16 _
) As IDirect3DDevice9

    Dim pP As D3DPRESENT_PARAMETERS
    pP.BackBufferCount = ppBackBufferCount
    pP.Windowed = ppWindowed
    pP.BackBufferFormat = ppBackBufferFormat
    pP.SwapEffect = ppSwapEffect
    pP.EnableAutoDepthStencil = ppEnableAutoDepthStencil
    pP.AutoDepthStencilFormat = ppAutoDepthStencilFormat

    Set D9CreateDevice = Direct93D.CreateDevice(Adapter, DeviceType, TargetFormhWnd, BehaviourFlags, pP)
End Function

Private Sub Class_Initialize()
    Set Direct93D = Direct3DCreate9()
End Sub
