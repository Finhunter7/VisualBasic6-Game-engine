VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectX11_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const DebugBuild = True

Private GameEngine As EngineClass

Private m_d3d11Device           As ID3D11Device1
Private m_d3d11DeviceContext    As ID3D11DeviceContext1
Private m_d3d11SwapChain        As IDXGISwapChain1
Private m_d3d11FrameBufferView  As ID3D11RenderTargetView
Private m_isRunning             As Boolean

Sub Initialize(TGameEngine As EngineClass)
    Set GameEngine = TGameEngine
End Sub

Sub Direct11_Start(TargetFormhWnd As Long)
    Dim hResult         As VBHRESULT
    
    '--- Create D3D11 Device and Context
    Dim FeatureLevels() As Long
    Dim creationFlags   As Long
    pvArrayLong FeatureLevels, D3D_FEATURE_LEVEL_11_0
    creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT
    #If DebugBuild Then
        creationFlags = creationFlags Or D3D11_CREATE_DEVICE_DEBUG
    #End If
    
'RetryCreateDevice:
    hResult = CreateDevice(Nothing, D3D_DRIVER_TYPE_HARDWARE, 0, creationFlags, FeatureLevels(0), UBound(FeatureLevels) + 1, D3D11_SDK_VERSION, 0)
    'hResult = D3D11CreateDevice(Nothing, D3D_DRIVER_TYPE_HARDWARE, 0, creationFlags, _
                                featureLevels(0), UBound(featureLevels) + 1, D3D11_SDK_VERSION, _
                                m_d3d11Device, 0, m_d3d11DeviceContext)
    'If hResult = DXGI_ERROR_SDK_COMPONENT_MISSING And (creationFlags And D3D11_CREATE_DEVICE_DEBUG) <> 0 Then
        'creationFlags = creationFlags And Not D3D11_CREATE_DEVICE_DEBUG
        'GoTo RetryCreateDevice
    'End If
    'If hResult < 0 Then
        'Err.Raise hResult, "D3D11CreateDevice"
    'End If
    
#If DebugBuild And False Then
    '--- Set up debug layer to break on D3D11 errors
    Dim d3dDebug        As ID3D11Debug
    Dim d3dInfoQueue    As ID3D11InfoQueue
    Set d3dDebug = m_d3d11Device
    If Not d3dDebug Is Nothing Then
        Set d3dInfoQueue = d3dDebug
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_CORRUPTION, 1
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_ERROR, 1
        Set d3dInfoQueue = Nothing
    End If
    Set d3dDebug = Nothing
#End If
    
    '--- Create Swap Chain
    'Dim dxgiFactory     As IDXGIFactory2
    'Dim dxgiDevice      As IDXGIDevice1
    'Dim dxgiAdapter     As IDXGIAdapter
    'Dim adapterDesc     As DXGI_ADAPTER_DESC
    'Dim d3d11SwapChainDesc As DXGI_SWAP_CHAIN_DESC1
    'Set dxgiDevice = m_d3d11Device
    'Set dxgiAdapter = dxgiDevice.GetAdapter()
    'Set dxgiDevice = Nothing
    
    'dxgiAdapter.GetDesc adapterDesc
    'Debug.Print "Graphics Device: " & Replace(adapterDesc.Description, vbNullChar, vbNullString)
    
    'Set dxgiFactory = dxgiAdapter.GetParent(IIDFromString(szIID_IDXGIFactory2))
    'With d3d11SwapChainDesc
        '.Width = 0 '--- use window width
        '.Height = 0 '--- use window height
        '.Format = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB
        '.SampleDesc.Count = 1
        '.SampleDesc.Quality = 0
        '.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT
        '.BufferCount = 2
        '.Scaling = DXGI_SCALING_STRETCH
        '.SwapEffect = DXGI_SWAP_EFFECT_DISCARD
        '.AlphaMode = DXGI_ALPHA_MODE_UNSPECIFIED
        '.flags = 0
    'End With
    'Set m_d3d11SwapChain = dxgiFactory.CreateSwapChainForHwnd(m_d3d11Device, TargetFormhWnd, d3d11SwapChainDesc, ByVal 0, Nothing)
    Set m_d3d11SwapChain = CreateSwapChain(TargetFormhWnd)
    
    '--- Create Framebuffer Render Target
    Dim d3d11FrameBuffer As ID3D11Texture2D
    Set d3d11FrameBuffer = m_d3d11SwapChain.GetBuffer(0, IIDFromString(szIID_ID3D11Texture2D))
    'Set d3d11FrameBuffer = CreateSwapChain(TargetFormhWnd).GetBuffer(0, IIDFromString(szIID_ID3D11Texture2D))
    Set m_d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
    
    '--- Main Loop
    m_isRunning = True
    Do While m_isRunning
        DoEvents
        Dim backgroundColor() As Single
        pvArraySingle backgroundColor, 0.1!, 0.2!, 0.6!, 1!
        m_d3d11DeviceContext.ClearRenderTargetView m_d3d11FrameBufferView, backgroundColor(0)
        m_d3d11SwapChain.Present 1, 0
    Loop
End Sub

Sub Direct11_Initialize(TargetFormhWnd As Long)
    Set m_d3d11Device = Nothing
    Set m_d3d11DeviceContext = Nothing
    Set m_d3d11SwapChain = Nothing
    Set m_d3d11FrameBufferView = Nothing
    
    Dim hResult         As VBHRESULT
    
    '--- Create D3D11 Device and Context
    Dim FeatureLevels() As Long
    Dim creationFlags   As Long
    pvArrayLong FeatureLevels, D3D_FEATURE_LEVEL_11_0
    creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT
    #If DebugBuild Then
        creationFlags = creationFlags Or D3D11_CREATE_DEVICE_DEBUG
    #End If
    
    hResult = CreateDevice(Nothing, D3D_DRIVER_TYPE_HARDWARE, 0, creationFlags, FeatureLevels(0), UBound(FeatureLevels) + 1, D3D11_SDK_VERSION, 0)
    
#If DebugBuild And False Then
    '--- Set up debug layer to break on D3D11 errors
    Dim d3dDebug        As ID3D11Debug
    Dim d3dInfoQueue    As ID3D11InfoQueue
    Set d3dDebug = m_d3d11Device
    If Not d3dDebug Is Nothing Then
        Set d3dInfoQueue = d3dDebug
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_CORRUPTION, 1
        d3dInfoQueue.SetBreakOnSeverity D3D11_MESSAGE_SEVERITY_ERROR, 1
        Set d3dInfoQueue = Nothing
    End If
    Set d3dDebug = Nothing
#End If
    Set m_d3d11SwapChain = CreateSwapChain(TargetFormhWnd)
    
    '--- Create Framebuffer Render Target
    Dim d3d11FrameBuffer As ID3D11Texture2D
    Set d3d11FrameBuffer = m_d3d11SwapChain.GetBuffer(0, IIDFromString(szIID_ID3D11Texture2D))
    Set m_d3d11FrameBufferView = m_d3d11Device.CreateRenderTargetView(d3d11FrameBuffer, ByVal 0)
End Sub

Function GetDeviceContext() As ID3D11DeviceContext1
    Set GetDeviceContext = m_d3d11DeviceContext
End Function

Function GetSwapChain() As IDXGISwapChain1
    If m_d3d11SwapChain Is Nothing Then
        Err.Raise 1, , "Direct11 Is Not Initialized"
        Exit Function
    End If
    Set GetSwapChain = m_d3d11SwapChain
End Function

Function GetFrameBuffer() As ID3D11RenderTargetView
    Set GetFrameBuffer = m_d3d11FrameBufferView
End Function

Function GetDevice() As ID3D11Device1
    Set GetDevice = m_d3d11Device
End Function

Function CreateDevice( _
    Optional pAdapter As Variant = Nothing, _
    Optional DriverType As D3D_DRIVER_TYPE = D3D_DRIVER_TYPE_HARDWARE, _
    Optional Software As Long = 0, _
    Optional creationFlags As Long = D3D11_CREATE_DEVICE_BGRA_SUPPORT, _
    Optional pfeatureLevels As Long = D3D_FEATURE_LEVEL_11_0, _
    Optional FeatureLevels As Long = D3D_FEATURE_LEVEL_11_0, _
    Optional SDKVersion As Long = D3D11_SDK_VERSION, _
    Optional pFeatureLevel As D3D_FEATURE_LEVEL = 0 _
) As VBHRESULT

    Dim hResult As VBHRESULT
RetryCreateDevice:
    hResult = D3D11CreateDevice(pAdapter, DriverType, Software, creationFlags, pfeatureLevels, FeatureLevels, SDKVersion, m_d3d11Device, pFeatureLevel, m_d3d11DeviceContext)
                                
    If hResult = DXGI_ERROR_SDK_COMPONENT_MISSING And (creationFlags And D3D11_CREATE_DEVICE_DEBUG) <> 0 Then
        creationFlags = creationFlags And Not D3D11_CREATE_DEVICE_DEBUG
        GoTo RetryCreateDevice
    End If
    If hResult < 0 Then
        Err.Raise hResult, "D3D11CreateDevice"
        Exit Function
    End If
    
    CreateDevice = hResult
    
End Function

Function CreateSwapChain( _
    TargetFormhWnd As Long, _
    Optional TWidth As Long = 0, _
    Optional THeight As Long = 0, _
    Optional TFormat As DXGI_FORMAT = DXGI_FORMAT_B8G8R8A8_UNORM_SRGB, _
    Optional SampleDCount As Long = 1, _
    Optional SampleDQuality As Long = 0, _
    Optional TBufferUsage As DXGI_USAGE = DXGI_USAGE_RENDER_TARGET_OUTPUT, _
    Optional TBufferCount As Long = 2, _
    Optional TScaling As DXGI_SCALING = DXGI_SCALING_STRETCH, _
    Optional TSwapEffect As DXGI_SWAP_EFFECT = DXGI_SWAP_EFFECT_DISCARD, _
    Optional TAlphaMode As DXGI_ALPHA_MODE = DXGI_ALPHA_MODE_UNSPECIFIED, _
    Optional TFlags As Long = 0 _
) As IDXGISwapChain1

    Dim dxgiFactory     As IDXGIFactory2
    Dim dxgiDevice      As IDXGIDevice1
    Dim dxgiAdapter     As IDXGIAdapter
    Dim adapterDesc     As DXGI_ADAPTER_DESC
    Dim d3d11SwapChainDesc As DXGI_SWAP_CHAIN_DESC1
    Set dxgiDevice = m_d3d11Device
    Set dxgiAdapter = dxgiDevice.GetAdapter()
    Set dxgiDevice = Nothing
    
    dxgiAdapter.GetDesc adapterDesc
    Debug.Print "Graphics Device: " & Replace(adapterDesc.Description, vbNullChar, vbNullString)
    GameEngine.EConsole.WriteLine "Graphics Device: " & Replace(adapterDesc.Description, vbNullChar, vbNullString)
    GameEngine.EConsole.WriteLine "Graphics Vendor: " & Replace(adapterDesc.VendorId, vbNullChar, vbNullString)
    GameEngine.EConsole.WriteLine "Graphics Revision: " & Replace(adapterDesc.Revision, vbNullChar, vbNullString)
    GameEngine.EConsole.WriteLine "Graphics Memory Avaible: " & Replace(adapterDesc.DedicatedVideoMemory, vbNullChar, vbNullString)
    GameEngine.EConsole.WriteLine "System Memory Avaible: " & Replace(adapterDesc.DedicatedSystemMemory, vbNullChar, vbNullString)
    
    
    Set dxgiFactory = dxgiAdapter.GetParent(IIDFromString(szIID_IDXGIFactory2))
    With d3d11SwapChainDesc
        .Width = TWidth '--- use window width
        .Height = THeight '--- use window height
        .Format = TFormat
        .SampleDesc.Count = SampleDCount
        .SampleDesc.Quality = SampleDQuality
        .BufferUsage = TBufferUsage
        .BufferCount = TBufferCount
        .Scaling = TScaling
        .SwapEffect = TSwapEffect
        .AlphaMode = TAlphaMode
        .flags = TFlags
    End With
    Set CreateSwapChain = dxgiFactory.CreateSwapChainForHwnd(m_d3d11Device, TargetFormhWnd, d3d11SwapChainDesc, ByVal 0, Nothing)
End Function

Private Sub pvArrayLong(aDest() As Long, ParamArray A() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(A)) As Long
    For lIdx = 0 To UBound(A)
        aDest(lIdx) = A(lIdx)
    Next
End Sub


Private Sub pvArraySingle(aDest() As Single, ParamArray A() As Variant)
    Dim lIdx            As Long
    
    ReDim aDest(0 To UBound(A)) As Single
    For lIdx = 0 To UBound(A)
        aDest(lIdx) = A(lIdx)
    Next
End Sub

Sub Direct_Close()
    m_isRunning = False
End Sub

