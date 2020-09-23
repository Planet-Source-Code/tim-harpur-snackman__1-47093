Attribute VB_Name = "DXext"
'***************************************************************************************************************
'DirectX Basic Interface for Animation/Music/Sound/Input
'                                                     - written by Tim Harpur for Logicon Enterprises
'
'Don't forget to add the appropriate Project->Reference to the DirectX7 or better library
'Also the file AnimationObject.cls is required to be included in the project
'
'NOTE: this library is an outdated/undocumented version included for use with SnackMan
'***************************************************************************************************************

Option Explicit
Option Compare Text

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private dx_DirectX As New DirectX7

'***************************************************************************************************************
'Direct Draw & Animation Control variables
Public Enum FadeModes
  FMBlack = 0
  FMGrey
  FMWhite
  FMBlackVer
  FMBlackHor
End Enum

Public Enum MapModes
  NoBGorFGMap = 0
  SolidBGMap = 1
  TransparentBGMap = 2
  TransparentFGMap = 4
End Enum

Public Enum MapImageType
  MIUseBaseOffset = 4096
End Enum

Public Enum BlitterFX
  BFXNoEffects = 0
  
  BFXStretch = 1
  BFXMirrorLeftRight = 2
  BFXMirrorTopBottom = 4
  BFXTargetColour = 8
  'BFXRotate180 = 16
  'BFXRotate270 = 32
  'BFXRotate90 = 64
  'BFXSmoothEdge = 128
  'BFXAlphaBlend = 256
  'BFXRotated = 512
  
  BFXTransparent = 32768
  BFXSolid = 65536
End Enum

Public Enum FillStyles
  FSSolid = 0
  FSNoFill
  FSHorizontalLine
  FSVerticalLine
  FSUpwardDiagonal
  FSDownwardDiagonal
  FSCross
  FSDiagonalCross
End Enum

Public Enum LineStyles
  LSSolid = 0
  LSDash
  LSDot
  LSDashDot
  LSDashDotDot
  LSNoLine
  LSInsideSolid
End Enum

Public Enum TransformFX
  TFXNoEffect = 0
  TFXSlide
  TFXBlinds
  TFXSplit
  TFXSplitSidelong
  
  
  TFXFreezeLayer1 = 128
  TFXFreezeLayer2 = 256
  TFXHorizontal = 512
  TFXVertical = 1024
End Enum

Private dx_DirectDraw As DirectDraw7

Private dx_DirectDrawEnabled As Boolean, dx_FullScreenMode As Boolean
Private dx_Width As Long, dx_Height As Long, dx_BitDepth As Long
Private m_ClippingWindow As Object

Private dx_DirectDrawPrimarySurface As DirectDrawSurface7
Private dx_DirectDrawPrimaryPalette As DirectDrawPalette
Private dx_DirectDrawPrimaryColourControl As DirectDrawColorControl
Private dx_DirectDrawPrimaryGammaControl As DirectDrawGammaControl
Private dx_DirectDrawBackSurface As DirectDrawSurface7
Private dx_DirectDrawFadeSurface As DirectDrawSurface7

Private dx_DirectDrawStaticSurface() As DirectDrawSurface7
Private dx_StaticSurfaceWidth() As Long, dx_StaticSurfaceHeight() As Long
Private m_TotalStaticSurfaces As Long, m_StaticSurfaceFileName() As String
Private m_StaticSurfaceValid() As Boolean, m_StaticSurfaceTrans() As Long
Private m_StaticSurfaceUseSystem() As Boolean

Private m_AnimationRectangleX As Long, m_AnimationRectangleY As Long
Private m_AnimationRectangleWidth As Long, m_AnimationRectangleHeight As Long

Private dx_NumScreenModes As Long
Private dx_ScreenModeWidth() As Long, dx_ScreenModeHeight() As Long, dx_ScreenModeDepth() As Long

Private m_FirstAnimationObject As AnimationObject, m_LastAnimationObject As AnimationObject

Public MapMode As MapModes

Private m_BGMapImageStaticSurface As Long
Private m_BGMapImageSurfaceWidth As Long, m_BGMapImageSurfaceHeight As Long
Private m_BGMapImageWidth As Long, m_BGMapImageHeight As Long
Private m_BGMapImagesPerRow As Long
Public BGMapBaseImageIndex As Long

Public BGMapShiftX As Long, BGMapShiftY As Long
Public BGMapDisplayWidth As Long, BGMapDisplayHeight As Long
Public BGMapStartRow As Long, BGMapStartColumn As Long
Private m_BGMapArray() As Long
Private m_BGMapWidth As Long, m_BGMapHeight As Long

Private m_FGMapImageStaticSurface As Long
Private m_FGMapImageSurfaceWidth As Long, m_FGMapImageSurfaceHeight As Long
Private m_FGMapImageWidth As Long, m_FGMapImageHeight As Long
Private m_FGMapImagesPerRow As Long
Public FGMapBaseImageIndex As Long

Public FGMapShiftX As Long, FGMapShiftY As Long
Public FGMapDisplayWidth As Long, FGMapDisplayHeight As Long
Public FGMapStartRow As Long, FGMapStartColumn As Long
Private m_FGMapArray() As Long
Private m_FGMapWidth As Long, m_FGMapHeight As Long

Private m_BGPictureStaticSurface As Long
Private m_BGPictureHeight As Long, m_BGPictureWidth As Long
Private m_BGPictureSourceX As Long, m_BGPictureSourceY As Long
Public BGPictureShiftX As Long, BGPictureShiftY As Long, BGPictureWrap As Boolean

Private m_BackColour As Long

'***************************************************************************************************************
'Direct Music & Sound variables
Private dx_DirectSound As DirectSound
Private dx_SoundBuffer() As DirectSoundBuffer
Private dx_SoundBufferDesc() As DSBUFFERDESC
Private dx_WaveFormat() As WAVEFORMATEX
Private m_SoundBufferFileName() As String
Private dx_TotalSoundBuffers As Long

Private dx_TotalMusicChannels As Long
Private dx_DirectMusicLoader As DirectMusicLoader
Private dx_DirectMusicPerformance() As DirectMusicPerformance
Private dx_DirectMusicSegment() As DirectMusicSegment
Private m_MusicChannelFileName() As String

'***************************************************************************************************************
'Direct Input variables - all these should be treated as read-only outside of this module
Private dx_DirectInput As DirectInput

Private dx_DirectKeyboard As DirectInputDevice
Private dx_DirectMouse As DirectInputDevice
Private dx_DirectJoystick As DirectInputDevice

Public dx_KeyboardState As DIKEYBOARDSTATE
Public dx_MouseState As DIMOUSESTATE
Public dx_JoystickState As DIJOYSTATE

Public dx_EnumJoysticks As DirectInputEnumDevices
Private dx_JoystickCaps As DIDEVCAPS

Public Type JoystickDesc
  buttons As Long
  povs As Long
  
  X As Boolean
  Y As Boolean
  z As Boolean
  
  deadzone_x As Long
  deadzone_y As Long
  deadzone_z As Long
  
  saturation_x As Long
  saturation_y As Long
  saturation_z As Long
  
  range_x As Long
  range_y As Long
  range_z As Long
  
  rx As Boolean
  ry As Boolean
  rz As Boolean
  
  deadzone_rx As Long
  deadzone_ry As Long
  deadzone_rz As Long
  
  saturation_rx As Long
  saturation_ry As Long
  saturation_rz As Long
  
  range_rx As Long
  range_ry As Long
  range_rz As Long
  
  slider0 As Boolean
  slider1 As Boolean
End Type

Public dx_JoystickDescribed As JoystickDesc

'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************

'Timing routine
Public Function DelayTillTime(returnTime As Long, Optional MaxCarryOver As Long = 0, Optional ByVal UseRelativeTime As Boolean = False, Optional ByVal callDoEvents As Boolean = True)
  Dim CarryOver As Long
  
  DelayTillTime = timeGetTime()
  If UseRelativeTime Then returnTime = DelayTillTime + returnTime
  
  Do While DelayTillTime < returnTime
    If callDoEvents Then DoEvents
    
    DelayTillTime = timeGetTime()
  Loop
  
  CarryOver = DelayTillTime - returnTime
  If CarryOver > MaxCarryOver Then CarryOver = MaxCarryOver
  
  If CarryOver > 0 Then DelayTillTime = DelayTillTime - CarryOver
End Function


'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************

'Initialize DirectX Sound and Music
Public Sub Init_DXSound(callingForm As Object, Optional ByVal NumSoundChannels As Long = 0, Optional ByVal NumMusicChannels As Long = 0)
  Dim loop1 As Long
  
  On Error Resume Next
  
  CleanUp_DXSound
  
  Set dx_DirectSound = dx_DirectX.DirectSoundCreate("")
  dx_DirectSound.SetCooperativeLevel callingForm.hWnd, DSSCL_PRIORITY
  
  If NumSoundChannels > 0 Then
    ReDim dx_SoundBuffer(1 To NumSoundChannels)
    ReDim dx_SoundBufferDesc(1 To NumSoundChannels)
    ReDim dx_WaveFormat(1 To NumSoundChannels)
    ReDim m_SoundBufferFileName(1 To NumSoundChannels)
  End If
  
  dx_TotalSoundBuffers = NumSoundChannels
  
  If NumMusicChannels > 0 Then
    ReDim dx_DirectMusicPerformance(1 To NumMusicChannels)
    ReDim dx_DirectMusicSegment(1 To NumMusicChannels)
    ReDim m_MusicChannelFileName(1 To NumMusicChannels)
    
    Set dx_DirectMusicLoader = dx_DirectX.DirectMusicLoaderCreate()
    
    For loop1 = 1 To NumMusicChannels
      Set dx_DirectMusicPerformance(loop1) = dx_DirectX.DirectMusicPerformanceCreate()
      
      dx_DirectMusicPerformance(loop1).Init dx_DirectSound, 0
      dx_DirectMusicPerformance(loop1).SetPort -1, 1
      'dx_DirectMusicPerformance(loop1).SetMasterAutoDownload True
      dx_DirectMusicPerformance(loop1).SetMasterVolume 1200
    Next loop1
    
    dx_TotalMusicChannels = NumMusicChannels
  End If
End Sub

Public Property Get TotalSoundBuffers() As Long
  TotalSoundBuffers = dx_TotalSoundBuffers
End Property

Public Property Get TotalMusicChannels() As Long
  TotalMusicChannels = dx_TotalMusicChannels
End Property

Public Sub Play_Music(ByVal ChannelNumber As Long)
  On Error Resume Next
  
  dx_DirectMusicPerformance(ChannelNumber).PlaySegment dx_DirectMusicSegment(ChannelNumber), 0, 0
End Sub

Public Sub Stop_Music(ByVal ChannelNumber As Long)
  On Error Resume Next
  
  dx_DirectMusicPerformance(ChannelNumber).Stop Nothing, Nothing, 0, 0
End Sub

Public Sub LoadMusicFromMidi(ByVal ChannelNumber As Long, MidiFilePathName As String, Optional ByVal Volume As Variant, Optional ByVal Tempo As Variant, Optional ByVal Groove As Variant)
  On Error Resume Next
  
  dx_DirectMusicPerformance(ChannelNumber).Stop Nothing, Nothing, 0, 0
  dx_DirectMusicSegment(ChannelNumber).Unload dx_DirectMusicPerformance(ChannelNumber)
  
  Set dx_DirectMusicSegment(ChannelNumber) = dx_DirectMusicLoader.LoadSegment(MidiFilePathName)
  dx_DirectMusicSegment(ChannelNumber).Download dx_DirectMusicPerformance(ChannelNumber)
  
  m_MusicChannelFileName(ChannelNumber) = MidiFilePathName
  
  If Not IsMissing(Volume) Then dx_DirectMusicPerformance(ChannelNumber).SetMasterVolume Volume
  If Not IsMissing(Tempo) Then dx_DirectMusicPerformance(ChannelNumber).SetMasterTempo Tempo
  If Not IsMissing(Tempo) Then dx_DirectMusicPerformance(ChannelNumber).SetMasterGrooveLevel Groove
End Sub

Public Sub Change_MusicSettings(ByVal ChannelNumber As Long, Optional ByVal Volume As Variant, Optional ByVal Tempo As Variant, Optional ByVal Groove As Variant)
  On Error Resume Next
  
  'Set any parameters
  'Volume is port specific
  'Tempo is from (.25) to (2) with (1) being the default
  'Groove is from (-99) to (99)
  If Not IsMissing(Volume) Then dx_DirectMusicPerformance(ChannelNumber).SetMasterVolume Volume
  If Not IsMissing(Tempo) Then dx_DirectMusicPerformance(ChannelNumber).SetMasterTempo Tempo
  If Not IsMissing(Tempo) Then dx_DirectMusicPerformance(ChannelNumber).SetMasterGrooveLevel Groove
End Sub

'Create the sound buffer - from file
Public Sub CreateSoundBuffer(ByVal ChannelNum As Long, SoundFile As String, Optional ByVal playbackFrequency As Variant, Optional ByVal playbackVolume As Variant, Optional ByVal panLeftRight As Variant)
  On Error Resume Next
  
  If ChannelNum > 0 And ChannelNum <= dx_TotalSoundBuffers Then
    If Not (dx_SoundBuffer(ChannelNum) Is Nothing) Then
      dx_SoundBuffer(ChannelNum).Stop
      
      Set dx_SoundBuffer(ChannelNum) = Nothing
    End If
    
    With dx_SoundBufferDesc(ChannelNum)
      .lFlags = DSBCAPS_STICKYFOCUS Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY
      .lReserved = 0
    End With
        
    Set dx_SoundBuffer(ChannelNum) = dx_DirectSound.CreateSoundBufferFromFile(SoundFile, dx_SoundBufferDesc(ChannelNum), dx_WaveFormat(ChannelNum))
    m_SoundBufferFileName(ChannelNum) = SoundFile
    
    'Set any dx_SoundBuffer() parameters
    If Not IsMissing(playbackFrequency) Then dx_SoundBuffer(ChannelNum).SetFrequency playbackFrequency
    If Not IsMissing(playbackVolume) Then dx_SoundBuffer(ChannelNum).SetVolume playbackVolume
    If Not IsMissing(panLeftRight) Then dx_SoundBuffer(ChannelNum).SetPan panLeftRight
  End If
End Sub

'Return Sound Buffer fileName
Public Function Get_SoundBufferFileName(ByVal ChannelNum As Long) As String
  Get_SoundBufferFileName = m_SoundBufferFileName(ChannelNum)
End Function

'Create the sound buffer - duplicate existing
Public Sub DuplicateSoundBuffer(ByVal ChannelNum As Long, ByVal ChannelNumSource As Long, Optional ByVal playbackFrequency As Variant, Optional ByVal playbackVolume As Variant, Optional ByVal panLeftRight As Variant)
  On Error Resume Next
  
  If ChannelNum > 0 And ChannelNum <= dx_TotalSoundBuffers Then
    If Not (dx_SoundBuffer(ChannelNum) Is Nothing) Then
      dx_SoundBuffer(ChannelNum).Stop
      
      Set dx_SoundBuffer(ChannelNum) = Nothing
    End If
    
    Set dx_SoundBuffer(ChannelNum) = dx_DirectSound.DuplicateSoundBuffer(dx_SoundBuffer(ChannelNumSource))
    m_SoundBufferFileName(ChannelNum) = m_SoundBufferFileName(ChannelNumSource)
    
    'Set any dx_SoundBuffer() parameters
    If Not IsMissing(playbackFrequency) Then dx_SoundBuffer(ChannelNum).SetFrequency playbackFrequency
    If Not IsMissing(playbackVolume) Then dx_SoundBuffer(ChannelNum).SetVolume playbackVolume
    If Not IsMissing(panLeftRight) Then dx_SoundBuffer(ChannelNum).SetPan panLeftRight
  End If
End Sub

'Releases all sound buffers and Direct Sound
Public Sub CleanUp_DXSound()
  Dim loop1 As Long
  
  On Error Resume Next
  
  'Perform for each allocated buffer
  For loop1 = 1 To dx_TotalSoundBuffers
    If Not (dx_SoundBuffer(loop1) Is Nothing) Then
      dx_SoundBuffer(loop1).Stop
      Set dx_SoundBuffer(loop1) = Nothing
    End If
  Next loop1
  
  dx_TotalSoundBuffers = 0
  
  For loop1 = 1 To dx_TotalMusicChannels
    dx_DirectMusicPerformance(loop1).Stop Nothing, Nothing, 0, 0
    dx_DirectMusicPerformance(loop1).CloseDown
    
    Set dx_DirectMusicSegment(loop1) = Nothing
    Set dx_DirectMusicPerformance(loop1) = Nothing
    m_MusicChannelFileName(loop1) = ""
  Next loop1
  
  Set dx_DirectMusicLoader = Nothing
  dx_TotalMusicChannels = 0
  
  Set dx_DirectSound = Nothing
End Sub

'Plays the selected sound buffer
Public Sub Play_SoundBuffer(ByVal ChannelNum As Long, ByVal LoopMode As Boolean)
  On Error Resume Next
  
  dx_SoundBuffer(ChannelNum).SetCurrentPosition 0
  
  If LoopMode Then
    dx_SoundBuffer(ChannelNum).Play DSBPLAY_LOOPING 'Play the sound buffer repeating
  Else
    dx_SoundBuffer(ChannelNum).Play DSBPLAY_DEFAULT 'Play the sound buffer once
  End If
End Sub

'Resumes playing the selected sound buffer
Public Sub Resume_SoundBuffer(ByVal ChannelNum As Long, ByVal LoopMode As Boolean)
  On Error Resume Next
  
  If LoopMode Then
    dx_SoundBuffer(ChannelNum).Play DSBPLAY_LOOPING 'Play the sound buffer repeating
  Else
    dx_SoundBuffer(ChannelNum).Play DSBPLAY_DEFAULT 'Play the sound buffer once
  End If
End Sub

'Stops the selected sound buffer
Public Sub Stop_SoundBuffer(ByVal ChannelNum As Long)
  On Error Resume Next
  
  dx_SoundBuffer(ChannelNum).Stop
End Sub

'Changes the settings for the selected sound buffer
Public Sub Change_SoundSettings(ByVal ChannelNum As Long, Optional ByVal playbackFrequency As Variant, Optional ByVal playbackVolume As Variant, Optional ByVal panLeftRight As Variant)
  On Error Resume Next
  
  'Set any dx_SoundBuffer() parameters
  'Frequency is in Hz
  'Volume is rated in 1/100ths of a dB from max (0) to min (-10000)
  'Pan is is rated in reduction in 1/100ths of a dB with -ve reducing left channel and +ve reducing right
  'channel (-10000) to (10000) with (0) being centered
  If Not IsMissing(playbackFrequency) Then dx_SoundBuffer(ChannelNum).SetFrequency playbackFrequency
  If Not IsMissing(playbackVolume) Then dx_SoundBuffer(ChannelNum).SetVolume playbackVolume
  If Not IsMissing(panLeftRight) Then dx_SoundBuffer(ChannelNum).SetPan panLeftRight
End Sub

'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************

'Initialize Direct Input for use with system keyboard, system mouse and any attached joysticks/controllers
Public Sub Init_DXInput(callingForm As Object, Optional exclusiveMode As Boolean = False, Optional activeController As Long = 1)
  On Error Resume Next
  
  CleanUp_DXInput
  
  Set dx_DirectInput = dx_DirectX.DirectInputCreate
  
  'Grab keyboard nonexclusively while in foreground for UNBUFFERED data
  Set dx_DirectKeyboard = dx_DirectInput.CreateDevice("GUID_SysKeyboard")
  dx_DirectKeyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
  
  If exclusiveMode Then
    dx_DirectKeyboard.SetCooperativeLevel callingForm.hWnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
  Else
    dx_DirectKeyboard.SetCooperativeLevel callingForm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
  End If
  
  dx_DirectKeyboard.Acquire
  
  'Grab mouse exclusively while in foreground for UNBUFFERED data
  Dim dx_DirectMouse_Property As DIPROPLONG
  
  Set dx_DirectMouse = dx_DirectInput.CreateDevice("GUID_SysMouse")
  dx_DirectMouse.SetCommonDataFormat DIFORMAT_MOUSE
  
  If exclusiveMode Then
    dx_DirectMouse.SetCooperativeLevel callingForm.hWnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
  Else
    dx_DirectMouse.SetCooperativeLevel callingForm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
  End If
  
  dx_DirectMouse_Property.lData = DIPROPAXISMODE_REL
  dx_DirectMouse_Property.lHow = DIPH_DEVICE
  dx_DirectMouse_Property.lObj = 0
  dx_DirectMouse_Property.lSize = Len(dx_DirectMouse_Property)
  dx_DirectMouse.SetProperty "DIPROP_AXISMODE", dx_DirectMouse_Property
  
  dx_DirectMouse_Property.lData = DIPROPAXISMODE_REL
  dx_DirectMouse_Property.lHow = DIPH_DEVICE
  dx_DirectMouse_Property.lObj = 0
  dx_DirectMouse_Property.lSize = Len(dx_DirectMouse_Property)
  dx_DirectMouse.SetProperty "DIPROP_AXISMODE", dx_DirectMouse_Property
  
  dx_DirectMouse.Acquire
  
  'Grab attached joystick(s) exclusively while in foreground for UNBUFFERED data
  Set dx_EnumJoysticks = dx_DirectInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
  
  Select_Joystick callingForm, activeController, exclusiveMode
End Sub

'select available joystick from enumerated list
Public Sub Select_Joystick(callingForm As Form, JoystickNum As Long, Optional exclusiveMode As Boolean = False)
  Dim didoEnum As DirectInputEnumDeviceObjects, dido As DirectInputDeviceObjectInstance, loop1 As Long
  
  On Error Resume Next
  
  If JoystickNum > 0 And JoystickNum <= dx_EnumJoysticks.GetCount Then
    Set dx_DirectJoystick = dx_DirectInput.CreateDevice(dx_EnumJoysticks.GetItem(JoystickNum).GetGuidInstance)
    
    dx_DirectJoystick.SetCommonDataFormat DIFORMAT_JOYSTICK
  
    If exclusiveMode Then
      dx_DirectJoystick.SetCooperativeLevel callingForm.hWnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND
    Else
      dx_DirectJoystick.SetCooperativeLevel callingForm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
    End If
    
    dx_DirectJoystick.GetCapabilities dx_JoystickCaps
    
    dx_JoystickDescribed.X = False
    dx_JoystickDescribed.Y = False
    dx_JoystickDescribed.z = False
    dx_JoystickDescribed.rx = False
    dx_JoystickDescribed.ry = False
    dx_JoystickDescribed.rz = False
    dx_JoystickDescribed.slider0 = False
    dx_JoystickDescribed.slider1 = False
    
    Set didoEnum = dx_DirectJoystick.GetDeviceObjectsEnum(DIDFT_AXIS)
    
    For loop1 = 1 To didoEnum.GetCount
      Set dido = didoEnum.GetItem(loop1)
      
      Select Case dido.GetOfs
        Case DIJOFS_X
          dx_JoystickDescribed.X = True
        Case DIJOFS_Y
          dx_JoystickDescribed.Y = True
        Case DIJOFS_Z
          dx_JoystickDescribed.z = True
        Case DIJOFS_RX
          dx_JoystickDescribed.rx = True
        Case DIJOFS_RY
          dx_JoystickDescribed.ry = True
        Case DIJOFS_RZ
          dx_JoystickDescribed.rz = True
        Case DIJOFS_SLIDER0
          dx_JoystickDescribed.slider0 = True
        Case DIJOFS_SLIDER1
          dx_JoystickDescribed.slider1 = True
      End Select
    Next loop1
    
    dx_JoystickDescribed.buttons = dx_JoystickCaps.lButtons
    dx_JoystickDescribed.povs = dx_JoystickCaps.lPOVs
    
    Dim DiProp_Abs As DIPROPLONG
  
    With DiProp_Abs
      .lData = DIPROPAXISMODE_ABS
      .lSize = Len(DiProp_Abs)
      .lHow = DIPH_DEVICE
      
      dx_DirectJoystick.SetProperty "DIPROP_AXISMODE", DiProp_Abs
    End With
    
    Get_JoystickDeadZoneSat
    Get_JoystickRange
    
    dx_DirectJoystick.Acquire
  End If
End Sub

'Release all Direct Input control
Public Sub CleanUp_DXInput()
  On Error Resume Next
  
  dx_DirectKeyboard.Unacquire
  dx_DirectMouse.Unacquire
  dx_DirectJoystick.Unacquire
  
  Set dx_DirectInput = Nothing
End Sub

'Sets active joystick's dead zone and saturation for selected axis
Public Sub Set_JoystickDeadZoneSat(Optional DeadZone As Long = 200, Optional Saturation As Long = 6000, Optional Apply_x As Boolean = False, Optional Apply_y As Boolean = False, Optional Apply_z As Boolean = False, Optional Apply_rx As Boolean = False, Optional Apply_ry As Boolean = False, Optional Apply_rz As Boolean = False)
  Dim DiProp_Dead As DIPROPLONG
  
  On Error Resume Next
  
  With DiProp_Dead
    .lData = DeadZone
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    
    If Apply_x And dx_JoystickDescribed.X Then
      dx_JoystickDescribed.deadzone_x = DeadZone
      
      .lObj = DIJOFS_X
      
      dx_DirectJoystick.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_y And dx_JoystickDescribed.Y Then
      dx_JoystickDescribed.deadzone_y = DeadZone
      
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_z And dx_JoystickDescribed.z Then
      dx_JoystickDescribed.deadzone_z = DeadZone
      
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_rx And dx_JoystickDescribed.rx Then
      dx_JoystickDescribed.deadzone_rx = DeadZone
      
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_ry And dx_JoystickDescribed.ry Then
      dx_JoystickDescribed.deadzone_ry = DeadZone
      
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    If Apply_rz And dx_JoystickDescribed.rz Then
      dx_JoystickDescribed.deadzone_rz = DeadZone
      
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End If
    
    .lData = Saturation
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    
    If Apply_x And dx_JoystickDescribed.X Then
      dx_JoystickDescribed.saturation_x = Saturation
      
      .lObj = DIJOFS_X
      
      dx_DirectJoystick.SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_y And dx_JoystickDescribed.Y Then
      dx_JoystickDescribed.saturation_y = Saturation
      
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick.SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_z And dx_JoystickDescribed.z Then
      dx_JoystickDescribed.saturation_z = Saturation
      
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick.SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_rx And dx_JoystickDescribed.rx Then
      dx_JoystickDescribed.saturation_rx = Saturation
      
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick.SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_ry And dx_JoystickDescribed.ry Then
      dx_JoystickDescribed.saturation_ry = Saturation
      
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick.SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
    
    If Apply_rz And dx_JoystickDescribed.rz Then
      dx_JoystickDescribed.saturation_rz = Saturation
      
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick.SetProperty "DIPROP_SATURATION", DiProp_Dead
    End If
  End With
End Sub

'Sets active joystick's range for selected axis
Public Sub Set_JoystickRange(Optional Range As Long = 10000, Optional Apply_x As Boolean = False, Optional Apply_y As Boolean = False, Optional Apply_z As Boolean = False, Optional Apply_rx As Boolean = False, Optional Apply_ry As Boolean = False, Optional Apply_rz As Boolean = False)
  Dim DiProp_Range As DIPROPRANGE
 
  On Error Resume Next
  
  With DiProp_Range
    .lMin = -Range
    .lMax = Range
    .lSize = Len(DiProp_Range)
    .lHow = DIPH_BYOFFSET
    
    If Apply_x And dx_JoystickDescribed.X Then
      dx_JoystickDescribed.range_x = Range
      
      .lObj = DIJOFS_X
      
      dx_DirectJoystick.SetProperty "DIPROP_RANGE", DiProp_Range
    End If
    
    If Apply_y And dx_JoystickDescribed.Y Then
      dx_JoystickDescribed.range_y = Range
      
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick.SetProperty "DIPROP_RANGE", DiProp_Range
    End If
    
    If Apply_z And dx_JoystickDescribed.z Then
      dx_JoystickDescribed.range_z = Range
      
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick.SetProperty "DIPROP_RANGE", DiProp_Range
    End If
    
    If Apply_rx And dx_JoystickDescribed.rx Then
      dx_JoystickDescribed.range_rx = Range
      
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick.SetProperty "DIPROP_RANGE", DiProp_Range
    End If
    
    If Apply_ry And dx_JoystickDescribed.ry Then
      dx_JoystickDescribed.range_ry = Range
      
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick.SetProperty "DIPROP_RANGE", DiProp_Range
    End If
    
    If Apply_rz And dx_JoystickDescribed.rz Then
      dx_JoystickDescribed.range_rz = Range
      
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick.SetProperty "DIPROP_RANGE", DiProp_Range
    End If
  End With
End Sub

'Gets active joystick's deadzones and saturation levels for selected axis and loads into dx_JoystickDescribed
Private Sub Get_JoystickDeadZoneSat()
  Dim DiProp_Dead As DIPROPLONG
  
  On Error Resume Next
  
  With DiProp_Dead
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    
    If dx_JoystickDescribed.X Then
      .lObj = DIJOFS_X
      
      dx_DirectJoystick.GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_JoystickDescribed.deadzone_x = .lData
      
      dx_DirectJoystick.GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_JoystickDescribed.saturation_x = .lData
    End If
    
    If dx_JoystickDescribed.Y Then
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick.GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_JoystickDescribed.deadzone_y = .lData
      
      dx_DirectJoystick.GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_JoystickDescribed.saturation_y = .lData
    End If
    
    If dx_JoystickDescribed.z Then
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick.GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_JoystickDescribed.deadzone_z = .lData
      
      dx_DirectJoystick.GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_JoystickDescribed.saturation_z = .lData
    End If
    
    If dx_JoystickDescribed.rx Then
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick.GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_JoystickDescribed.deadzone_rx = .lData
      
      dx_DirectJoystick.GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_JoystickDescribed.saturation_rx = .lData
    End If
    
    If dx_JoystickDescribed.ry Then
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick.GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_JoystickDescribed.deadzone_ry = .lData
      
      dx_DirectJoystick.GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_JoystickDescribed.saturation_ry = .lData
    End If
    
    If dx_JoystickDescribed.rz Then
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick.GetProperty "DIPROP_DEADZONE", DiProp_Dead
      
      dx_JoystickDescribed.deadzone_rz = .lData
      
      dx_DirectJoystick.GetProperty "DIPROP_SATURATION", DiProp_Dead
      
      dx_JoystickDescribed.saturation_rz = .lData
    End If
  End With
End Sub

'Gets active joystick's ranges for selected axis and loads into dx_JoystickDescribed
Public Sub Get_JoystickRange()
  Dim DiProp_Range As DIPROPRANGE
 
 On Error Resume Next
 
  With DiProp_Range
    .lSize = Len(DiProp_Range)
    .lHow = DIPH_BYOFFSET
    
    If dx_JoystickDescribed.X Then
      .lObj = DIJOFS_X
      
      dx_DirectJoystick.GetProperty "DIPROP_RANGE", DiProp_Range
      
      dx_JoystickDescribed.range_x = .lMax
    End If
    
    If dx_JoystickDescribed.Y Then
      .lObj = DIJOFS_Y
      
      dx_DirectJoystick.GetProperty "DIPROP_RANGE", DiProp_Range
      
      dx_JoystickDescribed.range_y = .lMax
    End If
    
    If dx_JoystickDescribed.z Then
      .lObj = DIJOFS_Z
      
      dx_DirectJoystick.GetProperty "DIPROP_RANGE", DiProp_Range
      
      dx_JoystickDescribed.range_z = .lMax
    End If
    
    If dx_JoystickDescribed.rx Then
      .lObj = DIJOFS_RX
      
      dx_DirectJoystick.GetProperty "DIPROP_RANGE", DiProp_Range
      
      dx_JoystickDescribed.range_rx = .lMax
    End If
    
    If dx_JoystickDescribed.ry Then
      .lObj = DIJOFS_RY
      
      dx_DirectJoystick.GetProperty "DIPROP_RANGE", DiProp_Range
      
      dx_JoystickDescribed.range_ry = .lMax
    End If
    
    If dx_JoystickDescribed.rz Then
      .lObj = DIJOFS_RZ
      
      dx_DirectJoystick.GetProperty "DIPROP_RANGE", DiProp_Range
      
      dx_JoystickDescribed.range_rz = .lMax
    End If
  End With
End Sub

'Polling of keyboard prepares for call to Read_Keyboard when event handling is disabled
Public Sub PollKeyboard()
  On Error Resume Next
  
  dx_DirectKeyboard.Poll
End Sub

'Polling of mouse prepares for call to Read_Mouse when event handling is disabled
Public Sub PollMouse()
  On Error Resume Next
  
  dx_DirectMouse.Poll
End Sub

'Polling of joystick is required as not all joysticks automatically poll themselves even if event handling
'is enabled
Public Sub PollJoystick()
  On Error Resume Next
  
  dx_DirectJoystick.Poll
End Sub

'Get immediate device state for keyboard
Public Sub Read_Keyboard()
  On Error Resume Next
  
  dx_DirectKeyboard.GetDeviceStateKeyboard dx_KeyboardState
End Sub

'Get immediate device state for mouse
Public Sub Read_Mouse()
  On Error Resume Next
  
  dx_DirectMouse.GetDeviceStateMouse dx_MouseState
End Sub

'Get immediate device state for joystick
Public Sub Read_Joystick()
  On Error Resume Next
  
  dx_DirectJoystick.GetDeviceStateJoystick dx_JoystickState
End Sub


'***************************************************************************************************************
'***************************************************************************************************************
'***************************************************************************************************************

'Initialize the direct draw routines for this animation window
Public Sub Init_DXDrawWindow(ParentForm As Object, Optional ClippingWindow As Object, Optional ByVal NumberOfStaticSurfaces As Long = 0, Optional ByVal refreshOnly As Boolean = False, Optional ByVal UseSystemMemory As Boolean = False)
  Dim loop1 As Long, m_object As AnimationObject
  Dim dx_DirectDrawPrimarySurfaceDesc As DDSURFACEDESC2
  Dim dx_DirectDrawBackSurfaceDesc As DDSURFACEDESC2
  Dim dx_DirectDrawPrimaryClipper As DirectDrawClipper
  
  On Error GoTo badInit
  
  If refreshOnly = True And dx_DirectDrawEnabled = True Then
    Set dx_DirectDrawPrimarySurface = Nothing
    Set dx_DirectDrawPrimaryPalette = Nothing
    Set dx_DirectDrawPrimaryColourControl = Nothing
    Set dx_DirectDrawPrimaryGammaControl = Nothing
    Set dx_DirectDrawBackSurface = Nothing
    Set dx_DirectDrawFadeSurface = Nothing
    
    For loop1 = 1 To m_TotalStaticSurfaces
      Set dx_DirectDrawStaticSurface(loop1) = Nothing
    Next loop1
    
    Set m_object = m_FirstAnimationObject
    
    Do While Not (m_object Is Nothing)
      Set m_object.DXSurface = Nothing
      
      Set m_object = m_object.NextObject
    Loop
    
    dx_Width = 0
    dx_Height = 0
    dx_BitDepth = 0
    
    If dx_FullScreenMode Then dx_DirectDraw.RestoreDisplayMode
    
    NumberOfStaticSurfaces = m_TotalStaticSurfaces
    
    Set dx_DirectDraw = Nothing
  Else
    refreshOnly = False
    
    Cleanup_AnimationWindow
  End If
  
  If ClippingWindow Is Nothing Then
    Set m_ClippingWindow = ParentForm
  Else
    Set m_ClippingWindow = ClippingWindow
  End If
  
  Set dx_DirectDraw = dx_DirectX.DirectDrawCreate("")
  
  dx_DirectDraw.SetCooperativeLevel ParentForm.hWnd, DDSCL_NORMAL
  
  'initailize the primary surface
  With dx_DirectDrawPrimarySurfaceDesc
    .lFlags = DDSD_CAPS
    
    If UseSystemMemory Then
      .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
    Else
      .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End If
  End With
  
  Set dx_DirectDrawPrimarySurface = dx_DirectDraw.CreateSurface(dx_DirectDrawPrimarySurfaceDesc)
  
  'create a full window clipping rectangle for the primary surface
  Set dx_DirectDrawPrimaryClipper = dx_DirectDraw.CreateClipper(0)
  dx_DirectDrawPrimaryClipper.SetHWnd ParentForm.hWnd
  dx_DirectDrawPrimarySurface.SetClipper dx_DirectDrawPrimaryClipper
  
  'initailize the back surface
  With dx_DirectDrawBackSurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
      
      If UseSystemMemory Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
      Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
      End If
    
    .lWidth = m_ClippingWindow.ScaleWidth
    .lHeight = m_ClippingWindow.ScaleHeight
  End With
  
  Set dx_DirectDrawBackSurface = dx_DirectDraw.CreateSurface(dx_DirectDrawBackSurfaceDesc)
  
  CommonInit NumberOfStaticSurfaces, refreshOnly, UseSystemMemory
  
  On Error Resume Next
  
  Dim caps1 As DDCAPS, caps2 As DDCAPS
  dx_DirectDraw.GetCaps caps1, caps2
  
  If caps1.lCaps2 And DDCAPS2_PRIMARYGAMMA Then Set dx_DirectDrawPrimaryGammaControl = dx_DirectDrawPrimarySurface.GetDirectDrawGammaControl
  If caps1.lCaps2 And DDCAPS2_COLORCONTROLPRIMARY Then Set dx_DirectDrawPrimaryColourControl = dx_DirectDrawPrimarySurface.GetDirectDrawColorControl
  
  dx_DirectDrawEnabled = True
  dx_FullScreenMode = False
  dx_Width = m_ClippingWindow.ScaleWidth
  dx_Height = m_ClippingWindow.ScaleHeight
  
  Exit Sub
  
badInit:
  dx_DirectDrawEnabled = False
End Sub

'Initialize the direct draw routines for this animation window as full screen
Public Sub Init_DXDrawScreen(ParentForm As Object, Optional ByVal PixelWidth As Long = 800, Optional ByVal PixelHeight As Long = 600, Optional ByVal PixelDepth As Long = 16, Optional ByVal preferredRefreshRate As Long = 0, Optional ByVal NumberOfStaticSurfaces As Long = 0, Optional ByVal refreshOnly As Boolean = False, Optional ByVal UseSystemMemory As Boolean = False)
  Dim loop1 As Long, m_object As AnimationObject
  Dim dx_DirectDrawPrimarySurfaceDesc As DDSURFACEDESC2
  
  On Error GoTo badInit
  
  If refreshOnly = True And dx_DirectDrawEnabled = True Then
    Set dx_DirectDrawPrimarySurface = Nothing
    Set dx_DirectDrawPrimaryPalette = Nothing
    Set dx_DirectDrawPrimaryColourControl = Nothing
    Set dx_DirectDrawPrimaryGammaControl = Nothing
    Set dx_DirectDrawBackSurface = Nothing
    Set dx_DirectDrawFadeSurface = Nothing
    
    For loop1 = 1 To m_TotalStaticSurfaces
      Set dx_DirectDrawStaticSurface(loop1) = Nothing
    Next loop1
    
    Set m_object = m_FirstAnimationObject
    
    Do While Not (m_object Is Nothing)
      Set m_object.DXSurface = Nothing
      
      Set m_object = m_object.NextObject
    Loop
    
    If dx_FullScreenMode Then dx_DirectDraw.RestoreDisplayMode
    
    NumberOfStaticSurfaces = m_TotalStaticSurfaces
    
    Set dx_DirectDraw = Nothing
  Else
    refreshOnly = False
    
    Cleanup_AnimationWindow
  End If
 
  Set dx_DirectDraw = dx_DirectX.DirectDrawCreate("")
  
  dx_DirectDraw.SetCooperativeLevel ParentForm.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
  
  On Error GoTo badRefresh
  dx_DirectDraw.SetDisplayMode PixelWidth, PixelHeight, PixelDepth, preferredRefreshRate, DDSDM_DEFAULT
  On Error GoTo badInit
  If preferredRefreshRate = -1 Then dx_DirectDraw.SetDisplayMode PixelWidth, PixelHeight, PixelDepth, 0, DDSDM_DEFAULT
  
  'initailize the primary surface
  With dx_DirectDrawPrimarySurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    
    If UseSystemMemory Then
      .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_SYSTEMMEMORY
    Else
      .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    End If
    
    .lBackBufferCount = 1
  End With
  
  Set dx_DirectDrawPrimarySurface = dx_DirectDraw.CreateSurface(dx_DirectDrawPrimarySurfaceDesc)
  
  If PixelDepth = 8 Then
    Dim ptemp(0 To 255) As PALETTEENTRY
    
    Set dx_DirectDrawPrimaryPalette = dx_DirectDraw.CreatePalette(DDPCAPS_8BIT Or DDPCAPS_ALLOW256, ptemp)
    
    ResetDefaultPalette
    
    dx_DirectDrawPrimarySurface.SetPalette dx_DirectDrawPrimaryPalette
  End If
  
  'initailize the back surface
  Dim Caps As DDSCAPS2
  Caps.lCaps = DDSCAPS_BACKBUFFER
  
  Set dx_DirectDrawBackSurface = dx_DirectDrawPrimarySurface.GetAttachedSurface(Caps)
  
  CommonInit NumberOfStaticSurfaces, refreshOnly, UseSystemMemory
  
  On Error Resume Next
  
  Dim caps1 As DDCAPS, caps2 As DDCAPS
  dx_DirectDraw.GetCaps caps1, caps2
  
  If caps1.lCaps2 And DDCAPS2_PRIMARYGAMMA Then Set dx_DirectDrawPrimaryGammaControl = dx_DirectDrawPrimarySurface.GetDirectDrawGammaControl
  If caps1.lCaps2 And DDCAPS2_COLORCONTROLPRIMARY Then Set dx_DirectDrawPrimaryColourControl = dx_DirectDrawPrimarySurface.GetDirectDrawColorControl
  
  dx_DirectDrawEnabled = True
  dx_FullScreenMode = True
  dx_Width = PixelWidth
  dx_Height = PixelHeight
  
  Exit Sub
  
badRefresh:
  preferredRefreshRate = -1
  
  Resume Next
  
badInit:
  dx_DirectDraw.RestoreDisplayMode
  dx_DirectDrawEnabled = False
  dx_Width = 0
  dx_Height = 0
  dx_BitDepth = 0
End Sub

Private Sub CommonInit(ByVal NumberOfStaticSurfaces As Long, ByVal refreshOnly As Boolean, Optional ByVal UseSystemMemory As Boolean)
  Dim loop1 As Long, loop2 As Long, t_Rect As RECT, colourkey As DDCOLORKEY
  Dim m_object As AnimationObject, pixelCaps As DDPIXELFORMAT
  Dim dx_DirectDrawSurfaceDesc As DDSURFACEDESC2
  
  On Error Resume Next
  
  dx_DirectDrawPrimarySurface.GetPixelFormat pixelCaps
  
  If pixelCaps.lFlags & DDPF_RGB Then
    dx_BitDepth = pixelCaps.lRGBBitCount
  ElseIf pixelCaps.lFlags & DDPF_PALETTEINDEXED8 Then
    dx_BitDepth = 8
  Else
    dx_BitDepth = 0
  End If
  
  If refreshOnly = True And dx_DirectDrawEnabled = True Then
    'attempt to restore all graphics
    For loop1 = 1 To NumberOfStaticSurfaces
      If m_StaticSurfaceValid(loop1) Then
        If m_StaticSurfaceFileName(loop1) = "" Then
          Init_StaticSurface loop1, dx_StaticSurfaceWidth(loop1), dx_StaticSurfaceHeight(loop1), m_StaticSurfaceTrans(loop1)
        Else
          Init_StaticSurfaceFromFile loop1, m_StaticSurfaceFileName(loop1), dx_StaticSurfaceWidth(loop1), dx_StaticSurfaceHeight(loop1), m_StaticSurfaceTrans(loop1)
        End If
        
        colourkey.low = m_StaticSurfaceTrans(loop1)
        colourkey.high = m_StaticSurfaceTrans(loop1)
        
        dx_DirectDrawStaticSurface(loop1).SetColorKey DDCKEY_SRCBLT, colourkey
      End If
    Next loop1
    
    Set m_object = m_FirstAnimationObject
    
    Do While Not (m_object Is Nothing)
      With m_object
        If .SurfaceStaticNum <> 0 Then
          Set .DXSurface = dx_DirectDrawStaticSurface(.SurfaceStaticNum)
        End If
      End With
      
      Set m_object = m_object.NextObject
    Loop
  Else
    SetAnimationWindow
    
    're-dim the static surfaces
    If NumberOfStaticSurfaces > 0 Then
      ReDim dx_DirectDrawStaticSurface(1 To NumberOfStaticSurfaces)
      ReDim dx_DirectDrawStaticSurfaceDesc(1 To NumberOfStaticSurfaces)
      ReDim dx_StaticSurfaceWidth(1 To NumberOfStaticSurfaces)
      ReDim dx_StaticSurfaceHeight(1 To NumberOfStaticSurfaces)
      ReDim m_StaticSurfaceValid(1 To NumberOfStaticSurfaces)
      ReDim m_StaticSurfaceFileName(1 To NumberOfStaticSurfaces)
      ReDim m_StaticSurfaceTrans(1 To NumberOfStaticSurfaces)
      ReDim m_StaticSurfaceUseSystem(1 To NumberOfStaticSurfaces)
    End If
    
    m_TotalStaticSurfaces = NumberOfStaticSurfaces
  End If
  
  'create a fade surface to use for FadeBG()
  With dx_DirectDrawSurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    
    If UseSystemMemory Then
      .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Else
      .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    End If
    
    .lWidth = 200
    .lHeight = 40
  End With
  
  Set dx_DirectDrawFadeSurface = dx_DirectDraw.CreateSurface(dx_DirectDrawSurfaceDesc)
  
  With dx_DirectDrawFadeSurface
    .SetColorKey DDCKEY_SRCBLT, colourkey
    
    .BltColorFill t_Rect, &H0
    
    For loop2 = 0 To 2
      Select Case loop2
        Case 0
          .SetForeColor &H80808
        Case 1
          .SetForeColor &H808080
        Case 2
          .SetForeColor &HFFFFFF
      End Select
      
      For loop1 = 1 To 39 Step 2
        .DrawLine loop2 * 40 + loop1, 0, loop2 * 40 + loop1, 40
        .DrawLine loop2 * 40, loop1, loop2 * 40 + 40, loop1
      Next loop1
    Next loop2
    
    .SetForeColor &H80808
    
    For loop1 = 1 To 39 Step 4
      .DrawLine 120 + loop1, 0, 120 + loop1, 40
      .DrawLine 121 + loop1, 0, 121 + loop1, 40
      .DrawLine 122 + loop1, 0, 122 + loop1, 40
      .DrawLine 160, loop1, 200, loop1
      .DrawLine 160, loop1 + 1, 200, loop1 + 1
      .DrawLine 160, loop1 + 2, 200, loop1 + 2
    Next loop1
  End With
End Sub

Public Sub ResetDefaultPalette()
  Dim m_palette(0 To 255) As PALETTEENTRY, loop1 As Long, temp1 As Long
  
  On Error Resume Next
  
  For loop1 = 0 To 31 'grey range
    With m_palette(loop1)
      temp1 = loop1 * 8 + loop1 \ 4
      .red = temp1
      .green = temp1
      .blue = temp1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'red primary range
    With m_palette(loop1 + 32)
      .red = loop1 * 16 + loop1
      .green = 0
      .blue = 0
    End With
  Next loop1
  
  For loop1 = 0 To 15 'red upper range
    With m_palette(loop1 + 48)
      .red = 255
      .green = loop1 * 16 + loop1
      .blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'green primary range
    With m_palette(loop1 + 64)
      .red = 0
      .green = loop1 * 16 + loop1
      .blue = 0
    End With
  Next loop1
  
  For loop1 = 0 To 15 'green upper range
    With m_palette(loop1 + 80)
      .red = loop1 * 16 + loop1
      .green = 255
      .blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'blue primary range
    With m_palette(loop1 + 96)
      .red = 0
      .green = 0
      .blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'blue upper range
    With m_palette(loop1 + 112)
      .red = loop1 * 16 + loop1
      .green = loop1 * 16 + loop1
      .blue = 255
    End With
  Next loop1
  
  For loop1 = 0 To 15 'purple primary range
    With m_palette(loop1 + 128)
      .red = loop1 * 16 + loop1
      .green = 0
      .blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'purple upper range
    With m_palette(loop1 + 144)
      .red = 255
      .green = loop1 * 16 + loop1
      .blue = 255
    End With
  Next loop1
  
  For loop1 = 0 To 15 'yellow primary range
    With m_palette(loop1 + 160)
      .red = loop1 * 16 + loop1
      .green = loop1 * 16 + loop1
      .blue = 0
    End With
  Next loop1
  
  For loop1 = 0 To 15 'yellow upper range
    With m_palette(loop1 + 176)
      .red = 255
      .green = 255
      .blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'cyan primary range
    With m_palette(loop1 + 192)
      .red = 0
      .green = loop1 * 16 + loop1
      .blue = loop1 * 16 + loop1
    End With
  Next loop1
  
  For loop1 = 0 To 15 'cyan upper range
    With m_palette(loop1 + 208)
      .red = loop1 * 16 + loop1
      .green = 255
      .blue = 255
    End With
  Next loop1
  
  For loop1 = 0 To 15 'brown primary range
    With m_palette(loop1 + 224)
      .red = loop1 * 16
      .green = loop1 * 12
      .blue = loop1 * 5
    End With
  Next loop1
  
  For loop1 = 0 To 7 'brown upper range
    With m_palette(loop1 + 240)
      .red = 255
      .green = 191 + loop1 * 8
      .blue = 96 + loop1 * 20
    End With
  Next loop1
  
  For loop1 = 0 To 7 'silver range
    With m_palette(loop1 + 248)
      .red = 84 + loop1 * 16
      .green = 96 + loop1 * 16
      .blue = 128 + loop1 * 16
    End With
  Next loop1
  
  dx_DirectDrawPrimaryPalette.SetEntries 0, 256, m_palette
End Sub

Public Sub SetPalette(paletteBlue() As Long, paletteGreen() As Long, paletteRed() As Long)
  Dim m_palette(0 To 255) As PALETTEENTRY, loop1 As Long
  
  On Error Resume Next
  
  For loop1 = 0 To 255
    With m_palette(loop1)
      .blue = paletteBlue(loop1)
      .green = paletteGreen(loop1)
      .red = paletteRed(loop1)
    End With
  Next loop1
  
  dx_DirectDrawPrimaryPalette.SetEntries 0, 256, m_palette
End Sub

Public Sub LoadPaletteFromBMP(bitmapFilePathName As String)
  Dim m_palette(0 To 255) As PALETTEENTRY
  Dim tPalette As DirectDrawPalette
  
  On Error Resume Next
  
  Set tPalette = dx_DirectDraw.LoadPaletteFromBitmap(bitmapFilePathName)
  
  tPalette.GetEntries 0, 256, m_palette
  dx_DirectDrawPrimaryPalette.SetEntries 0, 256, m_palette
End Sub

Public Sub GetPalette(paletteBlue() As Long, paletteGreen() As Long, paletteRed() As Long)
  Dim m_palette(0 To 255) As PALETTEENTRY, loop1 As Long
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryPalette.GetEntries 0, 256, m_palette
  
  For loop1 = 0 To 255
    With m_palette(loop1)
      paletteBlue(loop1) = .blue
      paletteGreen(loop1) = .green
      paletteRed(loop1) = .red
    End With
  Next loop1
End Sub

Public Sub SetPaletteEntry(ByVal PALETTEENTRY As Long, ByVal paletteBlue As Long, ByVal paletteGreen As Long, ByVal paletteRed As Long)
  Dim m_palette(0 To 0) As PALETTEENTRY
  
  On Error Resume Next
  
  With m_palette(0)
    .blue = paletteBlue
    .green = paletteGreen
    .red = paletteRed
  End With
  
  dx_DirectDrawPrimaryPalette.SetEntries PALETTEENTRY, 1, m_palette
End Sub

Public Sub GetPaletteEntry(ByVal PALETTEENTRY As Long, ByVal paletteBlue As Long, ByVal paletteGreen As Long, ByVal paletteRed As Long)
  Dim m_palette(0 To 0) As PALETTEENTRY
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryPalette.GetEntries PALETTEENTRY, 1, m_palette
  
  With m_palette(0)
    paletteBlue = .blue
    paletteGreen = .green
    paletteRed = .red
  End With
End Sub

Public Property Get FullScreenMode() As Boolean
  FullScreenMode = dx_FullScreenMode
End Property

Public Property Get ScreenWidth() As Long
  ScreenWidth = dx_Width
End Property

Public Property Get ScreenHeight() As Long
  ScreenHeight = dx_Height
End Property

Public Property Get ScreenDepth() As Long
  On Error Resume Next
  
  If dx_FullScreenMode Then
    ScreenDepth = dx_BitDepth
  Else
    Dim m_dispDesc As DDSURFACEDESC2
    
    On Error GoTo badDepth
    
    ScreenDepth = 0
    
    If dx_DirectDrawEnabled Then
      dx_DirectDraw.GetDisplayMode m_dispDesc
      
      With m_dispDesc
        If .ddpfPixelFormat.lFlags And DDPF_RGB Then
          ScreenDepth = m_dispDesc.ddpfPixelFormat.lRGBBitCount
        ElseIf .ddpfPixelFormat.lFlags And DDPF_YUV Then
          ScreenDepth = m_dispDesc.ddpfPixelFormat.lYUVBitCount
        ElseIf .ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED1 Then
          ScreenDepth = 1
        ElseIf .ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED2 Then
          ScreenDepth = 2
        ElseIf .ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED4 Then
          ScreenDepth = 4
        ElseIf .ddpfPixelFormat.lFlags And DDPF_PALETTEINDEXED8 Then
          ScreenDepth = 8
        End If
      End With
    End If
  End If
badDepth:
End Function

Public Property Get TotalDisplayMemory() As Long
  Dim m_caps As DDSCAPS2
  
  On Error Resume Next
  
  m_caps.lCaps = DDSCAPS_VIDEOMEMORY
  
  TotalDisplayMemory = dx_DirectDraw.GetAvailableTotalMem(m_caps)
End Property

Public Property Get FreeDisplayMemory() As Long
  Dim m_caps As DDSCAPS2
  
  On Error Resume Next
  
  m_caps.lCaps = DDSCAPS_VIDEOMEMORY
  
  FreeDisplayMemory = dx_DirectDraw.GetFreeMem(m_caps)
End Property

Public Property Get GammaControlAvailable() As Boolean
  If dx_DirectDrawPrimaryGammaControl Is Nothing Then
    GammaControlAvailable = False
  Else
    GammaControlAvailable = True
  End If
End Property

'only if gamma control available
Public Sub Get_GammaRamp(GammaBlue() As Long, GammaGreen() As Long, GammaRed() As Long)
  Dim m_gamma As DDGAMMARAMP, loop1 As Long
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryGammaControl.GetGammaRamp DDSGR_DEFAULT, m_gamma
  
  For loop1 = 0 To 255
    GammaBlue(loop1) = m_gamma.blue(loop1)
    GammaGreen(loop1) = m_gamma.green(loop1)
    GammaRed(loop1) = m_gamma.red(loop1)
  Next loop1
End Sub

'only if gamma control available
Public Sub Set_GammaRamp(GammaBlue() As Long, GammaGreen() As Long, GammaRed() As Long)
  Dim m_gamma As DDGAMMARAMP, loop1 As Long
  
  On Error Resume Next
  
  For loop1 = 0 To 255
    m_gamma.blue(loop1) = GammaBlue(loop1)
    m_gamma.green(loop1) = GammaGreen(loop1)
    m_gamma.red(loop1) = GammaRed(loop1)
  Next loop1
  
  dx_DirectDrawPrimaryGammaControl.SetGammaRamp DDSGR_DEFAULT, m_gamma
End Sub

Public Property Get ColourControlAvailable() As Boolean
  If dx_DirectDrawPrimaryColourControl Is Nothing Then
    ColourControlAvailable = False
  Else
    ColourControlAvailable = True
  End If
End Property

Public Property Get Brightness() As Long 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_BRIGHTNESS Then
    Brightness = ddcolour.lBrightness
  Else
    Brightness = -1
  End If
End Property

Public Property Get Contrast() As Long 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_CONTRAST Then
    Contrast = ddcolour.lContrast
  Else
    Contrast = -1
  End If
End Property

Public Property Get Gamma() As Long 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_GAMMA Then
    Gamma = ddcolour.lGamma
  Else
    Gamma = -1
  End If
End Property

Public Property Get ColorEnable() As Long 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_COLORENABLE Then
    ColorEnable = ddcolour.lColorEnable
  Else
    ColorEnable = -1
  End If
End Property

Public Property Get Saturation() As Long 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_SATURATION Then
    Saturation = ddcolour.lSaturation
  Else
    Saturation = -1
  End If
End Property

Public Property Get Sharpness() As Long 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_SHARPNESS Then
    Sharpness = ddcolour.lSharpness
  Else
    Sharpness = -1
  End If
End Property

Public Property Get Hue() As Long 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  dx_DirectDrawPrimaryColourControl.GetColorControls ddcolour
  
  If ddcolour.lFlags And DDCOLOR_HUE Then
    Hue = ddcolour.lHue
  Else
    Hue = -1
  End If
End Property

Public Property Let Gamma(ByVal Gamma As Long) 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_GAMMA
  ddcolour.lGamma = Gamma
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Property

Public Property Let ColourEnable(ByVal ColourEnable As Long) 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_COLORENABLE
  ddcolour.lColorEnable = ColourEnable
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Property

Public Property Let Brightness(ByVal Brightness As Long) 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_BRIGHTNESS
  ddcolour.lBrightness = Brightness
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Property

Public Property Let Contrast(ByVal Contrast As Long) 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_CONTRAST
  ddcolour.lContrast = Contrast
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Property

Public Property Let Saturation(ByVal Saturation As Long) 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_SATURATION
  ddcolour.lSaturation = Saturation
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Property

Public Property Let Sharpness(ByVal Sharpness As Long) 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_SHARPNESS
  ddcolour.lSharpness = Sharpness
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Property

Public Property Let Hue(ByVal Hue As Long) 'only if colour control available
  Dim ddcolour As DDCOLORCONTROL
  
  On Error Resume Next
  
  ddcolour.lFlags = DDCOLOR_HUE
  ddcolour.lHue = Hue
  
  dx_DirectDrawPrimaryColourControl.SetColorControls ddcolour
End Property

Public Property Get AvailableDisplayModes(ParentForm As Object) As Long
  Dim DisplayModesEnum As DirectDrawEnumModes, ddsd2 As DDSURFACEDESC2
  Dim loop1 As Long, dd As DirectDraw7

  On Error GoTo badMode
  
  If dx_NumScreenModes = 0 Then 'first time this is called - enumerate available display modes
    If dx_DirectDraw Is Nothing Then
      Set dd = dx_DirectX.DirectDrawCreate("")
      dd.SetCooperativeLevel ParentForm.hWnd, DDSCL_NORMAL
    Else
      Set dd = dx_DirectDraw
    End If
    
    Set DisplayModesEnum = dd.GetDisplayModesEnum(DDEDM_DEFAULT, ddsd2)
    
    dx_NumScreenModes = DisplayModesEnum.GetCount()
    
    ReDim dx_ScreenModeWidth(1 To dx_NumScreenModes)
    ReDim dx_ScreenModeHeight(1 To dx_NumScreenModes)
    ReDim dx_ScreenModeDepth(1 To dx_NumScreenModes)
    
    For loop1 = 1 To dx_NumScreenModes
      DisplayModesEnum.GetItem loop1, ddsd2
      
      dx_ScreenModeWidth(loop1) = ddsd2.lWidth
      dx_ScreenModeHeight(loop1) = ddsd2.lHeight
      dx_ScreenModeDepth(loop1) = ddsd2.ddpfPixelFormat.lRGBBitCount
    Next loop1
  End If
  
  AvailableDisplayModes = dx_NumScreenModes
  
  Exit Property
  
badMode:
  AvailableDisplayModes = 0
End Property

Public Property Get AvailableDisplayModeWidth(ModeNum As Long) As Long
  On Error GoTo badMode
  
  AvailableDisplayModeWidth = dx_ScreenModeWidth(ModeNum)
  
  Exit Property
  
badMode:
  AvailableDisplayModeWidth = 0
End Property

Public Property Get AvailableDisplayModeHeight(ModeNum As Long) As Long
  On Error GoTo badMode
  
  AvailableDisplayModeHeight = dx_ScreenModeHeight(ModeNum)
  
  Exit Property
  
badMode:
  AvailableDisplayModeHeight = 0
End Property

Public Property Get AvailableDisplayModeDepth(ModeNum As Long) As Long
  On Error GoTo badMode
  
  AvailableDisplayModeDepth = dx_ScreenModeDepth(ModeNum)
  
  Exit Property
  
badMode:
  AvailableDisplayModeDepth = 0
End Property

'Returns the status of the animation window - a false requires the window to be initialized or re-initialized
Public Property Get TestDisplayValid() As Boolean
  On Error GoTo notValid
  
  If dx_DirectDrawEnabled Then
    If dx_DirectDraw.TestCooperativeLevel() Then
      TestDisplayValid = False
    Else
      TestDisplayValid = True
    End If
  Else
    TestDisplayValid = False
  End If
  
  Exit Property
  
notValid:
  TestDisplayValid = False
End Property

'Copy the back buffer to the visible animation window on the screen or flip buffers if in full screen mode
Public Sub RefreshDisplay(Optional ByVal WaitForVB As Boolean = True)
  Dim t_Rect As RECT, s_Rect As RECT
  
  On Error Resume Next
  
  If Not dx_DirectDrawEnabled Then Exit Sub
  If dx_FullScreenMode Then Exit Sub
  
  'copy back buffer to window
  If WaitForVB Then dx_DirectDraw.WaitForVerticalBlank DDWAITVB_BLOCKBEGIN, 0
  
  dx_DirectX.GetWindowRect m_ClippingWindow.hWnd, t_Rect
  dx_DirectDrawPrimarySurface.Blt t_Rect, dx_DirectDrawBackSurface, s_Rect, DDBLT_WAIT
End Sub

'Copy the back buffer to the visible animation window on the screen or flip buffers if in full screen mode
Public Sub FlipBuffers(Optional ByVal WaitForVB As Boolean = False, Optional ByVal FastFlip As Boolean = False, Optional ByVal WaitTillFlipComplete As Boolean = False)
  On Error Resume Next
  
  If Not dx_DirectDrawEnabled Then Exit Sub
  If Not dx_FullScreenMode Then Exit Sub
  
  'flip back buffer to front and front to back buffer
  If WaitForVB Then dx_DirectDraw.WaitForVerticalBlank DDWAITVB_BLOCKBEGIN, 0
  
  If FastFlip Then
    dx_DirectDrawPrimarySurface.Flip Nothing, DDFLIP_WAIT Or DDFLIP_NOVSYNC
  Else
    dx_DirectDrawPrimarySurface.Flip Nothing, DDFLIP_WAIT
  End If
  
  If WaitTillFlipComplete Then
    Do While dx_DirectDrawPrimarySurface.GetFlipStatus(DDGFS_ISFLIPDONE) <> DD_OK
      DoEvents
    Loop
  End If
End Sub

'Sysncronize contents of the rendering buffer with the back buffer
Public Sub SyncronizeBuffers()
  Dim t_Rect As RECT, s_Rect As RECT
  
  On Error Resume Next
  
  If Not dx_DirectDrawEnabled Then Exit Sub
  If Not dx_FullScreenMode Then Exit Sub
  
  dx_DirectDrawPrimarySurface.Blt t_Rect, dx_DirectDrawBackSurface, s_Rect, DDBLT_WAIT
End Sub

'Set background colour
Public Property Let RGBBackColour(ByVal RGBColour As Long)
  Dim compRed As Long, compGreen As Long, compBlue As Long
  Dim palEntry As Long, palMatch As Long, palDelta As Long
  Dim DeltaR As Long, DeltaG As Long, DeltaB As Long, DeltaT As Long
  Dim redA(0 To 255) As Long, greenA(0 To 255) As Long, blueA(0 To 255) As Long
  
  compRed = (RGBColour And &HFF0000) \ 65536
  compGreen = (RGBColour And &HFF00) \ 256
  compBlue = RGBColour And &HFF
  
  Select Case dx_BitDepth
    Case 8
      GetPalette blueA, greenA, redA
      
      palMatch = 0
      palDelta = 999
      
      For palEntry = 0 To 255
        DeltaR = redA(palEntry) - compRed
        DeltaT = DeltaR
        If DeltaR < 0 Then DeltaR = -DeltaR
        
        DeltaG = greenA(palEntry) - compGreen
        DeltaT = DeltaT + DeltaG
        If DeltaG < 0 Then DeltaG = -DeltaG
        
        DeltaB = blueA(palEntry) - compBlue
        DeltaT = DeltaT + DeltaB
        If DeltaB < 0 Then DeltaB = -DeltaB
        
        DeltaT = DeltaT + DeltaR + DeltaG + DeltaB
        
        If palDelta > DeltaT Then
          palDelta = DeltaT
          palMatch = palEntry
        End If
      Next palEntry
      
      m_BackColour = palMatch
    Case 15
      compRed = (compRed * 5) \ 8
      compGreen = (compGreen * 5) \ 8
      compBlue = (compBlue * 5) \ 8
      
      m_BackColour = ((compRed * 32) Or compGreen) * 32 Or compBlue
    Case 16
      compRed = (compRed * 5) \ 8
      compGreen = (compGreen * 6) \ 8
      compBlue = (compBlue * 5) \ 8
      
      m_BackColour = ((compRed * 64) Or compGreen) * 32 Or compBlue
    Case 24, 32
      m_BackColour = RGBColour
  End Select
End Property

'Clear the display to the BG colour or BG Picture if one is set
Public Sub ClearDisplay(Optional ByVal FullWindow As Boolean = False)
  Dim t_Rect As RECT, pictSurface As DirectDrawSurface7
  Dim startCol As Long, RowY As Long, RowX As Long
  
  On Error Resume Next
  
  If FullWindow Then
    If m_BGPictureStaticSurface = 0 Then
      dx_DirectDrawBackSurface.BltColorFill t_Rect, m_BackColour
    Else
      If m_AnimationRectangleWidth <> dx_Width Or m_AnimationRectangleHeight <> dx_Height Then dx_DirectDrawBackSurface.BltColorFill t_Rect, m_BackColour
      
      If BGPictureWrap Then
        Set pictSurface = dx_DirectDrawStaticSurface(m_BGPictureStaticSurface)
        
        RowY = BGPictureShiftY Mod m_BGPictureHeight
        If RowY <> 0 Then RowY = RowY - m_BGPictureHeight
        startCol = BGPictureShiftX Mod m_BGPictureWidth
        If startCol <> 0 Then startCol = startCol - m_BGPictureWidth
        
        Do While RowY < m_AnimationRectangleHeight
          RowX = startCol
          
          Do While RowX < m_AnimationRectangleWidth
            BlitSolid pictSurface, RowX, RowY, m_BGPictureSourceX, m_BGPictureSourceY, m_BGPictureWidth, m_BGPictureHeight
            
            RowX = RowX + m_BGPictureWidth
          Loop
          
          RowY = RowY + m_BGPictureHeight
        Loop
      Else
        BlitSolid dx_DirectDrawStaticSurface(m_BGPictureStaticSurface), BGPictureShiftX, BGPictureShiftY, m_BGPictureSourceX, m_BGPictureSourceY, m_BGPictureWidth, m_BGPictureHeight
      End If
    End If
  Else
    If m_BGPictureStaticSurface = 0 Then
      With t_Rect
        .Left = m_AnimationRectangleX
        .Top = m_AnimationRectangleY
        .Right = m_AnimationRectangleWidth + .Left
        .Bottom = m_AnimationRectangleHeight + .Top
      End With
      
      dx_DirectDrawBackSurface.BltColorFill t_Rect, m_BackColour
    ElseIf BGPictureWrap Then
      Set pictSurface = dx_DirectDrawStaticSurface(m_BGPictureStaticSurface)
      
      RowY = BGPictureShiftY Mod m_BGPictureHeight
      If RowY <> 0 Then RowY = RowY - m_BGPictureHeight
      startCol = BGPictureShiftX Mod m_BGPictureWidth
      If startCol <> 0 Then startCol = startCol - m_BGPictureWidth
      
      Do While RowY < m_AnimationRectangleHeight
        RowX = startCol
        
        Do While RowX < m_AnimationRectangleWidth
          BlitSolid pictSurface, RowX, RowY, m_BGPictureSourceX, m_BGPictureSourceY, m_BGPictureWidth, m_BGPictureHeight
           
          RowX = RowX + m_BGPictureWidth
        Loop
        
        RowY = RowY + m_BGPictureHeight
      Loop
    Else
      BlitSolid dx_DirectDrawStaticSurface(m_BGPictureStaticSurface), BGPictureShiftX, BGPictureShiftY, m_BGPictureSourceX, m_BGPictureSourceY, m_BGPictureWidth, m_BGPictureHeight
    End If
  End If
End Sub

'Fade the display using the selected fade mode
Public Sub FadeDisplay(Optional ByVal FullWindow As Boolean = False, Optional ByVal FadeMode As FadeModes)
  Dim loopx As Long, loopy As Long, fadeOffset As Long
  
  On Error Resume Next
  
  fadeOffset = FadeMode * 40
  
  If FullWindow Then
    For loopy = 0 To dx_Height + 39 Step 40
      For loopx = 0 To dx_Width + 39 Step 40
        BlitTransparentFW dx_DirectDrawFadeSurface, loopx, loopy, fadeOffset, 0, 40, 40
      Next loopx
    Next loopy
  Else
    For loopy = 0 To m_AnimationRectangleHeight + 39 Step 40
      For loopx = 0 To m_AnimationRectangleWidth + 39 Step 40
        BlitTransparent dx_DirectDrawFadeSurface, loopx, loopy, fadeOffset, 0, 40, 40
      Next loopx
    Next loopy
  End If
End Sub

'BlitClear the area of display to the background colour
Public Sub BlitClear(ByVal xPos As Long, ByVal yPos As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal FullWindow As Boolean = False)
  Dim t_Rect As RECT
  
  On Error Resume Next
  
  With t_Rect
    If FullWindow Then
      .Left = xPos
      .Top = yPos
      .Right = Width + .Left
      .Bottom = Height + .Top
      
      If .Left < 0 Then .Left = 0
      If .Left >= dx_Width Then Exit Sub
      If .Right <= 0 Then Exit Sub
      If .Right > dx_Width Then .Right = dx_Width
      
      If .Top < 0 Then .Top = 0
      If .Top >= dx_Height Then Exit Sub
      If .Bottom <= 0 Then Exit Sub
      If .Bottom > dx_Height Then .Bottom = dx_Height
    Else
      .Left = m_AnimationRectangleX + xPos
      .Top = m_AnimationRectangleY + yPos
      .Right = Width + .Left
      .Bottom = Height + .Top
      
      If .Left < m_AnimationRectangleX Then .Left = m_AnimationRectangleX
      If .Left >= m_AnimationRectangleX + m_AnimationRectangleWidth Then Exit Sub
      If .Right <= m_AnimationRectangleX Then Exit Sub
      If .Right > m_AnimationRectangleX + m_AnimationRectangleWidth Then .Right = m_AnimationRectangleX + m_AnimationRectangleWidth
      
      If .Top < m_AnimationRectangleY Then .Top = m_AnimationRectangleY
      If .Top >= m_AnimationRectangleY + m_AnimationRectangleHeight Then Exit Sub
      If .Bottom <= m_AnimationRectangleY Then Exit Sub
      If .Bottom > m_AnimationRectangleY + m_AnimationRectangleHeight Then .Bottom = m_AnimationRectangleY + m_AnimationRectangleHeight
    End If
  End With
  
  dx_DirectDrawBackSurface.BltColorFill t_Rect, m_BackColour
End Sub

'Blit the source surface to the co-ordinates on the back buffer using transparency copy
Private Sub BlitTransparent(s_surface As DirectDrawSurface7, ByVal x_pos As Long, _
        ByVal y_pos As Long, ByVal s_xoffset As Long, s_yoffset As Long, ByVal s_width As Long, _
        ByVal s_height As Long)
        
  Dim s_Rect As RECT, t_Rect As RECT
  
  On Error Resume Next
  
  If x_pos >= m_AnimationRectangleWidth Or y_pos >= m_AnimationRectangleHeight Then Exit Sub
  
  With s_Rect
    If x_pos <= 0 Then
      .Left = s_xoffset - x_pos + 1
      s_width = s_width + x_pos - 1
      x_pos = 1
    Else
      .Left = s_xoffset
    End If
    
    If y_pos <= 0 Then
      .Top = s_yoffset - y_pos + 1
      s_height = s_height + y_pos - 1
      y_pos = 1
    Else
      .Top = s_yoffset
    End If
    
    If x_pos + s_width >= m_AnimationRectangleWidth Then s_width = m_AnimationRectangleWidth - x_pos - 1
    .Right = .Left + s_width
    
    If y_pos + s_height >= m_AnimationRectangleHeight Then s_height = m_AnimationRectangleHeight - y_pos - 1
    .Bottom = .Top + s_height
    
    If s_width <= 0 Or s_height <= 0 Then Exit Sub
  End With
  
  With t_Rect
    .Left = x_pos + m_AnimationRectangleX
    .Top = y_pos + m_AnimationRectangleY
    .Right = .Left + s_width
    .Bottom = .Top + s_height
  End With
  
  dx_DirectDrawBackSurface.Blt t_Rect, s_surface, s_Rect, DDBLT_WAIT Or DDBLT_KEYSRC
End Sub

'Blit the source surface to the co-ordinates on the back buffer using transparency copy (full window)
Private Sub BlitTransparentFW(s_surface As DirectDrawSurface7, ByVal x_pos As Long, _
        ByVal y_pos As Long, ByVal s_xoffset As Long, s_yoffset As Long, ByVal s_width As Long, _
        ByVal s_height As Long)
        
  Dim s_Rect As RECT, t_Rect As RECT
  
  On Error Resume Next
  
  If x_pos >= dx_Width Or y_pos >= dx_Height Then Exit Sub
  
  With s_Rect
    If x_pos <= 0 Then
      .Left = s_xoffset - x_pos + 1
      s_width = s_width + x_pos - 1
      x_pos = 1
    Else
      .Left = s_xoffset
    End If
    
    If y_pos <= 0 Then
      .Top = s_yoffset - y_pos + 1
      s_height = s_height + y_pos - 1
      y_pos = 1
    Else
      .Top = s_yoffset
    End If
    
    If x_pos + s_width >= dx_Width Then s_width = dx_Width - x_pos - 1
    .Right = .Left + s_width
    
    If y_pos + s_height >= dx_Height Then s_height = dx_Height - y_pos - 1
    .Bottom = .Top + s_height
    
    If s_width <= 0 Or s_height <= 0 Then Exit Sub
  End With
  
  With t_Rect
    .Left = x_pos
    .Top = y_pos
    .Right = .Left + s_width
    .Bottom = .Top + s_height
  End With
  
  dx_DirectDrawBackSurface.Blt t_Rect, s_surface, s_Rect, DDBLT_WAIT Or DDBLT_KEYSRC
End Sub

'Blit the source static surface to the co-ordinates on the target static surface using transparent copy
Public Sub BlitTransparentS2S(ByVal SourceStaticSurfaceIndex As Long, ByVal TargetStaticSurfaceIndex As Long, _
      ByVal SourceOffsetX As Long, ByVal SourceOffsetY As Long, ByVal SourceWidth As Long, _
      ByVal SourceHeight As Long, ByVal TargetOffsetX As Long, ByVal TargetOffsetY As Long)

  Dim s_Rect As RECT, t_Rect As RECT, m_width As Long, m_height As Long
  
  On Error Resume Next
  
  m_width = dx_StaticSurfaceWidth(TargetStaticSurfaceIndex)
  m_height = dx_StaticSurfaceHeight(TargetStaticSurfaceIndex)
  
  If TargetOffsetX >= m_width Or TargetOffsetY >= m_height Then Exit Sub
  
  With s_Rect
    If TargetOffsetX <= 0 Then
      .Left = SourceOffsetX - TargetOffsetX + 1
      SourceWidth = SourceWidth + TargetOffsetX - 1
      TargetOffsetX = 1
    Else
      .Left = SourceOffsetX
    End If
   
    If TargetOffsetY <= 0 Then
      .Top = SourceOffsetY - TargetOffsetY + 1
      SourceHeight = SourceHeight + TargetOffsetY - 1
      TargetOffsetY = 1
    Else
      .Top = SourceOffsetY
    End If
    
    If TargetOffsetX + SourceWidth >= m_width Then SourceWidth = m_width - TargetOffsetX - 1
    .Right = .Left + SourceWidth
    
    If TargetOffsetY + SourceHeight >= m_height Then SourceHeight = m_height - TargetOffsetY - 1
    .Bottom = .Top + SourceHeight
    
    If SourceWidth <= 0 Or SourceHeight <= 0 Then Exit Sub
  End With
  
  With t_Rect
    .Left = TargetOffsetX
    .Top = TargetOffsetY
    .Right = .Left + SourceWidth
    .Bottom = .Top + SourceHeight
  End With
  
  dx_DirectDrawStaticSurface(TargetStaticSurfaceIndex).Blt t_Rect, dx_DirectDrawStaticSurface(SourceStaticSurfaceIndex), s_Rect, DDBLT_WAIT Or DDBLT_KEYSRC
End Sub

'Blit the source surface to the co-ordinates on the back buffer using solid copy
Private Sub BlitSolid(s_surface As DirectDrawSurface7, ByVal x_pos As Long, ByVal y_pos As Long, _
        ByVal s_xoffset As Long, ByVal s_yoffset As Long, ByVal s_width As Long, ByVal s_height As Long)
        
  Dim s_Rect As RECT, t_Rect As RECT
  
  On Error Resume Next
  
  If x_pos >= m_AnimationRectangleWidth Or y_pos >= m_AnimationRectangleHeight Then Exit Sub
  
  With s_Rect
    If x_pos <= 0 Then
      .Left = s_xoffset - x_pos + 1
      s_width = s_width + x_pos - 1
      x_pos = 1
    Else
      .Left = s_xoffset
    End If
    
    If y_pos <= 0 Then
      .Top = s_yoffset - y_pos + 1
      s_height = s_height + y_pos - 1
      y_pos = 1
    Else
      .Top = s_yoffset
    End If
    
    If x_pos + s_width >= m_AnimationRectangleWidth Then s_width = m_AnimationRectangleWidth - x_pos - 1
    .Right = .Left + s_width
    
    If y_pos + s_height >= m_AnimationRectangleHeight Then s_height = m_AnimationRectangleHeight - y_pos - 1
    .Bottom = .Top + s_height
    
    If s_width <= 0 Or s_height <= 0 Then Exit Sub
  End With
  
  With t_Rect
    .Left = x_pos + m_AnimationRectangleX
    .Top = y_pos + m_AnimationRectangleY
    .Right = .Left + s_width
    .Bottom = .Top + s_height
  End With
  
  dx_DirectDrawBackSurface.Blt t_Rect, s_surface, s_Rect, DDBLT_WAIT
End Sub

'Blit the source surface to the co-ordinates on the back buffer using solid copy (full window)
Private Sub BlitSolidFW(s_surface As DirectDrawSurface7, ByVal x_pos As Long, ByVal y_pos As Long, _
        ByVal s_xoffset As Long, ByVal s_yoffset As Long, ByVal s_width As Long, ByVal s_height As Long)
        
  Dim s_Rect As RECT, t_Rect As RECT
  
  On Error Resume Next
  
  If x_pos >= dx_Width Or y_pos >= dx_Height Then Exit Sub
  
  With s_Rect
    If x_pos <= 0 Then
      .Left = s_xoffset - x_pos + 1
      s_width = s_width + x_pos - 1
      x_pos = 1
    Else
      .Left = s_xoffset
    End If
   
    If y_pos <= 0 Then
      .Top = s_yoffset - y_pos + 1
      s_height = s_height + y_pos - 1
      y_pos = 1
    Else
      .Top = s_yoffset
    End If
    
    If x_pos + s_width >= dx_Width Then s_width = dx_Width - x_pos - 1
    .Right = .Left + s_width
    
    If y_pos + s_height >= dx_Height Then s_height = dx_Height - y_pos - 1
    .Bottom = .Top + s_height
    
    If s_width <= 0 Or s_height <= 0 Then Exit Sub
  End With
  
  With t_Rect
    .Left = x_pos
    .Top = y_pos
    .Right = .Left + s_width
    .Bottom = .Top + s_height
  End With
  
  dx_DirectDrawBackSurface.Blt t_Rect, s_surface, s_Rect, DDBLT_WAIT
End Sub

'Blit the source static surface to the co-ordinates on the target static surface using solid copy
Public Sub BlitSolidS2S(ByVal SourceStaticSurfaceIndex As Long, ByVal TargetStaticSurfaceIndex As Long, _
      ByVal SourceOffsetX As Long, ByVal SourceOffsetY As Long, ByVal SourceWidth As Long, _
      ByVal SourceHeight As Long, ByVal TargetOffsetX As Long, ByVal TargetOffsetY As Long)

  Dim s_Rect As RECT, t_Rect As RECT, m_width As Long, m_height As Long
  
  On Error Resume Next
  
  m_width = dx_StaticSurfaceWidth(TargetStaticSurfaceIndex)
  m_height = dx_StaticSurfaceHeight(TargetStaticSurfaceIndex)
  
  If TargetOffsetX >= m_width Or TargetOffsetY >= m_height Then Exit Sub
  
  With s_Rect
    If TargetOffsetX <= 0 Then
      .Left = SourceOffsetX - TargetOffsetX + 1
      SourceWidth = SourceWidth + TargetOffsetX - 1
      TargetOffsetX = 1
    Else
      .Left = SourceOffsetX
    End If
   
    If TargetOffsetY <= 0 Then
      .Top = SourceOffsetY - TargetOffsetY + 1
      SourceHeight = SourceHeight + TargetOffsetY - 1
      TargetOffsetY = 1
    Else
      .Top = SourceOffsetY
    End If
    
    If TargetOffsetX + SourceWidth >= m_width Then SourceWidth = m_width - TargetOffsetX - 1
    .Right = .Left + SourceWidth
    
    If TargetOffsetY + SourceHeight >= m_height Then SourceHeight = m_height - TargetOffsetY - 1
    .Bottom = .Top + SourceHeight
    
    If SourceWidth <= 0 Or SourceHeight <= 0 Then Exit Sub
  End With
  
  With t_Rect
    .Left = TargetOffsetX
    .Top = TargetOffsetY
    .Right = .Left + SourceWidth
    .Bottom = .Top + SourceHeight
  End With
  
  dx_DirectDrawStaticSurface(TargetStaticSurfaceIndex).Blt t_Rect, dx_DirectDrawStaticSurface(SourceStaticSurfaceIndex), s_Rect, DDBLT_WAIT
End Sub

'Blit the animation object to the co-ordinates on the back buffer using transparency and SpecialFX
Private Sub BlitFX(s_surface As DirectDrawSurface7, ByVal x_pos As Long, ByVal y_pos As Long, _
        ByVal s_xoffset As Long, ByVal s_yoffset As Long, ByVal s_width As Long, ByVal s_height As Long, _
        ByVal SpecialFX As Long, Optional ByVal ScaleWidth As Double = 1, _
        Optional ByVal ScaleHeight As Double = 1, Optional ByVal TargetColour As Long)
        
  Dim s_Rect As RECT, t_Rect As RECT
  Dim t_flags As Long, t_FX As DDBLTFX
  
  On Error Resume Next
  
  If x_pos >= m_AnimationRectangleWidth Or y_pos >= m_AnimationRectangleHeight Then Exit Sub
  
  With s_Rect
    t_flags = DDBLT_WAIT
    
    If SpecialFX And BFXStretch Then
      If x_pos <= 0 Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = s_xoffset
        Else
          .Left = s_xoffset - (x_pos - 1) / ScaleWidth
        End If
        
        s_width = s_width + (x_pos - 1) / ScaleWidth
        
        x_pos = 1
      Else
        .Left = s_xoffset
      End If
      
      If y_pos <= 0 Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = s_yoffset
        Else
          .Top = s_yoffset - (y_pos - 1) / ScaleHeight
        End If
        
        s_height = s_height + (y_pos - 1) / ScaleHeight
        y_pos = 1
      Else
        .Top = s_yoffset
      End If
      
      If x_pos + s_width * ScaleWidth >= m_AnimationRectangleWidth Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = .Left + s_width
          s_width = (dx_Width - x_pos - 1) / ScaleWidth
          .Left = .Left - s_width
        Else
          s_width = (m_AnimationRectangleWidth - x_pos - 1) / ScaleWidth
        End If
      End If
      
      .Right = .Left + s_width
      
      If y_pos + s_height * ScaleHeight >= m_AnimationRectangleHeight Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = .Top + s_height
          s_height = (dx_Height - y_pos - 1) / ScaleHeight
          .Top = .Top - s_height
        Else
          s_height = (m_AnimationRectangleHeight - y_pos - 1) / ScaleHeight
        End If
      End If
      
      .Bottom = .Top + s_height
      
      With t_Rect
        .Top = y_pos + m_AnimationRectangleY
        .Left = x_pos + m_AnimationRectangleX
        .Bottom = .Top + s_height * ScaleHeight
        .Right = .Left + s_width * ScaleWidth
        
        If .Right >= m_AnimationRectangleWidth Then .Right = m_AnimationRectangleWidth - 1
        If .Bottom >= m_AnimationRectangleHeight Then .Bottom = m_AnimationRectangleHeight - 1
      End With
    Else
      If x_pos <= 0 Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = s_xoffset
        Else
          .Left = s_xoffset - x_pos + 1
        End If
        
        s_width = s_width + x_pos - 1
        x_pos = 1
      Else
        .Left = s_xoffset
      End If
      
      If y_pos <= 0 Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = s_yoffset
        Else
          .Top = s_yoffset - y_pos + 1
        End If
        
        s_height = s_height + y_pos - 1
        y_pos = 1
      Else
        .Top = s_yoffset
      End If
      
      If x_pos + s_width >= m_AnimationRectangleWidth Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = .Left + s_width
          s_width = m_AnimationRectangleWidth - x_pos - 1
          .Left = .Left - s_width
        Else
          s_width = m_AnimationRectangleWidth - x_pos - 1
        End If
      End If
      
      .Right = .Left + s_width
      
      If y_pos + s_height >= m_AnimationRectangleHeight Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = .Top + s_height
          s_height = m_AnimationRectangleHeight - y_pos - 1
          .Top = .Top - s_height
        Else
          s_height = m_AnimationRectangleHeight - y_pos - 1
        End If
      End If
      
      .Bottom = .Top + s_height
      
      With t_Rect
        .Top = y_pos + m_AnimationRectangleY
        .Left = x_pos + m_AnimationRectangleX
        .Bottom = .Top + s_height
        .Right = .Left + s_width
      End With
    End If
  End With
  
  If s_width <= 0 Or s_height <= 0 Then Exit Sub
    
  With t_FX
    If SpecialFX And BFXTransparent Then t_flags = t_flags Or DDBLT_KEYSRC
      
    If SpecialFX And BFXTargetColour Then
      t_flags = t_flags Or DDBLT_KEYDESTOVERRIDE
      
      .ddckDestColorKey_low = TargetColour
      .ddckDestColorKey_high = TargetColour
    End If
    
    If SpecialFX And BFXMirrorLeftRight Then
      .lDDFX = .lDDFX Or DDBLTFX_MIRRORLEFTRIGHT
      t_flags = t_flags Or DDBLT_DDFX
    End If
    
    If SpecialFX And BFXMirrorTopBottom Then
      .lDDFX = .lDDFX Or DDBLTFX_MIRRORUPDOWN
      t_flags = t_flags Or DDBLT_DDFX
    End If
    
    'future FX
  End With
  
  dx_DirectDrawBackSurface.BltFx t_Rect, s_surface, s_Rect, t_flags, t_FX
End Sub

'Blit the source surface to the co-ordinates on the back buffer using transparency and SpecialFX (full window)
Private Sub BlitFXFW(s_surface As DirectDrawSurface7, ByVal x_pos As Long, ByVal y_pos As Long, _
        ByVal s_xoffset As Long, ByVal s_yoffset As Long, ByVal s_width As Long, ByVal s_height As Long, _
        ByVal SpecialFX As Long, Optional ByVal ScaleWidth As Double = 1, _
        Optional ByVal ScaleHeight As Double = 1, Optional ByVal TargetColour As Long)
  
  Dim s_Rect As RECT, t_Rect As RECT
  Dim t_flags As Long, t_FX As DDBLTFX
  
  On Error Resume Next
  
  If x_pos >= dx_Width Or y_pos >= dx_Height Then Exit Sub
  
  With s_Rect
    t_flags = DDBLT_WAIT
    
    If SpecialFX And BFXStretch Then
      If x_pos <= 0 Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = s_xoffset
        Else
          .Left = s_xoffset - (x_pos - 1) / ScaleWidth
        End If
        
        s_width = s_width + (x_pos - 1) / ScaleWidth
        
        x_pos = 1
      Else
        .Left = s_xoffset
      End If
      
      If y_pos <= 0 Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = s_yoffset
        Else
          .Top = s_yoffset - (y_pos - 1) / ScaleHeight
        End If
        
        s_height = s_height + (y_pos - 1) / ScaleHeight
        y_pos = 1
      Else
        .Top = s_yoffset
      End If
      
      If x_pos + s_width * ScaleWidth >= dx_Width Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = .Left + s_width
          s_width = (dx_Width - x_pos - 1) / ScaleWidth
          .Left = .Left - s_width
        Else
          s_width = (dx_Width - x_pos - 1) / ScaleWidth
        End If
      End If
      
      .Right = .Left + s_width
      
      If y_pos + s_height * ScaleHeight >= dx_Height Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = .Top + s_height
          s_height = (dx_Height - y_pos - 1) / ScaleHeight
          .Top = .Top - s_height
        Else
          s_height = (dx_Height - y_pos - 1) / ScaleHeight
        End If
      End If
      
      .Bottom = .Top + s_height
      
      With t_Rect
        .Top = y_pos
        .Left = x_pos
        .Bottom = .Top + s_height * ScaleHeight
        .Right = .Left + s_width * ScaleWidth
        
        If .Right >= dx_Width Then .Right = dx_Width - 1
        If .Bottom >= dx_Height Then .Bottom = dx_Height - 1
      End With
    Else
      If x_pos <= 0 Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = s_xoffset
        Else
          .Left = s_xoffset - x_pos + 1
        End If
        
        s_width = s_width + x_pos - 1
        x_pos = 1
      Else
        .Left = s_xoffset
      End If
      
      If y_pos <= 0 Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = s_yoffset
        Else
          .Top = s_yoffset - y_pos + 1
        End If
        
        s_height = s_height + y_pos - 1
        y_pos = 1
      Else
        .Top = s_yoffset
      End If
      
      If x_pos + s_width >= dx_Width Then
        If SpecialFX And BFXMirrorLeftRight Then
          .Left = .Left + s_width
          s_width = dx_Width - x_pos - 1
          .Left = .Left - s_width
        Else
          s_width = dx_Width - x_pos - 1
        End If
      End If
      
      .Right = .Left + s_width
      
      If y_pos + s_height >= dx_Height Then
        If SpecialFX And BFXMirrorTopBottom Then
          .Top = .Top + s_height
          s_height = dx_Height - y_pos - 1
          .Top = .Top - s_height
        Else
          s_height = dx_Height - y_pos - 1
        End If
      End If
      
      .Bottom = .Top + s_height
      
      With t_Rect
        .Top = y_pos
        .Left = x_pos
        .Bottom = .Top + s_height
        .Right = .Left + s_width
      End With
    End If
  End With
  
  If s_width <= 0 Or s_height <= 0 Then Exit Sub
  
  With t_FX
    If SpecialFX And BFXTransparent Then t_flags = t_flags Or DDBLT_KEYSRC
    
    If SpecialFX And BFXTargetColour Then
      t_flags = t_flags Or DDBLT_KEYDESTOVERRIDE
      
      .ddckDestColorKey_low = TargetColour
      .ddckDestColorKey_high = TargetColour
    End If
    
    If SpecialFX And BFXMirrorLeftRight Then
      .lDDFX = .lDDFX Or DDBLTFX_MIRRORLEFTRIGHT
      t_flags = t_flags Or DDBLT_DDFX
    End If
    
    If SpecialFX And BFXMirrorTopBottom Then
      .lDDFX = .lDDFX Or DDBLTFX_MIRRORUPDOWN
      t_flags = t_flags Or DDBLT_DDFX
    End If
    
    'future FX
  End With
  
  dx_DirectDrawBackSurface.BltFx t_Rect, s_surface, s_Rect, t_flags, t_FX
End Sub

Public Property Get TotalStaticSurfaces() As Long
  TotalStaticSurfaces = m_TotalStaticSurfaces
End Property

'initialize a blank static surface surface
Public Sub Init_StaticSurface(ByVal StaticSurfaceIndex As Long, ByVal SurfaceWidth As Long, ByVal SurfaceHeight As Long, Optional ByVal TransparentColour As Long = 0, Optional ByVal SystemMemory As Boolean = False)
  Dim colourkey As DDCOLORKEY
  Dim dx_DirectDrawStaticSurfaceDesc As DDSURFACEDESC2
  
  On Error Resume Next
  
  If StaticSurfaceIndex > 0 And StaticSurfaceIndex <= m_TotalStaticSurfaces Then
    m_StaticSurfaceFileName(StaticSurfaceIndex) = ""
    
    With dx_DirectDrawStaticSurfaceDesc
      .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
      
      If SystemMemory Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
      Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
      End If
      
      .lWidth = SurfaceWidth
      .lHeight = SurfaceHeight
    End With
    
    m_StaticSurfaceUseSystem(StaticSurfaceIndex) = SystemMemory
    
    dx_StaticSurfaceWidth(StaticSurfaceIndex) = SurfaceWidth
    dx_StaticSurfaceHeight(StaticSurfaceIndex) = SurfaceHeight
    
    m_StaticSurfaceTrans(StaticSurfaceIndex) = TransparentColour
    
    Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = Nothing
    Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = dx_DirectDraw.CreateSurface(dx_DirectDrawStaticSurfaceDesc)
    
    colourkey.low = TransparentColour
    colourkey.high = TransparentColour
    
    dx_DirectDrawStaticSurface(StaticSurfaceIndex).SetColorKey DDCKEY_SRCBLT, colourkey
    
    m_StaticSurfaceValid(StaticSurfaceIndex) = True
  End If
  
  Exit Sub
  
badLoad:
  m_StaticSurfaceValid(StaticSurfaceIndex) = False
  
  Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = Nothing
End Sub

'initialize a static surface surface from a bitmap file
Public Sub Init_StaticSurfaceFromFile(ByVal StaticSurfaceIndex As Long, SurfaceFileName As String, ByVal SurfaceWidth As Long, ByVal SurfaceHeight As Long, Optional ByVal TransparentColour As Long = 0, Optional ByVal SystemMemory As Boolean = False)
  Dim colourkey As DDCOLORKEY
  Dim dx_DirectDrawStaticSurfaceDesc As DDSURFACEDESC2
  
  On Error GoTo badLoad
  
  If StaticSurfaceIndex > 0 And StaticSurfaceIndex <= m_TotalStaticSurfaces Then
    With dx_DirectDrawStaticSurfaceDesc
      .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
      
      If SystemMemory Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
      Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
      End If
      
      .lWidth = SurfaceWidth
      .lHeight = SurfaceHeight
    End With
    
    m_StaticSurfaceUseSystem(StaticSurfaceIndex) = SystemMemory
    
    dx_StaticSurfaceWidth(StaticSurfaceIndex) = SurfaceWidth
    dx_StaticSurfaceHeight(StaticSurfaceIndex) = SurfaceHeight
    
    m_StaticSurfaceTrans(StaticSurfaceIndex) = TransparentColour
    
    Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = Nothing
    
    If Right$(SurfaceFileName, 4) <> ".BMP" Then
      Dim srcDC As Long, trgDC As Long, srcPicture As StdPicture
      
      Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = dx_DirectDraw.CreateSurface(dx_DirectDrawStaticSurfaceDesc)
      
      Set srcPicture = LoadPicture(SurfaceFileName)
      
      srcDC = CreateCompatibleDC(ByVal 0&)
      SelectObject srcDC, srcPicture.Handle
      trgDC = dx_DirectDrawStaticSurface(StaticSurfaceIndex).GetDC
      
      BitBlt trgDC, 0, 0, SurfaceWidth, SurfaceHeight, srcDC, 0, 0, vbSrcCopy
      
      dx_DirectDrawStaticSurface(StaticSurfaceIndex).ReleaseDC trgDC
      
      DeleteDC srcDC
      
      Set srcPicture = Nothing
    Else
      Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = dx_DirectDraw.CreateSurfaceFromFile(SurfaceFileName, dx_DirectDrawStaticSurfaceDesc)
    End If
    
    colourkey.low = TransparentColour
    colourkey.high = TransparentColour
    
    dx_DirectDrawStaticSurface(StaticSurfaceIndex).SetColorKey DDCKEY_SRCBLT, colourkey
    
    m_StaticSurfaceFileName(StaticSurfaceIndex) = SurfaceFileName
    m_StaticSurfaceValid(StaticSurfaceIndex) = True
  End If
  
  Exit Sub
  
badLoad:
  m_StaticSurfaceFileName(StaticSurfaceIndex) = ""
  m_StaticSurfaceValid(StaticSurfaceIndex) = False
  
  Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = Nothing
End Sub

Public Sub Release_StaticSurface(ByVal StaticSurfaceIndex As Long)
  On Error Resume Next
  
  Set dx_DirectDrawStaticSurface(StaticSurfaceIndex) = Nothing
    
  m_StaticSurfaceFileName(StaticSurfaceIndex) = ""
  m_StaticSurfaceValid(StaticSurfaceIndex) = False
End Sub

Public Sub DisplayStaticImageSolid(ByVal StaticSurfaceIndex As Long, ByVal SourceOffsetX As Long, _
      ByVal SourceOffsetY As Long, ByVal SourceWidth As Long, ByVal SourceHeight As Long, _
      ByVal TargetOffsetX As Long, ByVal TargetOffsetY As Long, Optional ByVal FullWindow As Boolean = False)
      
  On Error Resume Next
  
  If FullWindow Then
    BlitSolidFW dx_DirectDrawStaticSurface(StaticSurfaceIndex), TargetOffsetX, TargetOffsetY, SourceOffsetX, SourceOffsetY, SourceWidth, SourceHeight
  Else
    BlitSolid dx_DirectDrawStaticSurface(StaticSurfaceIndex), TargetOffsetX, TargetOffsetY, SourceOffsetX, SourceOffsetY, SourceWidth, SourceHeight
  End If
End Sub

Public Sub DisplayStaticImageTransparent(ByVal StaticSurfaceIndex As Long, ByVal SourceOffsetX As Long, _
      ByVal SourceOffsetY As Long, ByVal SourceWidth As Long, ByVal SourceHeight As Long, _
      ByVal TargetOffsetX As Long, ByVal TargetOffsetY As Long, Optional ByVal FullWindow As Boolean = False)
      
  On Error Resume Next
  
  If FullWindow Then
    BlitTransparentFW dx_DirectDrawStaticSurface(StaticSurfaceIndex), TargetOffsetX, TargetOffsetY, SourceOffsetX, SourceOffsetY, SourceWidth, SourceHeight
  Else
    BlitTransparent dx_DirectDrawStaticSurface(StaticSurfaceIndex), TargetOffsetX, TargetOffsetY, SourceOffsetX, SourceOffsetY, SourceWidth, SourceHeight
  End If
End Sub

Public Sub DisplayStaticImageFX(ByVal StaticSurfaceIndex As Long, ByVal SourceOffsetX As Long, _
      ByVal SourceOffsetY As Long, ByVal SourceWidth As Long, ByVal SourceHeight As Long, _
      ByVal TargetOffsetX As Long, ByVal TargetOffsetY As Long, Optional ByVal FullWindow As Boolean = False, _
      Optional ByVal SpecialFX As BlitterFX = BFXTransparent, Optional ByVal ScaleWidth As Long, _
      Optional ByVal ScaleHeight As Long, Optional ByVal TargetColour As Long)
      
  On Error Resume Next
      
  If FullWindow Then
    BlitFXFW dx_DirectDrawStaticSurface(StaticSurfaceIndex), TargetOffsetX, TargetOffsetY, SourceOffsetX, SourceOffsetY, SourceWidth, SourceHeight, SpecialFX, ScaleWidth, ScaleHeight, TargetColour
  Else
    BlitFX dx_DirectDrawStaticSurface(StaticSurfaceIndex), TargetOffsetX, TargetOffsetY, SourceOffsetX, SourceOffsetY, SourceWidth, SourceHeight, SpecialFX, ScaleWidth, ScaleHeight, TargetColour
  End If
End Sub

'Clears the static surface
Public Sub ClearStaticSurface(ByVal StaticSurfaceIndex As Long, Optional ByVal Colour As Long = 0)
  Dim t_Rect As RECT
  
  On Error Resume Next
  
  dx_DirectDrawStaticSurface(StaticSurfaceIndex).BltColorFill t_Rect, Colour
End Sub

'Performs and displays the image transform
Public Sub DisplayTransform(ByVal Transform As TransformFX, ByVal TransformPercent As Long, _
      ByVal xPos As Long, ByVal yPos As Long, ByVal Width As Long, ByVal Height As Long, _
      Optional ByVal StaticSurfaceLayer1 As Long = 0, Optional ByVal Layer1SourceOffsetX As Long = 0, _
      Optional ByVal Layer1SourceOffsetY As Long = 0, Optional ByVal StaticSurfaceLayer2 As Long = 0, _
      Optional ByVal Layer2SourceOffsetX As Long = 0, Optional ByVal Layer2SourceOffsetY As Long = 0, _
      Optional ByVal FullWindow As Boolean = False)
  
  Dim s1Width As Long, s2Width As Long, s1Height As Long, s2Height As Long
  Dim ssWidth As Long, ssHeight As Long, ss2Width As Long, ss2Height As Long
  Dim t_Rect As RECT, m_surface1 As DirectDrawSurface7, m_surface2 As DirectDrawSurface7
  
  On Error Resume Next
  
  If StaticSurfaceLayer1 > 0 Then Set m_surface1 = dx_DirectDrawStaticSurface(StaticSurfaceLayer1)
  If StaticSurfaceLayer2 > 0 Then Set m_surface2 = dx_DirectDrawStaticSurface(StaticSurfaceLayer2)
  
  Select Case Transform And 63
    Case 0 'no effect
      
    Case 1 'slide
      If Transform And TFXHorizontal Then
        s1Width = (Width * TransformPercent) \ 100
        s2Width = Width - s1Width
        
        If Transform And TFXVertical Then 'slide diagonal
          s1Height = (Height * TransformPercent) \ 100
          s2Height = Height - s1Height
          
          If s1Width > 0 And s1Height > 0 Then
            If StaticSurfaceLayer1 = 0 Then
              BlitClear xPos, yPos, s1Width, Height, FullWindow
            ElseIf Transform And TFXFreezeLayer1 Then
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, s1Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, s1Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX + s2Width, Layer1SourceOffsetY + s2Height, s1Width, s1Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX + s2Width, Layer1SourceOffsetY + s2Height, s1Width, s1Height
              End If
            End If
          End If
          
          If s2Width > 0 And s2Height > 0 Then
            If StaticSurfaceLayer2 = 0 Then
              BlitClear xPos + s1Width, yPos + s2Height, s2Width, s2Height, FullWindow
            ElseIf Transform And TFXFreezeLayer2 Then
              If FullWindow Then
                BlitSolidFW m_surface2, xPos + s1Width, yPos + s1Height, Layer2SourceOffsetX + s1Width, Layer2SourceOffsetY + s1Height, s2Width, s2Height
              Else
                BlitSolid m_surface2, xPos + s1Width, yPos + s1Height, Layer2SourceOffsetX + s1Width, Layer2SourceOffsetY + s1Height, s2Width, s2Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface2, xPos + s1Width, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, s2Width, s2Height
              Else
                BlitSolid m_surface2, xPos + s1Width, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, s2Width, s2Height
              End If
            End If
          End If
        Else 'slide horizontal
          If s1Width > 0 Then
            If StaticSurfaceLayer1 = 0 Then
              BlitClear xPos, yPos, s1Width, Height, FullWindow
            ElseIf Transform And TFXFreezeLayer1 Then
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX + s2Width, Layer1SourceOffsetY, s1Width, Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX + s2Width, Layer1SourceOffsetY, s1Width, Height
              End If
            End If
          End If
          
          If s2Width > 0 Then
            If StaticSurfaceLayer2 = 0 Then
              BlitClear xPos + s1Width, yPos, s2Width, Height, FullWindow
            ElseIf Transform And TFXFreezeLayer2 Then
              If FullWindow Then
                BlitSolidFW m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX + s1Width, Layer2SourceOffsetY, s2Width, Height
              Else
                BlitSolid m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX + s1Width, Layer2SourceOffsetY, s2Width, Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX, Layer2SourceOffsetY, s2Width, Height
              Else
                BlitSolid m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX, Layer2SourceOffsetY, s2Width, Height
              End If
            End If
          End If
        End If
      Else 'slide vertical
        s1Height = (Height * TransformPercent) \ 100
        s2Height = Height - s1Height
        
        If s1Height > 0 Then
          If StaticSurfaceLayer1 = 0 Then
            BlitClear xPos, yPos, Width, s1Height, FullWindow
          ElseIf Transform And TFXFreezeLayer1 Then
            If FullWindow Then
              BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, Width, s1Height
            Else
              BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, Width, s1Height
            End If
          Else
            If FullWindow Then
              BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY + s2Height, Width, s1Height
            Else
              BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY + s2Height, Width, s1Height
            End If
          End If
        End If
        
        If s2Height > 0 Then
          If StaticSurfaceLayer2 = 0 Then
            BlitClear xPos, yPos + s1Height, Width, s2Height, FullWindow
          ElseIf Transform And TFXFreezeLayer2 Then
            If FullWindow Then
              BlitSolidFW m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY + s1Height, Width, s2Height
            Else
              BlitSolid m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY + s1Height, Width, s2Height
            End If
          Else
            If FullWindow Then
              BlitSolidFW m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, Width, s2Height
            Else
              BlitSolid m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, Width, s2Height
            End If
          End If
        End If
      End If
    Case 2 'blinds
      
      
      
    Case 3 'split
      If Transform And TFXHorizontal Then
        s1Width = (Width * TransformPercent) \ 200
        s2Width = Width - 2 * s1Width
        ss2Width = s2Width \ 2
        ssWidth = Width \ 2
        
        If Transform And TFXVertical Then 'split diagonal
          s1Height = (Height * TransformPercent) \ 200
          s2Height = Height - 2 * s1Height
          ss2Height = s2Height \ 2
          ssHeight = Height \ 2
          
          If ss2Width > 0 And ss2Height > 0 Then
            If StaticSurfaceLayer2 = 0 Then
              BlitClear xPos, yPos, Width, Height, FullWindow
            ElseIf Transform And TFXFreezeLayer2 Then
              If FullWindow Then
                BlitSolidFW m_surface2, xPos, yPos, Layer2SourceOffsetX, Layer2SourceOffsetY, Width, Height
              Else
                BlitSolid m_surface2, xPos, yPos, Layer2SourceOffsetX, Layer2SourceOffsetY, Width, Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface2, xPos + s1Width, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, ss2Width, ss2Height
                BlitSolidFW m_surface2, xPos + ssWidth, yPos + s1Height, Layer2SourceOffsetX + Width - ss2Width, Layer2SourceOffsetY, ss2Width, ss2Height
                BlitSolidFW m_surface2, xPos + s1Width, yPos + ssHeight, Layer2SourceOffsetX, Layer2SourceOffsetY + Height - ss2Height, ss2Width, ss2Height
                BlitSolidFW m_surface2, xPos + ssWidth, yPos + ssHeight, Layer2SourceOffsetX + Width - ss2Width, Layer2SourceOffsetY + Height - ss2Height, ss2Width, ss2Height
              Else
                BlitSolid m_surface2, xPos + s1Width, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, ss2Width, ss2Height
                BlitSolid m_surface2, xPos + ssWidth, yPos + s1Height, Layer2SourceOffsetX + Width - ss2Width, Layer2SourceOffsetY, ss2Width, ss2Height
                BlitSolid m_surface2, xPos + s1Width, yPos + ssHeight, Layer2SourceOffsetX, Layer2SourceOffsetY + Height - ss2Height, ss2Width, ss2Height
                BlitSolid m_surface2, xPos + ssWidth, yPos + ssHeight, Layer2SourceOffsetX + Width - ss2Width, Layer2SourceOffsetY + Height - ss2Height, ss2Width, ss2Height
              End If
            End If
          End If
          
          If s1Width > 0 And s1Height > 0 Then
            If StaticSurfaceLayer1 = 0 Then
              BlitClear xPos, yPos, s1Width, s1Height, FullWindow
              BlitClear xPos + Width - s1Width, yPos, s1Width, s1Height, FullWindow
              BlitClear xPos, yPos + Height - s1Height, s1Width, s1Height, FullWindow
              BlitClear xPos + Width - s1Width, yPos + Height - s1Height, s1Width, s1Height, FullWindow
            ElseIf Transform And TFXFreezeLayer1 Then
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, s1Height
                BlitSolidFW m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + Width - s1Width, Layer1SourceOffsetY, s1Width, s1Height
                BlitSolidFW m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX, Layer1SourceOffsetY + Height - s1Height, s1Width, s1Height
                BlitSolidFW m_surface1, xPos + Width - s1Width, yPos + Height - s1Height, Layer1SourceOffsetX + Width - s1Width, Layer1SourceOffsetY + Height - s1Height, s1Width, s1Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, s1Height
                BlitSolid m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + Width - s1Width, Layer1SourceOffsetY, s1Width, s1Height
                BlitSolid m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX, Layer1SourceOffsetY + Height - s1Height, s1Width, s1Height
                BlitSolid m_surface1, xPos + Width - s1Width, yPos + Height - s1Height, Layer1SourceOffsetX + Width - s1Width, Layer1SourceOffsetY + Height - s1Height, s1Width, s1Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX + ss2Width, Layer1SourceOffsetY + ss2Height, s1Width, s1Height
                BlitSolidFW m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + ssWidth, Layer1SourceOffsetY + ss2Height, s1Width, s1Height
                BlitSolidFW m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX + ss2Width, Layer1SourceOffsetY + ssHeight, s1Width, s1Height
                BlitSolidFW m_surface1, xPos + Width - s1Width, yPos + Height - s1Height, Layer1SourceOffsetX + ssWidth, Layer1SourceOffsetY + ssHeight, s1Width, s1Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX + ss2Width, Layer1SourceOffsetY + ss2Height, s1Width, s1Height
                BlitSolid m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + ssWidth, Layer1SourceOffsetY + ss2Height, s1Width, s1Height
                BlitSolid m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX + ss2Width, Layer1SourceOffsetY + ssHeight, s1Width, s1Height
                BlitSolid m_surface1, xPos + Width - s1Width, yPos + Height - s1Height, Layer1SourceOffsetX + ssWidth, Layer1SourceOffsetY + ssHeight, s1Width, s1Height
              End If
            End If
          End If
        Else 'split horizontal
          If s2Width > 0 Then
            If StaticSurfaceLayer2 = 0 Then
              BlitClear xPos + s1Width, yPos, s2Width, Height, FullWindow
            ElseIf Transform And TFXFreezeLayer2 Then
              If FullWindow Then
                BlitSolidFW m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX + s1Width, Layer2SourceOffsetY, s2Width, Height
              Else
                BlitSolid m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX + s1Width, Layer2SourceOffsetY, s2Width, Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX, Layer2SourceOffsetY, ss2Width, Height
                BlitSolidFW m_surface2, xPos + ssWidth, yPos, Layer2SourceOffsetX + Width - ss2Width, Layer2SourceOffsetY, ss2Width, Height
              Else
                BlitSolid m_surface2, xPos + s1Width, yPos, Layer2SourceOffsetX, Layer2SourceOffsetY, ss2Width, Height
                BlitSolid m_surface2, xPos + ssWidth, yPos, Layer2SourceOffsetX + Width - ss2Width, Layer2SourceOffsetY, ss2Width, Height
              End If
            End If
          End If
          
          If s1Width > 0 Then
            If StaticSurfaceLayer1 = 0 Then
              BlitClear xPos, yPos, s1Width, Height, FullWindow
              BlitClear xPos + Width - s1Width, yPos, s1Width, Height, FullWindow
            ElseIf Transform And TFXFreezeLayer1 Then
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, Height
                BlitSolidFW m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + Width - s1Width, Layer1SourceOffsetY, s1Width, Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, s1Width, Height
                BlitSolid m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + Width - s1Width, Layer1SourceOffsetY, s1Width, Height
              End If
            Else
              If FullWindow Then
                BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX + ss2Width, Layer1SourceOffsetY, s1Width, Height
                BlitSolidFW m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + ssWidth, Layer1SourceOffsetY, s1Width, Height
              Else
                BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX + ss2Width, Layer1SourceOffsetY, s1Width, Height
                BlitSolid m_surface1, xPos + Width - s1Width, yPos, Layer1SourceOffsetX + ssWidth, Layer1SourceOffsetY, s1Width, Height
              End If
            End If
          End If
        End If
      Else 'split vertical
        s1Height = (Height * TransformPercent) \ 200
        s2Height = Height - 2 * s1Height
        ss2Height = s2Height \ 2
        ssHeight = Height \ 2
        
        If s2Height > 0 Then
          If StaticSurfaceLayer2 = 0 Then
            BlitClear xPos, yPos + s1Height, Width, s2Height, FullWindow
          ElseIf Transform And TFXFreezeLayer2 Then
            If FullWindow Then
              BlitSolidFW m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY + s1Height, Width, s2Height
            Else
              BlitSolid m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY + s1Height, Width, s2Height
            End If
          Else
            If FullWindow Then
              BlitSolidFW m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, Width, ss2Height
              BlitSolidFW m_surface2, xPos, yPos + ssHeight, Layer2SourceOffsetX, Layer2SourceOffsetY + Height - ss2Height, Width, ss2Height
            Else
              BlitSolid m_surface2, xPos, yPos + s1Height, Layer2SourceOffsetX, Layer2SourceOffsetY, Width, ss2Height
              BlitSolid m_surface2, xPos, yPos + ssHeight, Layer2SourceOffsetX, Layer2SourceOffsetY + Height - ss2Height, Width, ss2Height
            End If
          End If
        End If
        
        If s1Height > 0 Then
          If StaticSurfaceLayer1 = 0 Then
            BlitClear xPos, yPos, Width, s1Height, FullWindow
            BlitClear xPos, yPos + Height - s1Height, Width, s1Height, FullWindow
          ElseIf Transform And TFXFreezeLayer1 Then
            If FullWindow Then
              BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, Width, s1Height
              BlitSolidFW m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX, Layer1SourceOffsetY + Height - s1Height, Width, s1Height
            Else
              BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY, Width, s1Height
              BlitSolid m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX, Layer1SourceOffsetY + Height - s1Height, Width, s1Height
            End If
          Else
            If FullWindow Then
              BlitSolidFW m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY + ss2Height, Width, s1Height
              BlitSolidFW m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX, Layer1SourceOffsetY + ssHeight, Width, s1Height
            Else
              BlitSolid m_surface1, xPos, yPos, Layer1SourceOffsetX, Layer1SourceOffsetY + ss2Height, Width, s1Height
              BlitSolid m_surface1, xPos, yPos + Height - s1Height, Layer1SourceOffsetX, Layer1SourceOffsetY + ssHeight, Width, s1Height
            End If
          End If
        End If
      End If
    Case 4 'split sidelong
      
  End Select
End Sub

Public Sub SetFont(mFont As IFont, Optional ByVal StaticSurface As Long = 0)
  Dim dds As DirectDrawSurface7
  
  On Error Resume Next
  
  If StaticSurface = 0 Then
    Set dds = dx_DirectDrawBackSurface
  Else
    Set dds = dx_DirectDrawStaticSurface(StaticSurface)
  End If
  
  dds.SetFont mFont
End Sub

Public Sub DisplayText(textString As String, ByVal xPos As Long, ByVal yPos As Long, ByVal ForeColour As Long, Optional ByVal BackColour = 0, Optional ByVal Transparency As Boolean = True, Optional ByVal StaticSurface As Long = 0, Optional ByVal FullWindow As Boolean = False)
  Dim dds As DirectDrawSurface7
  
  On Error Resume Next
  
  If StaticSurface = 0 Then
    Set dds = dx_DirectDrawBackSurface
  Else
    Set dds = dx_DirectDrawStaticSurface(StaticSurface)
    
    FullWindow = True
  End If
  
  With dds
    .SetForeColor ForeColour
    .SetFontBackColor BackColour
    .SetFontTransparency Transparency
    
    If FullWindow Then
      .DrawText xPos, yPos, textString, False
    Else
      .DrawText xPos + m_AnimationRectangleX, yPos + m_AnimationRectangleY, textString, False
    End If
  End With
End Sub

Public Sub DrawBox(ByVal xPos As Long, ByVal yPos As Long, ByVal boxWidth As Long, ByVal boxHeight As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal fillColour As Long = 0, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal fillStyle As FillStyles = FSNoFill, Optional ByVal StaticSurface As Long = 0, Optional ByVal FullWindow As Boolean = False)
  Dim dds As DirectDrawSurface7
  
  On Error Resume Next
  
  If StaticSurface = 0 Then
    Set dds = dx_DirectDrawBackSurface
  Else
    Set dds = dx_DirectDrawStaticSurface(StaticSurface)
    
    FullWindow = True
  End If
  
  With dds
    .SetForeColor lineColour
    .setDrawStyle lineStyle
    .setDrawWidth lineWidth
    .SetFillColor fillColour
    .SetFillStyle fillStyle
    
    If FullWindow Then
      .DrawBox xPos, yPos, xPos + boxWidth, yPos + boxHeight
    Else
      .DrawBox xPos + m_AnimationRectangleX, yPos + m_AnimationRectangleY, xPos + m_AnimationRectangleX + boxWidth, yPos + m_AnimationRectangleY + boxHeight
    End If
  End With
End Sub

Public Sub DrawEllipse(ByVal xPos As Long, ByVal yPos As Long, ByVal boundingWidth As Long, ByVal boundingHeight As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal fillColour As Long = 0, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal fillStyle As FillStyles = FSNoFill, Optional ByVal StaticSurface As Long = 0, Optional ByVal FullWindow As Boolean = False)
  Dim dds As DirectDrawSurface7
  
  On Error Resume Next
  
  If StaticSurface = 0 Then
    Set dds = dx_DirectDrawBackSurface
  Else
    Set dds = dx_DirectDrawStaticSurface(StaticSurface)
    
    FullWindow = True
  End If
  
  With dds
    .SetForeColor lineColour
    .setDrawStyle lineStyle
    .setDrawWidth lineWidth
    .SetFillColor fillColour
    .SetFillStyle fillStyle
    
    If FullWindow Then
      .DrawEllipse xPos, yPos, xPos + boundingWidth, yPos + boundingHeight
    Else
      .DrawEllipse xPos + m_AnimationRectangleX, yPos + m_AnimationRectangleY, xPos + m_AnimationRectangleX + boundingWidth, yPos + m_AnimationRectangleY + boundingHeight
    End If
  End With
End Sub

Public Sub DrawCircle(ByVal xPos As Long, ByVal yPos As Long, ByVal circleRadius As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal fillColour As Long = 0, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal fillStyle As FillStyles = FSNoFill, Optional ByVal StaticSurface As Long = 0, Optional ByVal FullWindow As Boolean = False)
  Dim dds As DirectDrawSurface7
  
  On Error Resume Next
  
  If StaticSurface = 0 Then
    Set dds = dx_DirectDrawBackSurface
  Else
    Set dds = dx_DirectDrawStaticSurface(StaticSurface)
    
    FullWindow = True
  End If
  
  With dds
    .SetForeColor lineColour
    .setDrawStyle lineStyle
    .setDrawWidth lineWidth
    .SetFillColor fillColour
    .SetFillStyle fillStyle
    
    If FullWindow Then
      .DrawCircle xPos, yPos, circleRadius
    Else
      .DrawCircle xPos + m_AnimationRectangleX, yPos + m_AnimationRectangleY, circleRadius
    End If
  End With
End Sub

Public Sub DrawLine(ByVal xPosStart As Long, ByVal yPosStart As Long, ByVal xPosStop As Long, ByVal yPosStop As Long, ByVal lineColour As Long, Optional ByVal lineWidth As Long = 1, Optional ByVal lineStyle As LineStyles = LSSolid, Optional ByVal StaticSurface As Long = 0, Optional ByVal FullWindow As Boolean = False)
  Dim dds As DirectDrawSurface7
  
  On Error Resume Next
  
  If StaticSurface = 0 Then
    Set dds = dx_DirectDrawBackSurface
  Else
    Set dds = dx_DirectDrawStaticSurface(StaticSurface)
    
    FullWindow = True
  End If
  
  With dds
    .SetForeColor lineColour
    .setDrawStyle lineStyle
    .setDrawWidth lineWidth
    
    If FullWindow Then
      .DrawLine xPosStart, yPosStart, xPosStop, yPosStop
    Else
      .DrawLine xPosStart + m_AnimationRectangleX, yPosStart + m_AnimationRectangleY, xPosStop + m_AnimationRectangleX, yPosStop + m_AnimationRectangleY
    End If
  End With
End Sub

Public Sub SetAnimationWindow(Optional ByVal xPos As Long = 0, Optional ByVal yPos As Long = 0, Optional ByVal Width As Long = 0, Optional ByVal Height As Long = 0)
  If dx_FullScreenMode Then
    If Width = 0 Or Height = 0 Or xPos < 0 Or xPos > dx_Width Or yPos < 0 Or yPos > dx_Height Then
      m_AnimationRectangleX = 0
      m_AnimationRectangleY = 0
      m_AnimationRectangleWidth = dx_Width
      m_AnimationRectangleHeight = dx_Height
    Else
      m_AnimationRectangleX = xPos
      m_AnimationRectangleY = yPos
      m_AnimationRectangleWidth = Width
      m_AnimationRectangleHeight = Height
    End If
  Else
    If Width = 0 Or Height = 0 Or xPos < 0 Or xPos > dx_Width Or yPos < 0 Or yPos > dx_Height Then
      m_AnimationRectangleX = 0
      m_AnimationRectangleY = 0
      m_AnimationRectangleWidth = dx_Width
      m_AnimationRectangleHeight = dx_Height
    Else
      m_AnimationRectangleX = xPos
      m_AnimationRectangleY = yPos
      m_AnimationRectangleWidth = Width
      m_AnimationRectangleHeight = Height
    End If
  End If
End Sub

Private Sub UserControl_Terminate()
  Cleanup_AnimationWindow
End Sub

'****************************************************************
'
'Animation and map control functions
'
'****************************************************************

'Initializes the background map's image surface using the static image surface
Public Sub InitializeBGMapImages(ByVal StaticSurfaceIndex, ByVal ImageWidth As Long, _
      ByVal ImageHeight As Long, Optional ByVal ImagesPerRow As Long = 0)
  
  On Error Resume Next
  
  If StaticSurfaceIndex <= 0 Or StaticSurfaceIndex > m_TotalStaticSurfaces Then
    m_BGMapImageStaticSurface = 0
    
    Exit Sub
  End If
  
  m_BGMapImageStaticSurface = StaticSurfaceIndex
  
  m_BGMapImageWidth = ImageWidth
  m_BGMapImageHeight = ImageHeight
  m_BGMapImageSurfaceWidth = dx_StaticSurfaceWidth(StaticSurfaceIndex)
  m_BGMapImageSurfaceHeight = dx_StaticSurfaceHeight(StaticSurfaceIndex)
  m_BGMapImageStaticSurface = StaticSurfaceIndex
  
  If ImagesPerRow = 0 Then
    m_BGMapImagesPerRow = dx_StaticSurfaceWidth(StaticSurfaceIndex) \ ImageWidth
  Else
    m_BGMapImagesPerRow = ImagesPerRow
  End If
End Sub

Public Sub InitializeBGMap(BGMap() As Variant)
  Dim BGMapLine() As Variant
  Dim BGWLower As Long, BGWUpper As Long
  Dim BGHLower As Long, BGHUpper As Long
  Dim Hline As Long, Wline As Long
  
  On Error GoTo baderror
  
  BGHLower = LBound(BGMap)
  BGHUpper = UBound(BGMap)
  
  BGMapLine = BGMap(BGHLower)
  BGWLower = LBound(BGMapLine)
  BGWUpper = UBound(BGMapLine)
  
  m_BGMapHeight = BGHUpper - BGHLower + 1
  m_BGMapWidth = BGWUpper - BGWLower + 1
  
  ReDim m_BGMapArray(0 To m_BGMapHeight - 1, 0 To m_BGMapWidth - 1)
  
  For Hline = 0 To m_BGMapHeight - 1
    BGMapLine = BGMap(Hline + BGHLower)
    
    For Wline = 0 To m_BGMapWidth - 1
      m_BGMapArray(Hline, Wline) = BGMapLine(Wline + BGWLower)
    Next Wline
  Next Hline
baderror:
End Sub

Public Property Let BGMapElement(ByVal Row As Long, ByVal Column As Long, ByVal BGImage As Long)
  On Error Resume Next
  
  m_BGMapArray(Row, Column) = BGImage
End Property

Public Property Get BGMapElement(ByVal Row As Long, ByVal Column As Long) As Long
  On Error Resume Next
  
  BGMapElement = m_BGMapArray(Row, Column)
End Property

Public Property Get BGMapWidth() As Long
  BGMapWidth = m_BGMapWidth
End Property

Public Property Get BGMapHeight() As Long
  BGMapHeight = m_BGMapHeight
End Property

'Initializes the foreground map's image surface using the static image surface
Public Sub InitializeFGMapImages(ByVal StaticSurfaceIndex, ByVal ImageWidth As Long, _
      ByVal ImageHeight As Long, Optional ByVal ImagesPerRow As Long = 0)
  
  On Error Resume Next
  
  If StaticSurfaceIndex <= 0 Or StaticSurfaceIndex > m_TotalStaticSurfaces Then
    m_FGMapImageStaticSurface = 0
    
    Exit Sub
  End If
  
  m_FGMapImageStaticSurface = StaticSurfaceIndex
  
  m_FGMapImageWidth = ImageWidth
  m_FGMapImageHeight = ImageHeight
  m_FGMapImageSurfaceWidth = dx_StaticSurfaceWidth(StaticSurfaceIndex)
  m_FGMapImageSurfaceHeight = dx_StaticSurfaceHeight(StaticSurfaceIndex)
  m_FGMapImageStaticSurface = StaticSurfaceIndex
  
  If ImagesPerRow = 0 Then
    m_FGMapImagesPerRow = dx_StaticSurfaceWidth(StaticSurfaceIndex) \ ImageWidth
  Else
    m_FGMapImagesPerRow = ImagesPerRow
  End If
End Sub

Public Sub InitializeFGMap(FGMap() As Variant)
  Dim FGMapLine() As Variant
  Dim FGWLower As Long, FGWUpper As Long
  Dim FGHLower As Long, FGHUpper As Long
  Dim Hline As Long, Wline As Long
  
  On Error GoTo baderror
  
  FGHLower = LBound(FGMap)
  FGHUpper = UBound(FGMap)
  
  FGMapLine = FGMap(FGHLower)
  FGWLower = LBound(FGMapLine)
  FGWUpper = UBound(FGMapLine)
  
  m_FGMapHeight = FGHUpper - FGHLower + 1
  m_FGMapWidth = FGWUpper - FGWLower + 1
  
  ReDim m_FGMapArray(0 To m_FGMapHeight - 1, 0 To m_FGMapWidth - 1)
  
  For Hline = 0 To m_FGMapHeight - 1
    FGMapLine = FGMap(Hline + FGHLower)
    
    For Wline = 0 To m_FGMapWidth - 1
      m_FGMapArray(Hline, Wline) = FGMapLine(Wline + FGWLower)
    Next Wline
  Next Hline
baderror:
End Sub

Public Property Let FGMapElement(ByVal Row As Long, ByVal Column As Long, ByVal FGImage As Long)
  On Error Resume Next
  
  m_FGMapArray(Row, Column) = FGImage
End Property

Public Property Get FGMapElement(ByVal Row As Long, ByVal Column As Long) As Long
  On Error Resume Next
  
  FGMapElement = m_FGMapArray(Row, Column)
End Property

Public Property Get FGMapWidth() As Long
  FGMapWidth = m_FGMapWidth
End Property

Public Property Get FGMapHeight() As Long
  FGMapHeight = m_FGMapHeight
End Property

Public Property Get FirstObject() As AnimationObject
  Set FirstObject = m_FirstAnimationObject
End Property

Public Property Get LastObject() As AnimationObject
  Set LastObject = m_LastAnimationObject
End Property

'Add new object to animation window using static image block
Public Function AnimationListObjectAdd(StaticSurfaceIndex As Long, _
      ByVal ImageWidth As Long, ByVal ImageHeight As Long, Optional ByVal ImagesPerRow As Long = 0, _
      Optional ByVal SourceXOffset As Long = 0, Optional ByVal SourceYOffset As Long = 0, _
      Optional ByVal ObjectID As Long = 0, Optional ByVal ObjectPriority = 0) As AnimationObject
      
  Dim cObject As AnimationObject
  
  On Error GoTo badObject
  
  If StaticSurfaceIndex <= 0 Or StaticSurfaceIndex > m_TotalStaticSurfaces Then
    Set AnimationListObjectAdd = Nothing
    
    Exit Function
  End If
  
  If ObjectID = 0 Then
    Randomize
    
    Do While ObjectID = 0
      Set cObject = m_FirstAnimationObject
      
      ObjectID = (Rnd() * 100000) + 1
      
      Do While Not (cObject Is Nothing)
        If cObject.ObjectID = ObjectID Then
          ObjectID = 0
          
          Exit Do
        End If
        
        Set cObject = cObject.NextObject
      Loop
    Loop
  End If
  
  Set AnimationListObjectAdd = New AnimationObject
  
  With AnimationListObjectAdd
    .ObjectID = ObjectID
    .LayerPriority = ObjectPriority
    
    Set .DXSurface = dx_DirectDrawStaticSurface(StaticSurfaceIndex)
    
    .ActionFrame = -1
    .ImageWidth = ImageWidth
    .ImageHeight = ImageHeight
    .SurfaceWidth = dx_StaticSurfaceWidth(StaticSurfaceIndex)
    .SurfaceHeight = dx_StaticSurfaceHeight(StaticSurfaceIndex)
    .SourceOffsetX = SourceXOffset
    .SourceOffsetY = SourceYOffset
    .SurfaceStaticNum = StaticSurfaceIndex
    
    If ImagesPerRow = 0 Then
      .ImagesPerRow = dx_StaticSurfaceWidth(StaticSurfaceIndex) \ ImageWidth
    Else
      .ImagesPerRow = ImagesPerRow
    End If
  End With
  
  'find new home
  Set cObject = m_FirstAnimationObject
  
  If (cObject Is Nothing) Then
    Set m_FirstAnimationObject = AnimationListObjectAdd
    Set m_LastAnimationObject = AnimationListObjectAdd
  ElseIf cObject.LayerPriority <= ObjectPriority Then
    Set AnimationListObjectAdd.NextObject = cObject
    Set cObject.PreviousObject = AnimationListObjectAdd
    Set m_FirstAnimationObject = AnimationListObjectAdd
  Else
    Do While Not (cObject.NextObject Is Nothing)
      Set cObject = cObject.NextObject
      
      If cObject.LayerPriority <= ObjectPriority Then
        Set AnimationListObjectAdd.NextObject = cObject
        Set AnimationListObjectAdd.PreviousObject = cObject.PreviousObject
        Set cObject.PreviousObject = AnimationListObjectAdd
        
        If Not (AnimationListObjectAdd.PreviousObject Is Nothing) Then Set AnimationListObjectAdd.PreviousObject.NextObject = AnimationListObjectAdd
        
        Exit Function
      End If
    Loop
    
    Set AnimationListObjectAdd.PreviousObject = cObject
    Set cObject.NextObject = AnimationListObjectAdd
    Set m_LastAnimationObject = AnimationListObjectAdd
  End If
  
  Exit Function
  
badObject:
  AnimationListObjectAdd = Nothing
End Function

'Add duplicate object to animation window
Public Function AnimationListObjectDuplicate(SourceObject As AnimationObject, Optional ByVal ObjectID As Long = 0, Optional ByVal ObjectPriority = 0) As AnimationObject
  Dim cObject As AnimationObject, iDDSurfaceDesc As DirectDrawSurface7
  
  If SourceObject Is Nothing Then Exit Function
  
  If ObjectID = 0 Then
    Randomize
    
    Do While ObjectID = 0
      Set cObject = m_FirstAnimationObject
      
      ObjectID = (Rnd() * 100000) + 1
      
      Do While Not (cObject Is Nothing)
        If cObject.ObjectID = ObjectID Then
          ObjectID = 0
          
          Exit Do
        End If
        
        Set cObject = cObject.NextObject
      Loop
    Loop
  End If
  
  Set AnimationListObjectDuplicate = New AnimationObject
  
  With AnimationListObjectDuplicate
    .ObjectID = ObjectID
    .LayerPriority = ObjectPriority
    
    Set .DXSurface = SourceObject.DXSurface
    
    If SourceObject.ActionFrame = -1 Then .ActionFrame = -1
    .ImageWidth = SourceObject.ImageWidth
    .ImageHeight = SourceObject.ImageHeight
    .SurfaceWidth = SourceObject.SurfaceWidth
    .SurfaceHeight = SourceObject.SurfaceHeight
    .ImagesPerRow = SourceObject.ImagesPerRow
    .CollisionBoxBottom = SourceObject.CollisionBoxBottom
    .CollisionBoxLeft = SourceObject.CollisionBoxLeft
    .CollisionBoxTop = SourceObject.CollisionBoxTop
    .CollsionBoxRight = SourceObject.CollsionBoxRight
    .CollisionMaskMe = SourceObject.CollisionMaskMe
    .CollisionMaskTarget = SourceObject.CollisionMaskTarget
    .SourceOffsetX = SourceObject.SourceOffsetX
    .SourceOffsetY = SourceObject.SourceOffsetY
    .SurfaceStaticNum = SourceObject.SurfaceStaticNum
  End With
  
  'find new home
  Set cObject = m_FirstAnimationObject
  
  If (cObject Is Nothing) Then
    Set m_FirstAnimationObject = AnimationListObjectDuplicate
    Set m_LastAnimationObject = AnimationListObjectDuplicate
  ElseIf cObject.LayerPriority <= ObjectPriority Then
    Set AnimationListObjectDuplicate.NextObject = cObject
    Set cObject.PreviousObject = AnimationListObjectDuplicate
    Set m_FirstAnimationObject = AnimationListObjectDuplicate
  Else
    Do While Not (cObject.NextObject Is Nothing)
      Set cObject = cObject.NextObject
      
      If cObject.LayerPriority <= ObjectPriority Then
        Set AnimationListObjectDuplicate.NextObject = cObject
        Set AnimationListObjectDuplicate.PreviousObject = cObject.PreviousObject
        Set cObject.PreviousObject = AnimationListObjectDuplicate
        
        If Not (AnimationListObjectDuplicate.PreviousObject Is Nothing) Then Set AnimationListObjectDuplicate.PreviousObject.NextObject = AnimationListObjectDuplicate
        
        Exit Function
      End If
    Loop
    
    Set AnimationListObjectDuplicate.PreviousObject = cObject
    Set cObject.NextObject = AnimationListObjectDuplicate
    Set m_LastAnimationObject = AnimationListObjectDuplicate
  End If
End Function

'Remove selected object from animation window
Public Sub AnimationListObjectRemove(ByVal ObjectID As Long)
  Dim m_object As AnimationObject
  
  Set m_object = AnimationListObject(ObjectID)
  
  If Not (m_object Is Nothing) Then
    If m_object.PreviousObject Is Nothing Then
      Set m_FirstAnimationObject = m_object.NextObject
    Else
      Set m_object.PreviousObject.NextObject = m_object.NextObject
    End If
    
    If m_object.NextObject Is Nothing Then
      Set m_LastAnimationObject = m_object.PreviousObject
    Else
      Set m_object.NextObject.PreviousObject = m_object.PreviousObject
    End If
    
    Set m_object.NextObject = Nothing
    Set m_object.PreviousObject = Nothing
    Set m_object.DXSurface = Nothing
    
    m_object.Release
  End If
End Sub

'Remove all objects from animation window
Public Sub AnimationListObjectRemoveAll()
  Do While Not (m_FirstAnimationObject Is Nothing)
    AnimationListObjectRemove m_FirstAnimationObject.ObjectID
  Loop
End Sub

'Turn off visibility for all objects in animation window
Public Sub AnimationListObjectDeactivateAll()
  Dim listObject As AnimationObject
  
  Set listObject = m_FirstAnimationObject
  
  Do While Not (listObject Is Nothing)
    listObject.Visible = False
    
    Set listObject = listObject.NextObject
  Loop
End Sub

'Adjust object priority in animation window
Public Property Let AnimationListObjectPriority(ByVal ObjectID As Long, ByVal ObjectPriority)
  Dim m_object As AnimationObject, t_Object As AnimationObject
  
  Set m_object = AnimationListObject(ObjectID)
  
  'remove from list
  If Not (m_object Is Nothing) Then
    If m_object.PreviousObject Is Nothing Then
      Set m_FirstAnimationObject = m_object.NextObject
    Else
      Set m_object.PreviousObject.NextObject = m_object.NextObject
    End If
    
    If m_object.NextObject Is Nothing Then
      Set m_LastAnimationObject = m_object.PreviousObject
    Else
      Set m_object.NextObject.PreviousObject = m_object.PreviousObject
    End If
    
    Set m_object.NextObject = Nothing
    Set m_object.PreviousObject = Nothing
    m_object.LayerPriority = ObjectPriority
    
    'find new home
    Set t_Object = m_FirstAnimationObject
    
    If (t_Object Is Nothing) Then
      Set m_FirstAnimationObject = m_object
      Set m_LastAnimationObject = m_object
    ElseIf t_Object.LayerPriority <= ObjectPriority Then
      Set m_object.NextObject = t_Object
      Set t_Object.PreviousObject = m_object
      Set m_FirstAnimationObject = m_object
    Else
      Do While Not (t_Object.NextObject Is Nothing)
        Set t_Object = t_Object.NextObject
        
        If t_Object.LayerPriority <= ObjectPriority Then
          Set m_object.NextObject = t_Object
          Set t_Object.PreviousObject = m_object
          
          Exit Property
        End If
      Loop
      
      Set m_object.PreviousObject = t_Object
      Set t_Object.NextObject = m_object
      Set m_LastAnimationObject = m_object
    End If
  End If
End Property

'Return the selected object from the animation window
Public Property Get AnimationListObject(ByVal ObjectID As Long) As AnimationObject
  Set AnimationListObject = m_FirstAnimationObject
  
  Do While Not (AnimationListObject Is Nothing)
    If AnimationListObject.ObjectID = ObjectID Then Exit Property
    
    Set AnimationListObject = AnimationListObject.NextObject
  Loop
End Property

'Redraw the background and the animation list
Public Sub ReDrawAnimationWindow()
  Dim m_object As AnimationObject, sSurface As DirectDrawSurface7
  Dim RowX As Long, RowY As Long, ImageXOffset As Long, ImageYOffset As Long
  Dim BGImage As Long, FGImage As Long, t_Rect As RECT, startCol As Long
  
  On Error Resume Next
  
  If Not dx_DirectDrawEnabled Then Exit Sub
  
  'render background map (only if in solid mode)
  If MapMode And SolidBGMap Then
    Set sSurface = dx_DirectDrawStaticSurface(m_BGMapImageStaticSurface)
    
    RowY = 0
      
      Do While RowY < BGMapDisplayHeight
        RowX = 0
        
        Do While RowX < BGMapDisplayWidth
          BGImage = m_BGMapArray(RowY + BGMapStartRow, RowX + BGMapStartColumn)
          
          If BGImage And MIUseBaseOffset Then BGImage = BGMapBaseImageIndex + (BGImage And 4095)
          
          ImageYOffset = (BGImage \ m_BGMapImagesPerRow) * m_BGMapImageHeight
          ImageXOffset = (BGImage Mod m_BGMapImagesPerRow) * m_BGMapImageWidth
          
          BlitSolid sSurface, BGMapShiftX + RowX * m_BGMapImageWidth, BGMapShiftY + RowY * m_BGMapImageHeight, ImageXOffset, ImageYOffset, m_BGMapImageWidth, m_BGMapImageHeight
          
          RowX = RowX + 1
        Loop
        
        RowY = RowY + 1
      Loop
  Else 'either clear background or redraw background picture
    If m_BGPictureStaticSurface = 0 Then
      With t_Rect
        .Left = m_AnimationRectangleX
        .Top = m_AnimationRectangleY
        .Right = m_AnimationRectangleWidth + .Left
        .Bottom = m_AnimationRectangleHeight + .Top
      End With
      
      dx_DirectDrawBackSurface.BltColorFill t_Rect, m_BackColour
    ElseIf BGPictureWrap Then
      Set sSurface = dx_DirectDrawStaticSurface(m_BGPictureStaticSurface)
      
      RowY = BGPictureShiftY Mod m_BGPictureHeight
      If RowY <> 0 Then RowY = RowY - m_BGPictureHeight
      startCol = BGPictureShiftX Mod m_BGPictureWidth
      If startCol <> 0 Then startCol = startCol - m_BGPictureWidth
      
      Do While RowY < m_AnimationRectangleHeight
        RowX = startCol
        
        Do While RowX < m_AnimationRectangleWidth
          BlitSolid sSurface, RowX, RowY, m_BGPictureSourceX, m_BGPictureSourceY, m_BGPictureWidth, m_BGPictureHeight
           
          RowX = RowX + m_BGPictureWidth
        Loop
        
        RowY = RowY + m_BGPictureHeight
      Loop
    Else
      BlitSolid dx_DirectDrawStaticSurface(m_BGPictureStaticSurface), BGPictureShiftX, BGPictureShiftY, m_BGPictureSourceX, m_BGPictureSourceY, m_BGPictureWidth, m_BGPictureHeight
    End If
    
    'render background map (only if in transparent mode)
    If MapMode And TransparentBGMap Then
      Set sSurface = dx_DirectDrawStaticSurface(m_BGMapImageStaticSurface)
      
      RowY = 0
      
      Do While RowY < BGMapDisplayHeight
        RowX = 0
        
        Do While RowX < BGMapDisplayWidth
          BGImage = m_BGMapArray(RowY + BGMapStartRow, RowX + BGMapStartColumn)
          
          If BGImage <> 0 Then
            If BGImage And MIUseBaseOffset Then BGImage = BGMapBaseImageIndex + (BGImage And 4095)
            
            ImageYOffset = (BGImage \ m_BGMapImagesPerRow) * m_BGMapImageHeight
            ImageXOffset = (BGImage Mod m_BGMapImagesPerRow) * m_BGMapImageWidth
            
            BlitTransparent sSurface, BGMapShiftX + RowX * m_BGMapImageWidth, BGMapShiftY + RowY * m_BGMapImageHeight, ImageXOffset, ImageYOffset, m_BGMapImageWidth, m_BGMapImageHeight
          End If
          
          RowX = RowX + 1
        Loop
        
        RowY = RowY + 1
      Loop
    End If
  End If
  
  'render animation objects
  Set m_object = m_LastAnimationObject
  
  Do While Not (m_object Is Nothing)
    With m_object
      If .Visible Then
        If .ActionFrame = -1 Then
          Select Case .SpecialFX
            Case BFXNoEffects, BFXTransparent
              BlitTransparent .DXSurface, .PosX_1000ths \ 1000, .PosY_1000ths \ 1000, .SourceOffsetX, .SourceOffsetY, .ImageWidth, .ImageHeight
            Case BFXSolid
              BlitSolid .DXSurface, .PosX_1000ths \ 1000, .PosY_1000ths \ 1000, .SourceOffsetX, .SourceOffsetY, .ImageWidth, .ImageHeight
            Case Else
              BlitFX .DXSurface, .PosX_1000ths \ 1000, .PosY_1000ths \ 1000, .SourceOffsetX, .SourceOffsetY, .ImageWidth, .ImageHeight, .SpecialFX, .FXScaleWidth, .FXScaleHeight, .FXTargetColour
          End Select
        Else
          ImageXOffset = .ActionSequenceFrame(.ActionFrame)
          ImageYOffset = (ImageXOffset \ .ImagesPerRow) * .ImageHeight
          ImageXOffset = (ImageXOffset Mod .ImagesPerRow) * .ImageWidth
          
          Select Case .SpecialFX
            Case BFXNoEffects, BFXTransparent
              BlitTransparent .DXSurface, .PosX_1000ths \ 1000, .PosY_1000ths \ 1000, ImageXOffset, ImageYOffset, .ImageWidth, .ImageHeight
            Case BFXSolid
              BlitSolid .DXSurface, .PosX_1000ths \ 1000, .PosY_1000ths \ 1000, ImageXOffset, ImageYOffset, .ImageWidth, .ImageHeight
            Case Else
              BlitFX .DXSurface, .PosX_1000ths \ 1000, .PosY_1000ths \ 1000, ImageXOffset, ImageYOffset, .ImageWidth, .ImageHeight, .SpecialFX, .FXScaleWidth, .FXScaleHeight, .FXTargetColour
          End Select
        End If
      End If
    End With
    
    Set m_object = m_object.PreviousObject
  Loop
  
  'render foreground map
  If MapMode And TransparentFGMap Then
    Set sSurface = dx_DirectDrawStaticSurface(m_FGMapImageStaticSurface)
    
    RowY = 0
    
    Do While RowY < FGMapDisplayHeight
      RowX = 0
      
      Do While RowX < FGMapDisplayWidth
        FGImage = m_FGMapArray(RowY + FGMapStartRow, RowX + FGMapStartColumn)
        
        If FGImage <> 0 Then
          If FGImage And MIUseBaseOffset Then FGImage = FGMapBaseImageIndex + (FGImage And 4095)
          
          ImageYOffset = (FGImage \ m_FGMapImagesPerRow) * m_FGMapImageHeight
          ImageXOffset = (FGImage Mod m_FGMapImagesPerRow) * m_FGMapImageWidth
          
          BlitTransparent sSurface, FGMapShiftX + RowX * m_FGMapImageWidth, FGMapShiftY + RowY * m_FGMapImageHeight, ImageXOffset, ImageYOffset, m_FGMapImageWidth, m_FGMapImageHeight
        End If
        
        RowX = RowX + 1
      Loop
      
      RowY = RowY + 1
    Loop
  End If
End Sub

Public Property Get StaticSurfaceFileName(ByVal SurfaceIndex As Long) As String
  If SurfaceIndex > 0 And SurfaceIndex <= m_TotalStaticSurfaces Then
    If m_StaticSurfaceValid(SurfaceIndex) Then
      StaticSurfaceFileName = m_StaticSurfaceFileName(SurfaceIndex)
    Else
      StaticSurfaceFileName = ""
    End If
  Else
    StaticSurfaceFileName = ""
  End If
End Property

Public Sub SetBGPicture(Optional ByVal StaticSurfaceIndex As Long = 0, _
        Optional ByVal pictureWidth As Long = 0, Optional ByVal pictureHeight As Long = 0, _
        Optional ByVal SourceOffsetX As Long = 0, Optional ByVal SourceOffsetY As Long = 0)
        
  If StaticSurfaceIndex > 0 And StaticSurfaceIndex <= m_TotalStaticSurfaces Then
    m_BGPictureStaticSurface = StaticSurfaceIndex
    
    If pictureWidth <= 0 Then
      m_BGPictureWidth = dx_StaticSurfaceWidth(StaticSurfaceIndex)
    Else
      m_BGPictureWidth = pictureWidth
    End If
    
    If pictureHeight <= 0 Then
      m_BGPictureHeight = dx_StaticSurfaceHeight(StaticSurfaceIndex)
    Else
      m_BGPictureHeight = pictureHeight
    End If
    
    m_BGPictureSourceX = SourceOffsetX
    m_BGPictureSourceY = SourceOffsetY
  Else
    m_BGPictureStaticSurface = 0
  End If
End Sub

Public Property Get BGPictureStaticSurfaceNum() As Long
  BGPictureStaticSurfaceNum = m_BGPictureStaticSurface
End Property

Public Property Get BGMapStaticSurfaceNum() As Long
  BGMapStaticSurfaceNum = m_BGMapImageStaticSurface
End Property

Public Property Get FGMapStaticSurfaceNum() As Long
  FGMapStaticSurfaceNum = m_FGMapImageStaticSurface
End Property

Public Property Get StaticSurfaceValid(ByVal SurfaceIndex As Long) As Boolean
  StaticSurfaceValid = m_StaticSurfaceValid(SurfaceIndex)
End Property

Public Sub Cleanup_AnimationWindow()
  Dim loop1 As Long
  
  On Error Resume Next
  
  If dx_FullScreenMode Then dx_DirectDraw.RestoreDisplayMode
  
  dx_DirectDrawEnabled = False
  dx_FullScreenMode = False
  dx_Width = 0
  dx_Height = 0
  dx_BitDepth = 0
  
  AnimationListObjectRemoveAll
  
  Set dx_DirectDrawPrimarySurface = Nothing
  Set dx_DirectDrawPrimaryPalette = Nothing
  Set dx_DirectDrawPrimaryColourControl = Nothing
  Set dx_DirectDrawPrimaryGammaControl = Nothing
  Set dx_DirectDrawBackSurface = Nothing
  Set dx_DirectDrawFadeSurface = Nothing
  
  For loop1 = 1 To m_TotalStaticSurfaces
    Set dx_DirectDrawStaticSurface(loop1) = Nothing
    
    m_StaticSurfaceFileName(loop1) = ""
    m_StaticSurfaceValid(loop1) = False
  Next loop1
  
  m_TotalStaticSurfaces = 0
  
  MapMode = NoBGorFGMap
  
  m_BGMapImageSurfaceWidth = 0
  m_BGMapImageSurfaceHeight = 0
  m_BGMapImageWidth = 0
  m_BGMapImageHeight = 0
  m_BGMapImagesPerRow = 0
  
  BGMapBaseImageIndex = 0
  m_BGMapImageStaticSurface = 0
  
  BGMapShiftX = 0
  BGMapShiftY = 0
  BGMapDisplayWidth = 0
  BGMapDisplayHeight = 0
  BGMapStartRow = 0
  BGMapStartColumn = 0
  m_BGMapWidth = 0
  m_BGMapHeight = 0
  
  m_FGMapImageSurfaceWidth = 0
  m_FGMapImageSurfaceHeight = 0
  m_FGMapImageWidth = 0
  m_FGMapImageHeight = 0
  m_FGMapImagesPerRow = 0
  
  FGMapBaseImageIndex = 0
  m_FGMapImageStaticSurface = 0
  
  FGMapShiftX = 0
  FGMapShiftY = 0
  FGMapDisplayWidth = 0
  FGMapDisplayHeight = 0
  FGMapStartRow = 0
  FGMapStartColumn = 0
  m_FGMapWidth = 0
  m_FGMapHeight = 0
  
  m_BGPictureStaticSurface = 0
  BGPictureShiftX = 0
  BGPictureShiftY = 0
  
  m_BackColour = 0
  
  Set m_ClippingWindow = Nothing
  Set dx_DirectDraw = Nothing
End Sub
  


