Attribute VB_Name = "modMain"
Option Explicit

Public GameSpeed As Long, TimeLoop As Long, GameCounter As Long, GameCounter2 As Long

Public ScreenSizeMode As Long, VerifyExit As Boolean, PauseModeX As Boolean
Public LastKey As Long

Public GameMode As Long, GameSubMode As Long, PauseMode As Boolean
Public Difficulty As Long, PlayerActive As Long, activeController As Long
Public BonusMode As Long, WraithMode As Boolean, SpeedMode As Boolean, PowerUpMode As Boolean
Public TreatMode As Long, EatGhostCount As Long, GhostFastMode(1 To 5) As Boolean, SnackmanFastMode As Boolean
Public SmartDirection As Long, maxLogicDepth As Long, logicTargetX As Long, logicTargetY As Long

Public m_Title(1 To 3) As AnimationObject, m_Title2(1 To 9) As AnimationObject
Public m_SnackMan As AnimationObject, m_Ghost(1 To 5) As AnimationObject, m_Treat(1 To 12) As AnimationObject
Public m_Title3(1 To 12) As AnimationObject, m_GhostT(1 To 5) As AnimationObject
Public m_PowerUp(1 To 4) As AnimationObject, m_Scores(1 To 5) As AnimationObject

Public m_SoundActive(1 To 20) As Boolean, m_SoundPaused As Boolean, LastDirection As Long

Public Type PData
  Map(0 To 7) As Variant
  MapLevel As Long
  
  SuperDot(1 To 4, 1 To 2) As Long
  SuperDotTime As Long
  
  DotCount As Long
  DotValue As Long
  
  CompletionBonus As Long
  Lives As Long
  NextMan As Long
  Score As Long
  
  GhostIntel(1 To 5) As Long
  GhostBaseSpeed(1 To 5) As Long
  WraithTimeOut As Long
  WraithTime As Long
  
  BestTreat As Long
  TreatTimeOut As Long
  TreatTime As Long
  
  SpeedTime As Long
  GameSpeed As Long
  
  GhostHomes(1 To 4, 1 To 2) As Long
  SnackManStart(1 To 2) As Long
End Type

Public PlayerData(1 To 2) As PData

Public Enum CollisionType
  CMSnackman = 1
  CMGhost = 2
  CMTreats = 4
  CMWalls = 8
End Enum

Sub Main()
  Dim loop1 As Long
  Dim green(0 To 255) As Long, red(0 To 255) As Long, blue(0 To 255) As Long, count As Long
  
  Difficulty = 1
  activeController = 1
  GameMode = 0
  PauseMode = False
  m_SoundPaused = False
  GameSpeed = 25
  ScreenSizeMode = 1
  
  frmMain.Show
  
  Do While GameMode >= 0
    TimeLoop = DelayTillTime(TimeLoop, 25) + 25
    
    If PauseMode = True And PauseModeX = True Then
      PollKeyboard
      Read_Keyboard
      PollJoystick
      Read_Joystick
      
      If LastKey = 0 Then
        For loop1 = 1 To 255
          If dx_KeyboardState.Key(loop1) <> 0 Then
            PauseMode = False
            UnPauseAllSound
            
            LastKey = 255
            
            Exit For
          End If
        Next loop1
        
        If LastKey <> 255 And dx_JoystickState.buttons(0) <> 0 Then
          PauseMode = False
          UnPauseAllSound
            
          LastKey = 255
        End If
      Else
        LastKey = 0
        
        For loop1 = 1 To 255
          If dx_KeyboardState.Key(loop1) <> 0 Then
            LastKey = loop1
            
            Exit For
          End If
        Next loop1
        
        If dx_JoystickState.buttons(0) <> 0 Then LastKey = 255
      End If
    End If
    
    Do While Not PauseMode
      If GameMode = 6 Then
        TimeLoop = DelayTillTime(TimeLoop, 25, , False) + GameSpeed
      Else
        TimeLoop = DelayTillTime(TimeLoop, 25, , False) + 25
      End If
      
      If GameMode > 0 Then
        RefreshDisplay
        FlipBuffers
      End If
      
      If GameMode > 1 Then
        PollKeyboard
        Read_Keyboard
        
        If dx_KeyboardState.Key(DIK_P) <> 0 Then
         If LastKey <> DIK_P Then
          PauseMode = True
          PauseModeX = True
          LastKey = DIK_P
          PauseAllSound
          
          FadeDisplay , FMBlack
          DisplayStaticImageTransparent 1, 0, 0, 240, 96, 256, 222, True
            
          RefreshDisplay
          FlipBuffers
          
          Exit Do
         End If
        ElseIf dx_KeyboardState.Key(DIK_ESCAPE) <> 0 Then
         If LastKey <> DIK_ESCAPE Then
          PauseMode = True
          PauseModeX = False
          LastKey = DIK_ESCAPE
          PauseAllSound
          
          With frmMain
            If FullScreenMode Then
              Init_DXDrawWindow frmMain, frmMain.picWindow, , True
              
              .Width = 11985
              .Height = 8745
              
              TimeLoop = DelayTillTime(50, , True)
      
              ClearDisplay True
              
              .Display_Border
              .Display_Score 0, 0, 0, True
              .Display_Specials 0, False, False, 1, False, True
              
              ReDrawAnimationWindow
            End If
            
            FadeDisplay , FMBlack
            DisplayStaticImageTransparent 1, 0, 0, 240, 96, 256, 222, True
            
            RefreshDisplay
            
            .mnuGame.Visible = True
            .mnuDisplay.Visible = True
            .mnuControllerTitle.Visible = True
            
            .mnuScreenmode(0).Checked = False
            .mnuScreenmode(1).Checked = False
            .mnuScreenmode(2).Checked = False
            .mnuScreenmode(ScreenSizeMode).Checked = True
            
            For loop1 = 0 To 10
              If loop1 <= dx_EnumJoysticks.GetCount() Then
                .mnuController(loop1).Visible = True
                
                If loop1 > 0 Then .mnuController.Item(loop1).Caption = dx_EnumJoysticks.GetItem(loop1).GetInstanceName
                
                If activeController = loop1 Then
                  .mnuController(loop1).Checked = True
                Else
                  .mnuController(loop1).Checked = False
                End If
              Else
                .mnuController(loop1).Visible = False
              End If
            Next loop1
          End With
          
          CleanUp_DXInput
          
          Exit Do
         End If
        ElseIf dx_KeyboardState.Key(DIK_F3) <> 0 Then
         If LastKey <> DIK_F3 Then
          LastKey = DIK_F3
          
          activeController = activeController + 1
          
          If activeController > dx_EnumJoysticks.GetCount Then activeController = 0
          
          Select_Joystick frmMain, activeController, True
          
          If FullScreenMode Then
            ClearDisplay True
            
            frmMain.Display_Border
            frmMain.Display_Score 0, 0, 0
            frmMain.Display_Specials 0, False, False, 1, False
          End If
         End If
        ElseIf dx_KeyboardState.Key(DIK_F4) <> 0 Then
         If LastKey <> DIK_F4 Then
          LastKey = DIK_F4
          
          If ScreenSizeMode <> 1 Then
            ScreenSizeMode = 3
            
            With frmMain
              Init_DXDrawScreen frmMain, 800, 600, 8, 60, , True
              
              TimeLoop = DelayTillTime(50, , True)
              
              ClearDisplay True
              
              .Display_Border
              .Display_Score 0, 0, 0, True
              .Display_Specials 0, False, False, 1, False, True
              
              .Reaquire_DXInput
            End With
          End If
         End If
        ElseIf dx_KeyboardState.Key(DIK_F8) <> 0 Then
         If LastKey <> DIK_F8 Then
          LastKey = DIK_F8
          
          If ScreenSizeMode <> 1 Then
            ScreenSizeMode = 1
            
            With frmMain
              Init_DXDrawScreen frmMain, 800, 600, 16, 60, , True
              
              TimeLoop = DelayTillTime(50, , True)
              
              ClearDisplay True
              
              .Display_Border
              .Display_Score 0, 0, 0, True
              .Display_Specials 0, False, False, 1, False, True
              
              .Reaquire_DXInput
            End With
          End If
         End If
        ElseIf dx_KeyboardState.Key(DIK_F9) <> 0 Then
         If LastKey <> DIK_F9 Then
          LastKey = DIK_F9
          
          If ScreenSizeMode <> 2 Then
            ScreenSizeMode = 2
            
            With frmMain
              Init_DXDrawScreen frmMain, 800, 600, 32, 60, , True
              
              TimeLoop = DelayTillTime(50, , True)
              
              ClearDisplay True
              
              .Display_Border
              .Display_Score 0, 0, 0, True
              .Display_Specials 0, False, False, 1, False, True
              
              .Reaquire_DXInput
            End With
          End If
         End If
        ElseIf LastKey <> 255 Then
          LastKey = 0
        End If
        
        PollJoystick
        Read_Joystick
      End If
      
      Select Case GameMode
        Case 0 'still waiting for frmMain to initialize
          
        Case 1 'initialize DirectX
          frmMain.GameMode_1
        Case 2 'reset for intro routine
          frmMain.GameMode_2
        Case 3 'play intro routine
          frmMain.GameMode_3
        Case 4 'reset for new game - not called this way
          'frmMain.GameMode_4 numPlayers
        Case 5 'prepare map and wait for player go
          frmMain.GameMode_5
        Case 6 'main game routine
          frmMain.GameMode_6
        Case 7 'cleared level delay routine
          frmMain.GameMode_7
        Case 8 'died routine
          frmMain.GameMode_8
        Case 9 'game over
          frmMain.GameMode_9
      End Select
      
      If GameMode > 0 Then If Not PauseMode Then ReDrawAnimationWindow
    Loop
    
    If GameMode > 0 Then RefreshDisplay
  Loop
End Sub

Public Sub LoadMapLevel()
  Dim loop1 As Long, loop2 As Long, dcount As Long
  Dim lLimit As Long, skipdot As Boolean, cellVal As Long
  Dim mapLine As Variant, baseVal As Long, effMapLevel As Long
  
  With PlayerData(PlayerActive) 'generate map
    Select Case .MapLevel Mod 10
      Case 1
        .Map(0) = Array(0, 8, 5, 9, 0, 0, 0, 0, 8, 5, 9, 0)
        .Map(1) = Array(8, 10, 1, 4, 2, 5, 5, 2, 6, 1, 7, 9)
        .Map(2) = Array(4, 2, 10, 1, 0, 1, 1, 0, 1, 7, 2, 6)
        .Map(3) = Array(7, 5, 2, 6, 0, 1, 1, 0, 4, 2, 5, 10)
        .Map(4) = Array(96 Or MIUseBaseOffset, 6, 0, 4, 2, 3, 3, 2, 6, 0, 4, 100 Or MIUseBaseOffset)
        .Map(5) = Array(8, 6, 0, 4, 2, 2, 2, 2, 6, 0, 4, 9)
        .Map(6) = Array(1, 1, 0, 1, 104, 105, 105, 106, 1, 0, 1, 1)
        .Map(7) = Array(7, 3, 2, 3, 2, 2, 2, 2, 3, 2, 3, 10)
        
        .GhostHomes(1, 1) = 4
        .GhostHomes(1, 2) = 6
        .GhostHomes(2, 1) = 5
        .GhostHomes(2, 2) = 6
        .GhostHomes(3, 1) = 6
        .GhostHomes(3, 2) = 6
        .GhostHomes(4, 1) = 7
        .GhostHomes(4, 2) = 6
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 7
        
        .SuperDot(1, 1) = 1
        .SuperDot(1, 2) = 1
        .SuperDot(2, 1) = 10
        .SuperDot(2, 2) = 1
        .SuperDot(3, 1) = 0
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
      Case 2
        .Map(0) = Array(8, 2, 2, 5, 2, 2, 2, 2, 5, 2, 2, 9)
        .Map(1) = Array(4, 2, 9, 4, 2, 2, 2, 2, 6, 8, 2, 6)
        .Map(2) = Array(4, 2, 6, 1, 104, 105, 105, 106, 1, 4, 2, 6)
        .Map(3) = Array(7, 2, 3, 3, 5, 2, 2, 5, 3, 3, 2, 10)
        .Map(4) = Array(96 Or MIUseBaseOffset, 2, 5, 2, 3, 2, 2, 3, 2, 5, 2, 100 Or MIUseBaseOffset)
        .Map(5) = Array(8, 5, 6, 8, 2, 2, 2, 2, 9, 4, 5, 9)
        .Map(6) = Array(1, 1, 1, 4, 2, 2, 2, 2, 6, 1, 1, 1)
        .Map(7) = Array(7, 3, 3, 3, 2, 2, 2, 2, 3, 3, 3, 10)
        
        .GhostHomes(1, 1) = 4
        .GhostHomes(1, 2) = 2
        .GhostHomes(2, 1) = 5
        .GhostHomes(2, 2) = 2
        .GhostHomes(3, 1) = 6
        .GhostHomes(3, 2) = 2
        .GhostHomes(4, 1) = 7
        .GhostHomes(4, 2) = 2
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 3
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 0
        .SuperDot(2, 1) = 0
        .SuperDot(2, 2) = 7
        .SuperDot(3, 1) = 11
        .SuperDot(3, 2) = 0
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
      Case 3
        .Map(0) = Array(8, 5, 2, 2, 5, 9, 8, 5, 2, 2, 5, 9)
        .Map(1) = Array(1, 4, 2, 2, 6, 7, 10, 4, 2, 2, 6, 1)
        .Map(2) = Array(7, 10, 8, 5, 3, 2, 2, 3, 5, 9, 7, 10)
        .Map(3) = Array(96 Or MIUseBaseOffset, 5, 10, 1, 104, 105, 105, 106, 1, 7, 5, 100 Or MIUseBaseOffset)
        .Map(4) = Array(8, 3, 9, 4, 2, 2, 2, 2, 6, 8, 3, 9)
        .Map(5) = Array(1, 8, 10, 7, 5, 2, 2, 5, 10, 7, 9, 1)
        .Map(6) = Array(1, 1, 8, 2, 6, 0, 0, 4, 2, 9, 1, 1)
        .Map(7) = Array(7, 3, 10, 0, 7, 2, 2, 10, 0, 7, 3, 10)
        
        .GhostHomes(1, 1) = 4
        .GhostHomes(1, 2) = 3
        .GhostHomes(2, 1) = 5
        .GhostHomes(2, 2) = 3
        .GhostHomes(3, 1) = 6
        .GhostHomes(3, 2) = 3
        .GhostHomes(4, 1) = 7
        .GhostHomes(4, 2) = 3
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 4
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 0
        .SuperDot(2, 1) = 11
        .SuperDot(2, 2) = 0
        .SuperDot(3, 1) = 0
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
      Case 4
        .Map(0) = Array(96 Or MIUseBaseOffset, 2, 9, 8, 2, 2, 2, 2, 9, 8, 2, 100 Or MIUseBaseOffset)
        .Map(1) = Array(8, 2, 11, 11, 2, 2, 2, 2, 11, 11, 2, 9)
        .Map(2) = Array(7, 5, 10, 4, 2, 2, 2, 2, 6, 7, 5, 10)
        .Map(3) = Array(8, 11, 9, 1, 104, 105, 105, 106, 1, 8, 11, 9)
        .Map(4) = Array(1, 1, 1, 7, 5, 2, 2, 5, 10, 1, 1, 1)
        .Map(5) = Array(1, 1, 4, 2, 3, 2, 2, 3, 2, 6, 1, 1)
        .Map(6) = Array(7, 3, 11, 2, 2, 5, 5, 2, 2, 11, 3, 10)
        .Map(7) = Array(96 Or MIUseBaseOffset, 2, 3, 2, 2, 10, 7, 2, 2, 3, 2, 100 Or MIUseBaseOffset)
        
        .GhostHomes(1, 1) = 4
        .GhostHomes(1, 2) = 3
        .GhostHomes(2, 1) = 5
        .GhostHomes(2, 2) = 3
        .GhostHomes(3, 1) = 6
        .GhostHomes(3, 2) = 3
        .GhostHomes(4, 1) = 7
        .GhostHomes(4, 2) = 3
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 4
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 1
        .SuperDot(2, 1) = 0
        .SuperDot(2, 2) = 6
        .SuperDot(3, 1) = 11
        .SuperDot(3, 2) = 1
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 6
      Case 5
        .Map(0) = Array(8, 2, 2, 2, 2, 2, 9, 8, 2, 2, 2, 9)
        .Map(1) = Array(7, 9, 104, 105, 105, 106, 1, 4, 2, 2, 5, 10)
        .Map(2) = Array(96 Or MIUseBaseOffset, 11, 2, 2, 2, 5, 3, 10, 8, 2, 11, 100 Or MIUseBaseOffset)
        .Map(3) = Array(8, 3, 2, 5, 2, 11, 2, 2, 11, 2, 3, 9)
        .Map(4) = Array(1, 8, 2, 10, 8, 11, 5, 5, 11, 2, 2, 6)
        .Map(5) = Array(1, 7, 9, 8, 6, 1, 1, 1, 4, 2, 2, 10)
        .Map(6) = Array(1, 8, 10, 1, 1, 1, 1, 1, 4, 5, 2, 9)
        .Map(7) = Array(7, 3, 2, 3, 3, 3, 3, 3, 10, 7, 2, 10)
        
        .GhostHomes(1, 1) = 2
        .GhostHomes(1, 2) = 1
        .GhostHomes(2, 1) = 3
        .GhostHomes(2, 2) = 1
        .GhostHomes(3, 1) = 4
        .GhostHomes(3, 2) = 1
        .GhostHomes(4, 1) = 5
        .GhostHomes(4, 2) = 1
        
        .SnackManStart(1) = 2
        .SnackManStart(2) = 2
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 0
        .SuperDot(2, 1) = 11
        .SuperDot(2, 2) = 0
        .SuperDot(3, 1) = 0
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
      Case 6
        .Map(0) = Array(8, 2, 2, 9, 8, 9, 8, 9, 8, 2, 2, 9)
        .Map(1) = Array(7, 2, 2, 3, 6, 1, 1, 4, 3, 2, 2, 10)
        .Map(2) = Array(8, 2, 2, 9, 1, 1, 1, 1, 8, 2, 2, 9)
        .Map(3) = Array(1, 104, 106, 1, 1, 1, 1, 1, 1, 104, 106, 1)
        .Map(4) = Array(7, 2, 2, 6, 1, 1, 1, 1, 4, 2, 2, 10)
        .Map(5) = Array(96 Or MIUseBaseOffset, 2, 2, 11, 3, 3, 3, 3, 11, 2, 2, 100 Or MIUseBaseOffset)
        .Map(6) = Array(8, 2, 5, 3, 2, 9, 8, 2, 3, 5, 2, 9)
        .Map(7) = Array(7, 2, 10, 0, 0, 7, 10, 0, 0, 7, 2, 10)
        
        .GhostHomes(1, 1) = 1
        .GhostHomes(1, 2) = 3
        .GhostHomes(2, 1) = 2
        .GhostHomes(2, 2) = 3
        .GhostHomes(3, 1) = 9
        .GhostHomes(3, 2) = 3
        .GhostHomes(4, 1) = 10
        .GhostHomes(4, 2) = 3
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 5
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 0
        .SuperDot(2, 1) = 11
        .SuperDot(2, 2) = 0
        .SuperDot(3, 1) = 0
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
      Case 7
        .Map(0) = Array(96 Or MIUseBaseOffset, 5, 5, 2, 5, 2, 2, 5, 2, 5, 5, 100 Or MIUseBaseOffset)
        .Map(1) = Array(8, 6, 4, 2, 11, 2, 2, 11, 2, 6, 4, 9)
        .Map(2) = Array(1, 1, 1, 0, 4, 2, 2, 6, 0, 1, 1, 1)
        .Map(3) = Array(1, 1, 1, 8, 3, 2, 2, 3, 9, 1, 1, 1)
        .Map(4) = Array(4, 6, 1, 7, 2, 2, 2, 2, 10, 1, 4, 6)
        .Map(5) = Array(1, 1, 4, 5, 2, 2, 2, 2, 5, 6, 1, 1)
        .Map(6) = Array(1, 1, 1, 1, 104, 105, 105, 106, 1, 1, 1, 1)
        .Map(7) = Array(7, 3, 3, 3, 2, 2, 2, 2, 3, 3, 3, 10)
        
        .GhostHomes(1, 1) = 4
        .GhostHomes(1, 2) = 6
        .GhostHomes(2, 1) = 5
        .GhostHomes(2, 2) = 6
        .GhostHomes(3, 1) = 6
        .GhostHomes(3, 2) = 6
        .GhostHomes(4, 1) = 7
        .GhostHomes(4, 2) = 6
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 7
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 1
        .SuperDot(2, 1) = 11
        .SuperDot(2, 2) = 1
        .SuperDot(3, 1) = 0
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
      Case 8
        .Map(0) = Array(8, 2, 2, 5, 2, 5, 2, 5, 2, 5, 2, 9)
        .Map(1) = Array(1, 8, 5, 10, 8, 10, 8, 10, 8, 10, 8, 6)
        .Map(2) = Array(4, 10, 4, 2, 3, 2, 3, 2, 3, 2, 6, 1)
        .Map(3) = Array(1, 8, 3, 5, 2, 2, 2, 2, 2, 5, 6, 1)
        .Map(4) = Array(7, 6, 8, 3, 2, 2, 2, 2, 2, 6, 7, 10)
        .Map(5) = Array(8, 3, 11, 2, 2, 2, 2, 2, 2, 3, 2, 9)
        .Map(6) = Array(7, 9, 4, 2, 2, 5, 2, 2, 2, 2, 5, 10)
        .Map(7) = Array(96 Or MIUseBaseOffset, 10, 7, 2, 2, 10, 104, 105, 105, 106, 7, 100 Or MIUseBaseOffset)
        
        .GhostHomes(1, 1) = 6
        .GhostHomes(1, 2) = 7
        .GhostHomes(2, 1) = 7
        .GhostHomes(2, 2) = 7
        .GhostHomes(3, 1) = 8
        .GhostHomes(3, 2) = 7
        .GhostHomes(4, 1) = 9
        .GhostHomes(4, 2) = 7
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 3
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 0
        .SuperDot(2, 1) = 11
        .SuperDot(2, 2) = 0
        .SuperDot(3, 1) = 2
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 4
      Case 9
        .Map(0) = Array(8, 2, 2, 5, 2, 2, 9, 8, 9, 8, 5, 9)
        .Map(1) = Array(1, 104, 106, 1, 8, 2, 11, 10, 7, 6, 1, 1)
        .Map(2) = Array(4, 2, 2, 11, 6, 8, 3, 2, 2, 6, 1, 1)
        .Map(3) = Array(7, 2, 2, 10, 1, 1, 8, 5, 2, 3, 3, 10)
        .Map(4) = Array(96 Or MIUseBaseOffset, 5, 5, 2, 3, 10, 1, 1, 8, 2, 2, 100 Or MIUseBaseOffset)
        .Map(5) = Array(8, 6, 4, 2, 2, 5, 10, 4, 11, 2, 2, 9)
        .Map(6) = Array(1, 1, 4, 9, 8, 11, 2, 10, 1, 104, 106, 1)
        .Map(7) = Array(7, 3, 10, 7, 10, 7, 2, 2, 3, 2, 2, 10)
        
        .GhostHomes(1, 1) = 1
        .GhostHomes(1, 2) = 1
        .GhostHomes(2, 1) = 2
        .GhostHomes(2, 2) = 1
        .GhostHomes(3, 1) = 9
        .GhostHomes(3, 2) = 6
        .GhostHomes(4, 1) = 10
        .GhostHomes(4, 2) = 6
        
        .SnackManStart(1) = 7
        .SnackManStart(2) = 2
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 0
        .SuperDot(2, 1) = 11
        .SuperDot(2, 2) = 0
        .SuperDot(3, 1) = 0
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
      Case Else
        .Map(0) = Array(8, 2, 2, 5, 9, 8, 9, 8, 5, 2, 2, 9)
        .Map(1) = Array(4, 2, 2, 6, 1, 1, 1, 1, 4, 2, 2, 6)
        .Map(2) = Array(4, 2, 2, 6, 1, 1, 1, 1, 4, 2, 2, 6)
        .Map(3) = Array(7, 2, 2, 6, 7, 3, 3, 10, 4, 2, 2, 10)
        .Map(4) = Array(96 Or MIUseBaseOffset, 2, 5, 6, 104, 105, 105, 106, 4, 5, 2, 100 Or MIUseBaseOffset)
        .Map(5) = Array(8, 5, 6, 4, 2, 2, 2, 2, 6, 4, 5, 9)
        .Map(6) = Array(1, 1, 1, 1, 8, 2, 2, 9, 1, 1, 1, 1)
        .Map(7) = Array(7, 3, 10, 7, 3, 2, 2, 3, 10, 7, 3, 10)
        
        .GhostHomes(1, 1) = 4
        .GhostHomes(1, 2) = 4
        .GhostHomes(2, 1) = 5
        .GhostHomes(2, 2) = 4
        .GhostHomes(3, 1) = 6
        .GhostHomes(3, 2) = 4
        .GhostHomes(4, 1) = 7
        .GhostHomes(4, 2) = 4
        
        .SnackManStart(1) = 6
        .SnackManStart(2) = 5
        
        .SuperDot(1, 1) = 0
        .SuperDot(1, 2) = 0
        .SuperDot(2, 1) = 11
        .SuperDot(2, 2) = 0
        .SuperDot(3, 1) = 0
        .SuperDot(3, 2) = 7
        .SuperDot(4, 1) = 11
        .SuperDot(4, 2) = 7
    End Select
    
    For loop1 = 0 To 7 'add dots
      mapLine = .Map(loop1)
      lLimit = LBound(mapLine)
    
      For loop2 = 0 To 11
        skipdot = False
        
        For dcount = 1 To 4
          If .SuperDot(dcount, 1) = loop2 And .SuperDot(dcount, 2) = loop1 Then
            skipdot = True
            
            Exit For
          End If
        Next dcount
        
        If .SnackManStart(1) = loop2 And .SnackManStart(2) = loop1 Then skipdot = True
        
        If Not skipdot Then
          cellVal = mapLine(loop2)
          
          If cellVal > 0 And cellVal < 12 Then
            cellVal = cellVal + 12
            
            mapLine(loop2) = cellVal
          End If
        End If
      Next loop2
      
      .Map(loop1) = mapLine
    Next loop1
    
    baseVal = ((((.MapLevel + 1) \ 2) - 1) Mod 4) * 24
    
    .DotCount = 4
    
    For loop1 = 0 To 7 'add level colour
      mapLine = .Map(loop1)
      lLimit = LBound(mapLine)
    
      For loop2 = 0 To 11
        cellVal = mapLine(loop2)
        
        If cellVal > 0 And cellVal < 24 Then
          If cellVal > 11 Then .DotCount = .DotCount + 1
          
          cellVal = cellVal + baseVal
          
          mapLine(loop2) = cellVal
        End If
      Next loop2
      
      .Map(loop1) = mapLine
    Next loop1
    
    'level settings
    If .MapLevel > 10 Then
      .BestTreat = 11
    Else
      .BestTreat = .MapLevel + 1
    End If
    
    effMapLevel = .MapLevel + 4 * Difficulty
    
    .TreatTime = 40 * 8
    .TreatTimeOut = 15 * 8
    
    If effMapLevel > 20 Then
      .DotValue = 20
      .CompletionBonus = 1000
      
      .SuperDotTime = 10 * 8
      .WraithTimeOut = (15 - Difficulty * 5) * 8
      .WraithTime = (25 + Difficulty * 5) * 8
      .SpeedTime = 10 * 8
    Else
      .DotValue = effMapLevel
      .CompletionBonus = effMapLevel * 50
      
      .SuperDotTime = (20 - effMapLevel / 2) * 8
      .WraithTimeOut = (55 - Difficulty * 5 - (2 * effMapLevel)) * 8
      .WraithTime = (5 + Difficulty * 5 + effMapLevel) * 8
      .SpeedTime = (20 - effMapLevel / 2) * 8
    End If
    
    Select Case effMapLevel
        Case 1, 2
          .GhostBaseSpeed(1) = 1000 'dopy
          .GhostBaseSpeed(2) = 500 'poky
          .GhostBaseSpeed(3) = 1000 'brain
          .GhostBaseSpeed(4) = 2000 'speedy
          .GhostBaseSpeed(5) = 1 'wraith
          .GameSpeed = 30
          .GhostIntel(1) = 1
          .GhostIntel(2) = 2
          .GhostIntel(3) = 7
          .GhostIntel(4) = 2
          .GhostIntel(5) = 3
        Case 3, 4
          .GhostBaseSpeed(1) = 1600 'dopy
          .GhostBaseSpeed(2) = 1000 'poky
          .GhostBaseSpeed(3) = 1600 'brain
          .GhostBaseSpeed(4) = 2800 'speedy
          .GhostBaseSpeed(5) = 1400 'wraith
          .GameSpeed = 28
          .GhostIntel(1) = 1
          .GhostIntel(2) = 3
          .GhostIntel(3) = 8
          .GhostIntel(4) = 2
          .GhostIntel(5) = 4
        Case 5, 6
          .GhostBaseSpeed(1) = 2000 'dopy
          .GhostBaseSpeed(2) = 1400 'poky
          .GhostBaseSpeed(3) = 2000 'brain
          .GhostBaseSpeed(4) = 2800 'speedy
          .GhostBaseSpeed(5) = 1400 'wraith
          .GameSpeed = 25
          .GhostIntel(1) = 1
          .GhostIntel(2) = 4
          .GhostIntel(3) = 9
          .GhostIntel(4) = 3
          .GhostIntel(5) = 5
        Case 7, 8
          .GhostBaseSpeed(1) = 2000 'dopy
          .GhostBaseSpeed(2) = 1400 'poky
          .GhostBaseSpeed(3) = 2000 'brain
          .GhostBaseSpeed(4) = 4000 'speedy
          .GhostBaseSpeed(5) = 1600 'wraith
          .GameSpeed = 23
          .GhostIntel(1) = 1
          .GhostIntel(2) = 5
          .GhostIntel(3) = 10
          .GhostIntel(4) = 3
          .GhostIntel(5) = 6
        Case 9, 10
          .GhostBaseSpeed(1) = 2800 'dopy
          .GhostBaseSpeed(2) = 1600 'poky
          .GhostBaseSpeed(3) = 2800 'brain
          .GhostBaseSpeed(4) = 4000 'speedy
          .GhostBaseSpeed(5) = 1600 'wraith
          .GameSpeed = 22
          .GhostIntel(1) = 2
          .GhostIntel(2) = 5
          .GhostIntel(3) = 11
          .GhostIntel(4) = 3
          .GhostIntel(5) = 7
        Case 11, 12
          .GhostBaseSpeed(1) = 2800 'dopy
          .GhostBaseSpeed(2) = 1600 'poky
          .GhostBaseSpeed(3) = 2800 'brain
          .GhostBaseSpeed(4) = 4000 'speedy
          .GhostBaseSpeed(5) = 2000 'wraith
          .GameSpeed = 21
          .GhostIntel(1) = 2
          .GhostIntel(2) = 6
          .GhostIntel(3) = 12
          .GhostIntel(4) = 4
          .GhostIntel(5) = 8
        Case 13, 14
          .GhostBaseSpeed(1) = 4000 'dopy
          .GhostBaseSpeed(2) = 2000 'poky
          .GhostBaseSpeed(3) = 2800 'brain
          .GhostBaseSpeed(4) = 4000 'speedy
          .GhostBaseSpeed(5) = 2000 'wraith
          .GameSpeed = 20
          .GhostIntel(1) = 2
          .GhostIntel(2) = 7
          .GhostIntel(3) = 13
          .GhostIntel(4) = 4
          .GhostIntel(5) = 9
        Case 15, 16
          .GhostBaseSpeed(1) = 4000 'dopy
          .GhostBaseSpeed(2) = 2000 'poky
          .GhostBaseSpeed(3) = 4000 'brain
          .GhostBaseSpeed(4) = 4000 'speedy
          .GhostBaseSpeed(5) = 2800 'wraith
          .GameSpeed = 19
          .GhostIntel(1) = 2
          .GhostIntel(2) = 8
          .GhostIntel(3) = 14
          .GhostIntel(4) = 5
          .GhostIntel(5) = 9
        Case 17, 18
          .GhostBaseSpeed(1) = 4000 'dopy
          .GhostBaseSpeed(2) = 2000 'poky
          .GhostBaseSpeed(3) = 2800 'brain
          .GhostBaseSpeed(4) = 4000 'speedy
          .GhostBaseSpeed(5) = 2800 'wraith
          .GameSpeed = 18
          .GhostIntel(1) = 3
          .GhostIntel(2) = 8
          .GhostIntel(3) = 14
          .GhostIntel(4) = 5
          .GhostIntel(5) = 10
        Case Else
          .GhostBaseSpeed(1) = 4000 'dopy
          .GhostBaseSpeed(2) = 2800 'poky
          .GhostBaseSpeed(3) = 4000 'brain
          .GhostBaseSpeed(4) = 4000 'speedy
          .GhostBaseSpeed(5) = 4000 'wraith
          .GameSpeed = 17
          .GhostIntel(1) = 3
          .GhostIntel(2) = 8
          .GhostIntel(3) = 14
          .GhostIntel(4) = 5
          .GhostIntel(5) = 10
      End Select
      
      Select Case Difficulty 'difficulty limiters
        Case 0
          If .GhostBaseSpeed(1) > 2800 Then .GhostBaseSpeed(1) = 2800
          If .GhostBaseSpeed(3) > 2800 Then .GhostBaseSpeed(3) = 2800
          If .GhostBaseSpeed(5) > 2800 Then .GhostBaseSpeed(5) = 2800
          If .GhostIntel(2) > 6 Then .GhostIntel(2) = 6
          If .GhostIntel(3) > 9 Then .GhostIntel(3) = 9
          If .GhostIntel(5) > 6 Then .GhostIntel(5) = 6
        Case 1
          If .GhostBaseSpeed(3) > 2800 Then .GhostBaseSpeed(3) = 2800
          If .GhostIntel(5) > 8 Then .GhostIntel(5) = 8
      End Select
    End With
End Sub

Public Sub PrepPlayer()
  Dim loop1 As Long, AnimObj As AnimationObject, NAnimObj As AnimationObject
  
  AllSoundsOff
  
  'clean up all unwanted aanimation objects
  Set AnimObj = DXext.FirstObject
  
  Do While Not (AnimObj Is Nothing)
    If AnimObj.ObjectID >= 40 Then
      Set NAnimObj = AnimObj.NextObject
      
      DXext.AnimationListObjectRemove AnimObj.ObjectID
      
      Set AnimObj = NAnimObj
    Else
      AnimObj.Visible = False
      
      Set AnimObj = AnimObj.NextObject
    End If
  Loop
  
  'initialize player to start/continue on this map level
  With m_SnackMan
    .ActionSequence = Array(14, 15, 16, 17, 18, 17, 16, 15)
    .ActionFrame = .ActionSequenceStart
    
    .UserLong1 = 5
    .UserLong2 = 5
    
    .PosX_1000ths = PlayerData(PlayerActive).SnackManStart(1) * 56000 + 4000
    .PosY_1000ths = PlayerData(PlayerActive).SnackManStart(2) * 56000 + 4000
    
    .Visible = True
  End With
  
  For loop1 = 1 To 4 'ghosts
    With m_Ghost(loop1)
      .ActionSequence = Array(loop1 + 4, loop1 + 18, loop1 + 32, loop1 + 46)
      .ActionFrame = .ActionSequenceStart + loop1 - 1
      
      .UserLong1 = 0
      .UserLong2 = 0
      
      .PosX_1000ths = PlayerData(PlayerActive).GhostHomes(loop1, 1) * 56000 + 4000
      .PosY_1000ths = PlayerData(PlayerActive).GhostHomes(loop1, 2) * 56000 + 4000
      
      .UserLong4 = .PosX_1000ths
      .UserLong5 = .PosY_1000ths
      
      GhostFastMode(loop1) = False
      
      .Visible = True
    End With
    
    'super power-ups
    With m_PowerUp(loop1)
      If PlayerData(PlayerActive).SuperDot(loop1, 1) <> -1 Then
        .PosX_1000ths = PlayerData(PlayerActive).SuperDot(loop1, 1) * 56000 + 4000
        .PosY_1000ths = PlayerData(PlayerActive).SuperDot(loop1, 2) * 56000 + 4000
        
        .Visible = True
      End If
    End With
  Next loop1
  
  BonusMode = 1
  PowerUpMode = False
  SpeedMode = False
  SnackmanFastMode = False
  WraithMode = False
  LastDirection = 0
  
  DXext.InitializeBGMap PlayerData(PlayerActive).Map
  DXext.MapMode = SolidBGMap
  
  With PlayerData(PlayerActive)
    frmMain.Display_Score .Lives, .Score, .MapLevel
    TreatMode = .TreatTimeOut
    GameSpeed = .GameSpeed
    m_Ghost(5).UserLong3 = .WraithTimeOut
  End With
End Sub

'get joystick1 and/or keyboard controls1
Public Function Get_Direction()
  If dx_KeyboardState.Key(DIK_UP) Then
    Get_Direction = 1
  ElseIf dx_KeyboardState.Key(DIK_RIGHT) Then
    Get_Direction = 2
  ElseIf dx_KeyboardState.Key(DIK_DOWN) Then
    Get_Direction = 3
  ElseIf dx_KeyboardState.Key(DIK_LEFT) Then
    Get_Direction = 4
  ElseIf activeController > 0 Then
    If dx_JoystickState.Y < -3000 Then
      Get_Direction = 1
    ElseIf dx_JoystickState.X > 3000 Then
      Get_Direction = 2
    ElseIf dx_JoystickState.Y > 3000 Then
      Get_Direction = 3
    ElseIf dx_JoystickState.X < -3000 Then
      Get_Direction = 4
    Else
      Get_Direction = 0
    End If
  End If
End Function

'get keyboard controls2 (for simultaneous 2 player action)
Public Function Get_Direction2()
  If dx_KeyboardState.Key(DIK_NUMPAD8) Then
    Get_Direction = 1
  ElseIf dx_KeyboardState.Key(DIK_NUMPAD6) Then
    Get_Direction = 2
  ElseIf dx_KeyboardState.Key(DIK_NUMPAD2) Then
    Get_Direction = 3
  ElseIf dx_KeyboardState.Key(DIK_NUMPAD4) Then
    Get_Direction = 4
  Else
    Get_Direction = 0
  End If
End Function

'check snackman's direction for validity and possible dot eat'n
Public Function Test_Direction1(ByVal newDirection As Long, Optional TakeAction As Boolean = False) As Boolean
  Dim baseX As Long, baseY As Long, ElementValue As Long
  
  Test_Direction1 = False
  
  With m_SnackMan
    baseX = .PosX_1000ths - 4000
    baseY = .PosY_1000ths - 4000
    
    Select Case newDirection
      Case 1
        If baseX Mod 56000 = 0 Then
          If baseY Mod 56000 = 0 Then
            ElementValue = DXext.BGMapElement(baseY \ 56000, baseX \ 56000)
            
            If ElementValue < 96 Then ElementValue = ElementValue Mod 12
            
            Select Case ElementValue
              Case 1, 3, 4, 6, 7, 10, 11
                Test_Direction1 = True
            End Select
          Else
            Test_Direction1 = True
            
            If TakeAction Then
              If baseY Mod 56000 <= 20000 Then
                ElementValue = DXext.BGMapElement(baseY \ 56000, baseX \ 56000)
                
                If ElementValue Mod 24 >= 12 And ElementValue < 96 Then EatDot baseY, baseX, ElementValue
              End If
            End If
          End If
        End If
      Case 2
        If baseY Mod 56000 = 0 Then
          If baseX < 0 Or baseX >= 672000 Then
            Test_Direction1 = True
          ElseIf baseX Mod 56000 = 0 Then
            ElementValue = DXext.BGMapElement(baseY \ 56000, baseX \ 56000)
            
            If ElementValue < 96 Then ElementValue = ElementValue Mod 12
            
            Select Case ElementValue
              Case 2, 3, 4, 5, 7, 8, 11, (96 Or MIUseBaseOffset), (100 Or MIUseBaseOffset)
                Test_Direction1 = True
            End Select
          Else
            Test_Direction1 = True
            
            If TakeAction Then
              If baseX Mod 56000 >= 36000 And baseX < 616000 Then
                ElementValue = DXext.BGMapElement(baseY \ 56000, (baseX + 20000) \ 56000)
                
                If ElementValue Mod 24 >= 12 And ElementValue < 96 Then EatDot baseY, baseX + 20000, ElementValue
              End If
            End If
          End If
        End If
      Case 3
        If baseX Mod 56000 = 0 Then
          If baseY Mod 56000 = 0 Then
            ElementValue = DXext.BGMapElement(baseY \ 56000, baseX \ 56000)
            
            If ElementValue < 96 Then ElementValue = ElementValue Mod 12
            
            Select Case ElementValue
              Case 1, 4, 5, 6, 8, 9, 11
                Test_Direction1 = True
            End Select
          Else
            Test_Direction1 = True
            
            If TakeAction Then
              If baseY Mod 56000 >= 36000 Then
                ElementValue = DXext.BGMapElement((baseY + 20000) \ 56000, baseX \ 56000)
                
                If ElementValue Mod 24 >= 12 And ElementValue < 96 Then EatDot baseY + 20000, baseX, ElementValue
              End If
            End If
          End If
        End If
      Case Else
        If baseY Mod 56000 = 0 Then
          If baseX < 0 Or baseX >= 672000 Then
            Test_Direction1 = True
          ElseIf baseX Mod 56000 = 0 Then
            ElementValue = DXext.BGMapElement(baseY \ 56000, baseX \ 56000)
            
            If ElementValue < 96 Then ElementValue = ElementValue Mod 12
            
            Select Case ElementValue
              Case 2, 3, 5, 6, 9, 10, 11, (96 Or MIUseBaseOffset), (100 Or MIUseBaseOffset)
                Test_Direction1 = True
            End Select
          Else
            Test_Direction1 = True
            
            If TakeAction And baseX > 0 Then
              If baseX Mod 56000 <= 20000 Then
                ElementValue = DXext.BGMapElement(baseY \ 56000, baseX \ 56000)
                
                If ElementValue Mod 24 >= 12 And ElementValue < 96 Then EatDot baseY, baseX, ElementValue
              End If
            End If
          End If
        End If
    End Select
  End With
End Function

Public Sub IntroMusic(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Play_SoundBuffer 1, False
  Else
    Stop_SoundBuffer 1
  End If
  
  m_SoundActive(1) = playMode
End Sub

Public Sub SnackManSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Play_SoundBuffer 2, True
  Else
    Stop_SoundBuffer 2
  End If
  
  m_SoundActive(2) = playMode
End Sub

Public Sub SnackManPan()
  Change_SoundSettings 2, , , m_SnackMan.PosX_1000ths \ 1000 - 336
End Sub

Public Sub EatDotSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 3, , , m_SnackMan.PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 3, False
  Else
    Stop_SoundBuffer 3
  End If
  
  m_SoundActive(3) = playMode
End Sub

Public Sub EatPowerUpSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 4, , , m_SnackMan.PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 4, False
  Else
    Stop_SoundBuffer 4
  End If
  
  m_SoundActive(4) = playMode
End Sub

Public Sub EatGhostSound(ByVal GhostNum As Long, Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 4 + GhostNum, , , m_SnackMan.PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 4 + GhostNum, False
  Else
    Stop_SoundBuffer 4 + GhostNum
  End If
  
  m_SoundActive(4 + GhostNum) = playMode
End Sub

Public Sub WraithDeathSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 9, , , m_Ghost(5).PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 9, False
  Else
    Stop_SoundBuffer 9
  End If
  
  m_SoundActive(9) = playMode
End Sub

Public Sub EatTreatSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 10, , , m_SnackMan.PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 10, False
  Else
    Stop_SoundBuffer 10
  End If
  
  m_SoundActive(10) = playMode
End Sub

Public Sub WraithAlertSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 11, , , m_Ghost(5).PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 11, False
  Else
    Stop_SoundBuffer 11
  End If
  
  m_SoundActive(11) = playMode
End Sub

Public Sub PlayerStartMusic(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Play_SoundBuffer 12, False
  Else
    Stop_SoundBuffer 12
  End If
  
  m_SoundActive(12) = playMode
End Sub

Public Sub PowerUpModeSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    If Not m_SoundActive(13) Then Play_SoundBuffer 13, True
  Else
    Stop_SoundBuffer 13
  End If
  
  m_SoundActive(13) = playMode
End Sub

Public Sub SnackManDeathSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 14, , , m_SnackMan.PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 14, False
  Else
    Stop_SoundBuffer 14
  End If
  
  m_SoundActive(14) = playMode
End Sub

Public Sub ClearedLevelSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Play_SoundBuffer 15, False
  Else
    Stop_SoundBuffer 15
  End If
  
  m_SoundActive(15) = playMode
End Sub

Public Sub TreatBounceSound(ByVal TreatNum As Long, Optional ByVal playMode As Boolean = True)
  If playMode Then
    Change_SoundSettings 16, , , m_Treat(TreatNum).PosX_1000ths \ 1000 - 336
    Play_SoundBuffer 16, False
  Else
    Stop_SoundBuffer 16
  End If
  
  m_SoundActive(16) = playMode
End Sub

Public Sub BonusManSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Play_SoundBuffer 17, False
  Else
    Stop_SoundBuffer 17
  End If
  
  m_SoundActive(17) = playMode
End Sub

Public Sub BonusValueSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Play_SoundBuffer 18, False
  Else
    Stop_SoundBuffer 18
  End If
  
  m_SoundActive(18) = playMode
End Sub

 Public Sub GameOverSound(Optional ByVal playMode As Boolean = True)
  If playMode Then
    Play_SoundBuffer 19, False
  Else
    Stop_SoundBuffer 19
  End If
  
  m_SoundActive(19) = playMode
End Sub

Public Sub AllSoundsOff()
  Dim loop1 As Long
  
  For loop1 = 1 To 20
    Stop_SoundBuffer loop1
    
    m_SoundActive(loop1) = False
  Next loop1
End Sub

Public Sub PauseAllSound()
  Dim loop1 As Long
  
  m_SoundPaused = True
  
  For loop1 = 1 To 20
    Stop_SoundBuffer loop1
  Next loop1
End Sub

Public Sub UnPauseAllSound()
  Dim loop1 As Long
  
  m_SoundPaused = False
  
  For loop1 = 1 To 20
    If m_SoundActive(loop1) Then
      Select Case loop1
        Case 2, 13
          Resume_SoundBuffer loop1, True
        Case Else
          'Resume_SoundBuffer loop1, False
      End Select
    End If
  Next loop1
End Sub

Public Sub EatDot(ByVal baseY As Long, ByVal baseX As Long, ByVal ElementValue As Long)
  DXext.BGMapElement(baseY \ 56000, baseX \ 56000) = ElementValue - 12
  
  With PlayerData(PlayerActive)
    .Score = .Score + .DotValue * BonusMode
                
    .DotCount = .DotCount - 1
  End With
  
  EatDotSound True
End Sub

'check anything's direction for validity
Public Function Test_Direction4(ByVal newDirection As Long, ByVal baseX As Long, ByVal baseY As Long) As Boolean
  Dim ElementValue As Long
  
  Test_Direction4 = False
  
  baseX = (baseX - 4000) \ 56000
  baseY = (baseY - 4000) \ 56000
  
  ElementValue = DXext.BGMapElement(baseY, baseX)
  If ElementValue < 96 Then ElementValue = ElementValue Mod 12
            
  Select Case newDirection
      Case 1
            Select Case ElementValue
              Case 1, 3, 4, 6, 7, 10, 11
                Test_Direction4 = True
            End Select
      Case 2
            Select Case ElementValue
              Case 2, 3, 4, 5, 7, 8, 11, (96 Or MIUseBaseOffset), (100 Or MIUseBaseOffset)
                Test_Direction4 = True
            End Select
      Case 3
            Select Case ElementValue
              Case 1, 4, 5, 6, 8, 9, 11
                Test_Direction4 = True
            End Select
      Case Else
            Select Case ElementValue
              Case 2, 3, 5, 6, 9, 10, 11, (96 Or MIUseBaseOffset), (100 Or MIUseBaseOffset)
                Test_Direction4 = True
            End Select
  End Select
End Function

Public Sub StoreMap()
  Dim loop1 As Long, loop2 As Long, mapLine(0 To 11) As Variant
  
  With PlayerData(PlayerActive)
    For loop1 = 0 To 7
      For loop2 = 0 To 11
        mapLine(loop2) = DXext.BGMapElement(loop1, loop2)
      Next loop2
      
      .Map(loop1) = mapLine()
    Next loop1
  End With
End Sub
