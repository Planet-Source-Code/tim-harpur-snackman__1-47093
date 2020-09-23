VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SnackMan v1.10 developed by Logicon Enterprises        (press Esc to pause and access settings)"
   ClientHeight    =   8325
   ClientLeft      =   150
   ClientTop       =   90
   ClientWidth     =   11895
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Amelia BT"
      Size            =   48
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picWindow 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   8100
      Left            =   90
      ScaleHeight     =   540
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   780
      TabIndex        =   0
      Top             =   120
      Width           =   11700
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Visible         =   0   'False
      Begin VB.Menu mnuResume 
         Caption         =   "Resume"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "&Display"
      Visible         =   0   'False
      Begin VB.Menu mnuScreenmode 
         Caption         =   "Window 800x600"
         Index           =   0
      End
      Begin VB.Menu mnuScreenmode 
         Caption         =   "Full Screen 800x600x16bit (F8)"
         Index           =   1
      End
      Begin VB.Menu mnuScreenmode 
         Caption         =   "Full Screen 800x600x32bit (F9)"
         Index           =   2
      End
      Begin VB.Menu mnuScreenmode 
         Caption         =   "Full Screen 800x600x8bit (F4)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuControllerTitle 
      Caption         =   "&Controller"
      Visible         =   0   'False
      Begin VB.Menu mnuController 
         Caption         =   "None"
         Index           =   0
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuController 
         Caption         =   ""
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
  Shift = 0
End Sub

Private Sub Form_Load()
  GameMode = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not VerifyExit Then Exit Sub
  VerifyExit = False
  
  PauseAllSound
  PauseMode = True
  
  DXext.ReDrawAnimationWindow
  DXext.FadeDisplay , FMBlack
  
  DXext.DisplayStaticImageTransparent 1, 0, 192, 576, 96, 88, 222, True
  
  DXext.RefreshDisplay
  DXext.FlipBuffers
  
  Reaquire_DXInput
  
  Do While True
    TimeLoop = DelayTillTime(TimeLoop, 25) + 25
    
    PollKeyboard
    Read_Keyboard
    
    If dx_KeyboardState.Key(DIK_Y) <> 0 Then
      GameMode = -1
      
      Exit Do
    ElseIf dx_KeyboardState.Key(DIK_N) <> 0 Then
      Cancel = 1
      
      Reaquire_DXInput
      
      UnPauseAllSound
      PauseMode = False
      
      TimeLoop = DelayTillTime(0, , True)
      
      Exit Do
    End If
    
    DXext.RefreshDisplay
  Loop
End Sub

Public Sub ResetGame()
  PauseAllSound
  
  DXext.ReDrawAnimationWindow
  DXext.FadeDisplay , FMBlack
  
  DXext.DisplayStaticImageTransparent 1, 0, 96, 384, 96, 184, 222, True
  
  DXext.RefreshDisplay
  DXext.FlipBuffers
  
  Reaquire_DXInput
  
  Do While True
    TimeLoop = DelayTillTime(TimeLoop, 25) + 25
    
    PollKeyboard
    Read_Keyboard
    
    If dx_KeyboardState.Key(DIK_Y) <> 0 Then
      GameMode = 2
      
      Reaquire_DXInput
      
      Exit Do
    ElseIf dx_KeyboardState.Key(DIK_N) <> 0 Then
      Reaquire_DXInput
      
      UnPauseAllSound
      
      Exit Do
    End If
    
    DXext.RefreshDisplay
  Loop
  
  PauseMode = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  GameMode = -1
  PauseMode = True
  
  CleanUp_DXSound
  CleanUp_DXInput
  
  DXext.Cleanup_AnimationWindow
End Sub

Public Sub GameMode_1()
  Static FailCount As Long
  
  Init_DXInput Me, True, activeController
  Init_DXSound frmMain, 20, 2
  
  Set_JoystickRange 10000, True, True
  Set_JoystickDeadZoneSat 1000, 8000, True, True
  
  DXext.Init_DXDrawScreen Me, 800, 600, 16, 60, 6
  DXext.DelayTillTime 50, 0, True
  
  If DXext.TestDisplayValid() Then
    GameMode = 2
    FailCount = 0
  Else
    FailCount = FailCount + 1
    
    If FailCount >= 3 Then
      CleanUp_DXInput
      CleanUp_DXSound
      
      MsgBox "Failed to initialize a DirectX link. Please check that DirectX7 or better is properly installed on your machine before running this software.", vbOKOnly, "DirectX Error"
      
      End
    End If
  End If
End Sub

Public Sub GameMode_2() 'load and initialize program vars (including those for titles)
  Dim m_object As AnimationObject, loop1 As Long
  
  Randomize
  
  If DXext.StaticSurfaceFileName(1) <> (App.Path & "\Static.bmp") Then DXext.Init_StaticSurfaceFromFile 1, App.Path & "\Static.bmp", 576, 720, 0
  If DXext.StaticSurfaceFileName(2) <> (App.Path & "\Animation.bmp") Then DXext.Init_StaticSurfaceFromFile 2, App.Path & "\Animation.bmp", 672, 432, 0
  If DXext.StaticSurfaceFileName(3) <> (App.Path & "\Map.bmp") Then DXext.Init_StaticSurfaceFromFile 3, App.Path & "\Map.bmp", 672, 504, 0
  If DXext.StaticSurfaceFileName(4) <> (App.Path & "\LogiconPlaque.gif") Then DXext.Init_StaticSurfaceFromFile 4, App.Path & "\LogiconPlaque.gif", 600, 198, 0
  If DXext.StaticSurfaceFileName(5) <> (App.Path & "\Presents.gif") Then DXext.Init_StaticSurfaceFromFile 5, App.Path & "\Presents.gif", 156, 31, 0
  If DXext.StaticSurfaceFileName(6) <> (App.Path & "\Snackman.gif") Then DXext.Init_StaticSurfaceFromFile 6, App.Path & "\Snackman.gif", 324, 54, 0
  
  DXext.ClearDisplay True
  DXext.SetAnimationWindow 22, 22, 672, 448
  DXext.RGBBackColour = &H0
  DXext.MapMode = NoBGorFGMap
  
  If DXext.BGMapStaticSurfaceNum <> 3 Then
    DXext.InitializeBGMapImages 3, 56, 56
    DXext.BGMapDisplayWidth = 12
    DXext.BGMapDisplayHeight = 8
  End If
  
  If DXext.AnimationListObject(100) Is Nothing Then Set m_Title(1) = DXext.AnimationListObjectAdd(4, 600, 198, , , , 100, 0)
  If DXext.AnimationListObject(101) Is Nothing Then Set m_Title(2) = DXext.AnimationListObjectAdd(5, 156, 31, , , , 101, 1)
  If DXext.AnimationListObject(102) Is Nothing Then Set m_Title(3) = DXext.AnimationListObjectAdd(6, 324, 54, , , , 102, 0)
  
  If DXext.AnimationListObject(1) Is Nothing Then Set m_SnackMan = DXext.AnimationListObjectAdd(2, 48, 48, , , , 1, 100)
  m_SnackMan.CollisionMaskMe = CMSnackman
  m_SnackMan.CollisionMaskTarget = CMGhost Or CMTreats Or CMWalls
  
  For loop1 = 1 To 5 'ghosts
    If DXext.AnimationListObject(1 + loop1) Is Nothing Then Set m_Ghost(loop1) = DXext.AnimationListObjectAdd(2, 48, 48, , , , loop1 + 1, 89 + loop1)
    m_Ghost(loop1).CollisionMaskMe = CMGhost
    
    If loop1 = 5 Then
      m_Ghost(5).CollisionMaskTarget = CMSnackman
    Else
      m_Ghost(loop1).CollisionMaskTarget = CMSnackman Or CMWalls
    End If
  Next loop1
  
  For loop1 = 1 To 4 'super power-ups
    If DXext.AnimationListObject(6 + loop1) Is Nothing Then
      Set m_PowerUp(loop1) = DXext.AnimationListObjectAdd(2, 48, 48, , , , loop1 + 6)
      
      With m_PowerUp(loop1)
        .ActionSequence = Array(97, 111, 125, 124, 125, 111)
        .ActionFrame = .ActionSequenceStart
        .CollisionMaskMe = CMTreats
        .CollisionMaskTarget = CMSnackman
      End With
    End If
  Next loop1
  
  If DXext.AnimationListObject(60) Is Nothing Then Set m_GhostT(1) = DXext.AnimationListObjectAdd(1, 84, 48, , 384, 672, 60)  'Poky - purple
  If DXext.AnimationListObject(61) Is Nothing Then Set m_GhostT(2) = DXext.AnimationListObjectAdd(1, 84, 48, , 96, 672, 61) 'Dopy - blue
  If DXext.AnimationListObject(62) Is Nothing Then Set m_GhostT(3) = DXext.AnimationListObjectAdd(1, 84, 48, , 180, 672, 62) 'Brain -green
  If DXext.AnimationListObject(63) Is Nothing Then Set m_GhostT(4) = DXext.AnimationListObjectAdd(1, 108, 48, , 268, 672, 63) 'Speedy -red
  If DXext.AnimationListObject(64) Is Nothing Then Set m_GhostT(5) = DXext.AnimationListObjectAdd(1, 114, 48, , 462, 672, 64) 'Wraith
  
  If DXext.AnimationListObject(20) Is Nothing Then Set m_Treat(1) = DXext.AnimationListObjectAdd(2, 48, 48, , 576, 192, 20, 95)  '?
  If DXext.AnimationListObject(40) Is Nothing Then Set m_Title3(1) = DXext.AnimationListObjectAdd(2, 48, 48, , 192, 384, 40) '???
  If DXext.AnimationListObject(21) Is Nothing Then Set m_Treat(2) = DXext.AnimationListObjectAdd(2, 48, 48, , 576, 240, 21, 95) 'chocolate bar
  If DXext.AnimationListObject(41) Is Nothing Then Set m_Title3(2) = DXext.AnimationListObjectAdd(2, 48, 48, , 0, 384, 41) 'speed
  If DXext.AnimationListObject(22) Is Nothing Then Set m_Treat(3) = DXext.AnimationListObjectAdd(2, 48, 48, , 480, 240, 22, 95) 'bon bon
  If DXext.AnimationListObject(42) Is Nothing Then Set m_Title3(3) = DXext.AnimationListObjectAdd(2, 48, 48, , 0, 192, 42) '50
  If DXext.AnimationListObject(23) Is Nothing Then Set m_Treat(4) = DXext.AnimationListObjectAdd(2, 48, 48, , 480, 192, 23, 95)  'donut
  If DXext.AnimationListObject(43) Is Nothing Then Set m_Title3(4) = DXext.AnimationListObjectAdd(2, 48, 48, , 48, 192, 43) '100
  If DXext.AnimationListObject(24) Is Nothing Then Set m_Treat(5) = DXext.AnimationListObjectAdd(2, 48, 48, , 528, 240, 24, 95) 'popsicle
  If DXext.AnimationListObject(44) Is Nothing Then Set m_Title3(5) = DXext.AnimationListObjectAdd(2, 48, 48, , 96, 192, 44) '200
  If DXext.AnimationListObject(25) Is Nothing Then Set m_Treat(6) = DXext.AnimationListObjectAdd(2, 48, 48, , 528, 192, 25, 95) 'hot dog
  If DXext.AnimationListObject(45) Is Nothing Then Set m_Title3(6) = DXext.AnimationListObjectAdd(2, 48, 48, , 144, 192, 45) '300
  If DXext.AnimationListObject(26) Is Nothing Then Set m_Treat(7) = DXext.AnimationListObjectAdd(2, 48, 48, , 576, 288, 26, 95) 'french fries
  If DXext.AnimationListObject(46) Is Nothing Then Set m_Title3(7) = DXext.AnimationListObjectAdd(2, 48, 48, , 192, 192, 46)  '400
  If DXext.AnimationListObject(27) Is Nothing Then Set m_Treat(8) = DXext.AnimationListObjectAdd(2, 48, 48, , 576, 336, 27, 95) 'hamburger
  If DXext.AnimationListObject(47) Is Nothing Then Set m_Title3(8) = DXext.AnimationListObjectAdd(2, 48, 48, , 240, 192, 47)  '500
  If DXext.AnimationListObject(28) Is Nothing Then Set m_Treat(9) = DXext.AnimationListObjectAdd(2, 48, 48, , 480, 336, 28, 95) 'pretzel
  If DXext.AnimationListObject(48) Is Nothing Then Set m_Title3(9) = DXext.AnimationListObjectAdd(2, 48, 48, , 288, 192, 48)  '1000
  If DXext.AnimationListObject(29) Is Nothing Then Set m_Treat(10) = DXext.AnimationListObjectAdd(2, 48, 48, , 528, 288, 29, 95)  'all day sucker
  If DXext.AnimationListObject(49) Is Nothing Then Set m_Title3(10) = DXext.AnimationListObjectAdd(2, 48, 48, , 336, 192, 49)  '2000
  If DXext.AnimationListObject(30) Is Nothing Then Set m_Treat(11) = DXext.AnimationListObjectAdd(2, 48, 48, , 528, 336, 30, 95) 'super burger
  If DXext.AnimationListObject(50) Is Nothing Then Set m_Title3(11) = DXext.AnimationListObjectAdd(2, 48, 48, , 384, 192, 50) '4000
  If DXext.AnimationListObject(31) Is Nothing Then Set m_Treat(12) = DXext.AnimationListObjectAdd(2, 48, 48, , 480, 288, 31, 95) 'cherries
  If DXext.AnimationListObject(51) Is Nothing Then Set m_Title3(12) = DXext.AnimationListObjectAdd(2, 48, 48, , 432, 192, 51) '8000
  
  For loop1 = 1 To 12
    m_Treat(loop1).CollisionMaskMe = CMTreats
    m_Treat(loop1).CollisionMaskTarget = CMSnackman Or CMWalls
  Next loop1
  
  For loop1 = 1 To 5
    If DXext.AnimationListObject(34 + loop1) Is Nothing Then Set m_Scores(loop1) = DXext.AnimationListObjectAdd(2, 48, 48, , , , 34 + loop1, 101)  'score
  Next loop1
  
  If DXext.AnimationListObject(200) Is Nothing Then Set m_Title2(1) = DXext.AnimationListObjectAdd(1, 480, 96, , 0, 576, 200)
  If DXext.AnimationListObject(201) Is Nothing Then Set m_Title2(2) = DXext.AnimationListObjectAdd(1, 384, 48, , 0, 288, 201)
  If DXext.AnimationListObject(202) Is Nothing Then Set m_Title2(3) = DXext.AnimationListObjectAdd(1, 384, 48, , 0, 384, 202)
  If DXext.AnimationListObject(203) Is Nothing Then Set m_Title2(4) = DXext.AnimationListObjectAdd(1, 192, 48, , 384, 384, 203)
  If DXext.AnimationListObject(204) Is Nothing Then Set m_Title2(5) = DXext.AnimationListObjectAdd(1, 240, 48, , 0, 478, 204)
  If DXext.AnimationListObject(205) Is Nothing Then Set m_Title2(6) = DXext.AnimationListObjectAdd(1, 192, 48, , 384, 288, 205)
  If DXext.AnimationListObject(206) Is Nothing Then Set m_Title2(7) = DXext.AnimationListObjectAdd(1, 180, 48, , 396, 0, 206)
  If DXext.AnimationListObject(207) Is Nothing Then Set m_Title2(8) = DXext.AnimationListObjectAdd(1, 150, 48, , 240, 0, 207)
  If DXext.AnimationListObject(208) Is Nothing Then Set m_Title2(9) = DXext.AnimationListObjectAdd(1, 48, 48, , 480, 576, 208)
  
  Display_Border
  
  'load game sounds
  If Get_SoundBufferFileName(1) <> (App.Path & "\GameStartIntro.wav") Then CreateSoundBuffer 1, App.Path & "\GameStartIntro.wav"
  If Get_SoundBufferFileName(2) <> (App.Path & "\SnackMan.wav") Then CreateSoundBuffer 2, App.Path & "\SnackMan.wav", , -500
  If Get_SoundBufferFileName(3) <> (App.Path & "\EatDot.wav") Then CreateSoundBuffer 3, App.Path & "\EatDot.wav"
  If Get_SoundBufferFileName(4) <> (App.Path & "\EatPowerUp.wav") Then CreateSoundBuffer 4, App.Path & "\EatPowerUp.wav"
  If Get_SoundBufferFileName(5) <> (App.Path & "\EatGhost.wav") Then CreateSoundBuffer 5, App.Path & "\EatGhost.wav"
  
  If Get_SoundBufferFileName(6) <> (App.Path & "\EatGhost.wav") Then DuplicateSoundBuffer 6, 5
  If Get_SoundBufferFileName(7) <> (App.Path & "\EatGhost.wav") Then DuplicateSoundBuffer 7, 5
  If Get_SoundBufferFileName(8) <> (App.Path & "\EatGhost.wav") Then DuplicateSoundBuffer 8, 5
  
  If Get_SoundBufferFileName(9) <> (App.Path & "\WraithDeath.wav") Then CreateSoundBuffer 9, App.Path & "\WraithDeath.wav"
  If Get_SoundBufferFileName(10) <> (App.Path & "\EatTreat.wav") Then CreateSoundBuffer 10, App.Path & "\EatTreat.wav"
  If Get_SoundBufferFileName(11) <> (App.Path & "\Wraith.wav") Then CreateSoundBuffer 11, App.Path & "\Wraith.wav"
  If Get_SoundBufferFileName(12) <> (App.Path & "\PlayerStartIntro.wav") Then CreateSoundBuffer 12, App.Path & "\PlayerStartIntro.wav"
  If Get_SoundBufferFileName(13) <> (App.Path & "\SnackManPowerUp.wav") Then CreateSoundBuffer 13, App.Path & "\SnackManPowerUp.wav", -500
  If Get_SoundBufferFileName(14) <> (App.Path & "\SnackManDie.wav") Then CreateSoundBuffer 14, App.Path & "\SnackManDie.wav"
  If Get_SoundBufferFileName(15) <> (App.Path & "\LevelUp.wav") Then CreateSoundBuffer 15, App.Path & "\LevelUp.wav"
  If Get_SoundBufferFileName(16) <> (App.Path & "\TreatBounce.wav") Then CreateSoundBuffer 16, App.Path & "\TreatBounce.wav"
  If Get_SoundBufferFileName(17) <> (App.Path & "\BonusMan.wav") Then CreateSoundBuffer 17, App.Path & "\BonusMan.wav"
  If Get_SoundBufferFileName(18) <> (App.Path & "\BonusValue.wav") Then CreateSoundBuffer 18, App.Path & "\BonusValue.wav"
  If Get_SoundBufferFileName(19) <> (App.Path & "\GameOver.wav") Then CreateSoundBuffer 19, App.Path & "\GameOver.wav"
  
  DXext.AnimationListObjectDeactivateAll
  
  Display_Score 0, 0, 0, True
  Display_Specials 0, False, False, 1, False, True
  
  TimeLoop = DelayTillTime(TimeLoop, False)
  
  GameMode = 3
  GameSubMode = -1
End Sub

Public Sub GameMode_3() 'do title sequence
  Dim m_object As AnimationObject, loop1 As Long
  Static val1 As Long, val2 As Long, counter1 As Long
  
  If dx_KeyboardState.Key(DIK_SPACE) <> 0 Then
    If GameSubMode <> 7 And GameSubMode <> 8 Then GameSubMode = 7
    GameCounter = 100
  ElseIf dx_KeyboardState.Key(DIK_F1) <> 0 Then
    'start 1 player game
    
    GameMode_4 1
    
    Exit Sub
  ElseIf dx_KeyboardState.Key(DIK_F2) <> 0 Then
    'start 2 player game
    
    GameMode_4 2
    
    Exit Sub
  ElseIf dx_JoystickState.buttons(1) <> 0 Then
    GameMode_4 1
    
    Exit Sub
  ElseIf dx_JoystickState.buttons(2) <> 0 Then
    GameMode_4 2
    
    Exit Sub
  ElseIf dx_KeyboardState.Key(DIK_F5) <> 0 Then
    'set difficulty easy
    If GameSubMode <> 7 And GameSubMode <> 8 Then GameSubMode = 7
    GameCounter = 100
    
    TreatBounceSound 1, True
    
    Difficulty = 0
  ElseIf dx_KeyboardState.Key(DIK_F6) <> 0 Then
    'set difficulty medium
    If GameSubMode <> 7 And GameSubMode <> 8 Then GameSubMode = 7
    GameCounter = 100
    
    TreatBounceSound 1, True
    
    Difficulty = 1
  ElseIf dx_KeyboardState.Key(DIK_F7) <> 0 Then
    'set difficulty hard
    If GameSubMode <> 7 And GameSubMode <> 8 Then GameSubMode = 7
    GameCounter = 100
    
    TreatBounceSound 1, True
    
    Difficulty = 2
  End If
  
  Select Case GameSubMode
    Case -1
      AllSoundsOff
      
      With m_Title(1)
        .ActionSequence = Array(0)
        .ActionFrame = .ActionSequenceStart
    
        .PosX_1000ths = 56000
        .PosY_1000ths = -220000
    
        .Visible = True
      End With
      
      With m_Title(2)
        .ActionSequence = Array(0)
        .ActionFrame = .ActionSequenceStart
    
        .PosX_1000ths = 680000
        .PosY_1000ths = 230000
    
        .Visible = False
      End With
      
      With m_Title(3)
        .ActionSequence = Array(0)
        .ActionFrame = .ActionSequenceStart
    
        .PosX_1000ths = -380000
        .PosY_1000ths = 320000
    
        .SpecialFX = BFXMirrorLeftRight
    
        .Visible = False
      End With
      
      With m_SnackMan
        .ActionSequence = Array(0, 1, 2, 3, 4, 3, 2, 1)
        .ActionFrame = .ActionSequenceStart
    
        .PosX_1000ths = -50000
        .PosY_1000ths = 320000
    
        .UserLong1 = 0
    
        .Visible = False
      End With
      
      For loop1 = 1 To 9
        m_Title2(loop1).Visible = False
      Next loop1
      
      For loop1 = 1 To 5
        m_Ghost(loop1).Visible = False
        m_GhostT(loop1).Visible = False
      Next loop1
        
      For loop1 = 1 To 12
        m_Treat(loop1).Visible = False
        m_Title3(loop1).Visible = False
      Next loop1
      
      IntroMusic True
      
      GameSubMode = 0
    Case 0
      If m_Title(1).PosY_1000ths < 10000 Then
        m_Title(1).PosY_1000ths = m_Title(1).PosY_1000ths + 3000
      Else
        GameSubMode = 1
      End If
    Case 1
      m_Title(2).Visible = True
      
      If m_Title(2).PosX_1000ths > 254000 Then
        m_Title(2).PosX_1000ths = m_Title(2).PosX_1000ths - 5000
      Else
        GameSubMode = 2
      End If
    Case 2
      m_Title(3).Visible = True
      m_SnackMan.Visible = True
      
      SnackManPan
      SnackManSound True

      GameSubMode = 3
    Case 3
      With m_Title(3)
        .PosX_1000ths = .PosX_1000ths + 6000
        
        If .PosX_1000ths >= 880000 Then GameSubMode = 4
      End With
      
      With m_SnackMan
        .PosX_1000ths = .PosX_1000ths + 6000
        SnackManPan
        
        If .UserLong1 = 0 Then
          .ActionFrame = .ActionFrame + 1
          
          If .ActionFrame = .ActionSequenceStop Then .UserLong1 = 1
        Else
          .ActionFrame = .ActionFrame - 1
          
          If .ActionFrame = .ActionSequenceStart Then .UserLong1 = 0
        End If
      End With
    Case 4
      With m_Title(3)
        .PosX_1000ths = 786000
        
        .SpecialFX = BFXNoEffects
      End With
      
      With m_SnackMan
        .PosX_1000ths = 736000
        .UserLong1 = 0
        
        .ActionSequence = Array(14, 15, 16, 17, 18, 17, 16, 15)
        .ActionFrame = .ActionSequenceStart
      End With
      
      GameSubMode = 5
    Case 5
      With m_Title(3)
        .PosX_1000ths = .PosX_1000ths - 6000
        
        If .PosX_1000ths <= 174000 Then GameSubMode = 6
      End With
      
      With m_SnackMan
        .PosX_1000ths = .PosX_1000ths - 6000
        SnackManPan
        
        If .UserLong1 = 0 Then
          .ActionFrame = .ActionFrame + 1
          
          If .ActionFrame = .ActionSequenceStop Then .UserLong1 = 1
        Else
          .ActionFrame = .ActionFrame - 1
          
          If .ActionFrame = .ActionSequenceStart Then .UserLong1 = 0
        End If
      End With
    Case 6
      With m_SnackMan
        .PosX_1000ths = .PosX_1000ths - 6000
        SnackManPan
        
        If .PosX_1000ths < -50000 Then
          GameSubMode = 7
          GameCounter = 0
          
          SnackManSound False
        End If
        
        If .UserLong1 = 0 Then
          .ActionFrame = .ActionFrame + 1
          
          If .ActionFrame = .ActionSequenceStop Then .UserLong1 = 1
        Else
          .ActionFrame = .ActionFrame - 1
          
          If .ActionFrame = .ActionSequenceStart Then .UserLong1 = 0
        End If
      End With
    Case 7
      GameCounter = GameCounter + 1
      
      If GameCounter >= 100 Then
        GameSubMode = 8
        GameCounter = 0
        DXext.AnimationListObjectDeactivateAll
        AllSoundsOff
        
        For loop1 = 1 To 12
          m_Treat(loop1).Visible = False
          m_Title3(loop1).Visible = False
        Next loop1
        
        For loop1 = 1 To 3
          m_Title(loop1).Visible = False
        Next loop1
          
        m_SnackMan.Visible = False
        
        For loop1 = 1 To 9
          m_Title2(loop1).Visible = True
        Next loop1
          
        m_Title2(1).PosX_1000ths = 96000
        m_Title2(1).PosY_1000ths = 0
        m_Title2(2).PosX_1000ths = 144000
        m_Title2(2).PosY_1000ths = 90000
        m_Title2(3).PosX_1000ths = 144000
        m_Title2(3).PosY_1000ths = 138000
        m_Title2(4).PosX_1000ths = 260000
        m_Title2(4).PosY_1000ths = 199000
        m_Title2(5).PosX_1000ths = 216000
        m_Title2(5).PosY_1000ths = 247000
        m_Title2(6).PosX_1000ths = 260000
        m_Title2(6).PosY_1000ths = 295000
        m_Title2(7).PosX_1000ths = 246000
        m_Title2(7).PosY_1000ths = 356000
        m_Title2(8).PosX_1000ths = 261000
        m_Title2(8).PosY_1000ths = 404000
        m_Title2(9).PosX_1000ths = 148000
        
        Display_Screen2
        
        val1 = 0
        val2 = 4
        counter1 = 0
        
        ClearedLevelSound True
      End If
    Case 8
      GameCounter = GameCounter + 1
      
      If GameCounter > 560 Then
        GameCounter = 0
        GameSubMode = 9
      Else
        counter1 = counter1 + 1
        
        If counter1 >= 5 Then
          counter1 = 0
          
          val1 = val1 + 1
          If val1 > 7 Then val1 = 1
          
          val2 = val2 + 1
          If val2 > 7 Then val2 = 1
          
          Display_Screen2 val1, val2
          
          Select Case val1
            Case 1
              Display_Specials 1, False, False, 1, False
            Case 2
              Display_Specials 2, False, False, 1, False
            Case 3
              Display_Specials 0, True, False, 1, False
            Case 4
              Display_Specials 0, False, True, 1, False
            Case 5
              Display_Specials 0, False, False, 2, False
            Case 6
              Display_Specials 0, False, False, 1, True
            Case 7
              Display_Specials 0, False, False, 1, False
          End Select
        End If
      End If
    Case 9
      GameSubMode = 10
      
      Display_Specials 0, False, False, 1, False
      
      For loop1 = 2 To 9
        m_Title2(loop1).Visible = False
      Next loop1
      
      For loop1 = 1 To 5
        With m_Ghost(loop1)
          .Visible = True
          
          .PosY_1000ths = 90000
          .PosX_1000ths = 40000 - loop1 * 94000
          
          SetNormalGhost loop1
        End With
        
        With m_GhostT(loop1)
          .PosY_1000ths = 138000
          .PosX_1000ths = loop1 * 112000 - (.ImageWidth \ 2) * 1000
        End With
      Next loop1
    Case 10
      For loop1 = 1 To 5
        With m_Ghost(loop1)
          If .PosX_1000ths < 112000 * loop1 - 24000 Then
            .PosX_1000ths = .PosX_1000ths + 4000
          Else
            m_GhostT(loop1).Visible = True
          End If
          
          .UserLong1 = .UserLong1 + 1
          
          If .UserLong1 > 5 Then
            .UserLong1 = 0
            
            .ActionFrame = .ActionFrame + 1
            
            If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
          End If
        End With
      Next loop1
      
      If m_Ghost(5).PosX_1000ths >= 112000 * 5 - 24000 Then
        GameSubMode = 11
        
        m_GhostT(5).Visible = True
      End If
    Case 11
      GameSubMode = 12
      GameCounter = 0
      
      For loop1 = 1 To 12
        With m_Treat(loop1)
          .Visible = True
          .PosX_1000ths = ((loop1 - 1) Mod 4) * 133000 + 113000
          .PosY_1000ths = -48000 - loop1 * 24000
          
          .UserLong1 = ((loop1 - 1) \ 4) * 80000 + 190000
          
          m_Title3(loop1).PosY_1000ths = .UserLong1 + 38000
          m_Title3(loop1).PosX_1000ths = .PosX_1000ths
        End With
      Next loop1
    Case 12
      For loop1 = 1 To 5
        With m_Ghost(loop1)
          .UserLong1 = .UserLong1 + 1
          
          If .UserLong1 > 5 Then
            .UserLong1 = 0
            
            .ActionFrame = .ActionFrame + 1
            
            If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
          End If
        End With
      Next loop1
      
      For loop1 = 1 To 12
        With m_Treat(loop1)
          If .PosY_1000ths >= .UserLong1 Then
            If Not m_Title3(loop1).Visible Then
              TreatBounceSound loop1, True
              m_Title3(loop1).Visible = True
              
              If loop1 = 12 Then GameSubMode = 13
            End If
          Else
            .PosY_1000ths = .PosY_1000ths + 4000
          End If
        End With
      Next loop1
    Case 13
      GameCounter = GameCounter + 1
      
      If GameCounter > 260 Then
        GameCounter = 0
        GameSubMode = 14
        
        SnackManPan
        SnackManSound True
        EatPowerUpSound True
        PowerUpModeSound True
        
        m_SnackMan.Visible = True
        m_SnackMan.PosX_1000ths = -48000
        m_SnackMan.PosY_1000ths = 90000
        m_SnackMan.ActionSequence = Array(0, 1, 2, 3, 4, 3, 2, 1)
        
        For loop1 = 1 To 4
          With m_Ghost(loop1)
            .ActionSequence = Array(10, 24, 38, 52)
            .UserLong2 = 0
          End With
        Next loop1
      Else
        For loop1 = 1 To 5
          With m_Ghost(loop1)
            .UserLong1 = .UserLong1 + 1
            
            If .UserLong1 > 5 Then
              .UserLong1 = 0
              
              .ActionFrame = .ActionFrame + 1
              
              If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
            End If
          End With
        Next loop1
      End If
    Case 14
      m_SnackMan.PosX_1000ths = m_SnackMan.PosX_1000ths + 4000
      SnackManPan
      m_SnackMan.ActionFrame = m_SnackMan.ActionFrame + 1
      If m_SnackMan.ActionFrame > m_SnackMan.ActionSequenceStop Then m_SnackMan.ActionFrame = m_SnackMan.ActionSequenceStart
      
      For loop1 = 1 To 5
        With m_Ghost(loop1)
          .UserLong1 = .UserLong1 + 1
          
          If .UserLong1 > 5 Then
            .UserLong1 = 0
            
            .ActionFrame = .ActionFrame + 1
            
            If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
          End If
          
          If .UserLong2 = 0 Then
            If m_SnackMan.PosX_1000ths >= .PosX_1000ths - 44000 Then
              .ActionSequence = Array(11, 25, 39, 53)
              .UserLong2 = 1
              
              m_GhostT(loop1).Visible = False
              
              EatGhostSound loop1, True
              
              If loop1 = 4 Then
                GameSubMode = 15
                m_GhostT(5).Visible = False
                WraithAlertSound True
                
                PowerUpModeSound False
                m_SnackMan.ActionSequence = Array(14, 15, 16, 17, 18, 17, 16, 15)
              End If
            End If
          Else
            .PosX_1000ths = .PosX_1000ths - 10000
          End If
        End With
      Next loop1
    Case 15
      m_SnackMan.PosX_1000ths = m_SnackMan.PosX_1000ths - 4000
      SnackManPan
      m_SnackMan.ActionFrame = m_SnackMan.ActionFrame + 1
      If m_SnackMan.ActionFrame > m_SnackMan.ActionSequenceStop Then m_SnackMan.ActionFrame = m_SnackMan.ActionSequenceStart
      
      For loop1 = 1 To 5
        With m_Ghost(loop1)
          .UserLong1 = .UserLong1 + 1
          
          If .UserLong1 > 5 Then
            .UserLong1 = 0
            
            .ActionFrame = .ActionFrame + 1
            
            If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
          End If
          
          If loop1 = 5 Then
            .PosX_1000ths = .PosX_1000ths - 5000
          Else
            .PosX_1000ths = .PosX_1000ths - 10000
          End If
        End With
      Next loop1
      
      If m_Ghost(5).PosX_1000ths < -48000 Then
        GameSubMode = -1
        SnackManSound False
      End If
  End Select
End Sub

Public Sub Display_Score(ByVal Lives As Long, ByVal Score As Long, ByVal Level As Long, Optional ByVal dual As Boolean = True)
  If Lives >= 100 Then Lives = Lives Mod 100
  
  DXext.DisplayStaticImageSolid 1, (Lives \ 10) * 48, 528, 28, 48, 160, 488, True
  DXext.DisplayStaticImageSolid 1, (Lives Mod 10) * 48, 528, 28, 48, 190, 488, True
  
  If Level >= 100 Then Level = Level Mod 100
  
  If Level = -1 Then
    DXext.DisplayStaticImageSolid 1, 504, 504, 28, 48, 346, 488, True
    DXext.DisplayStaticImageSolid 1, 504, 504, 28, 48, 376, 488, True
  Else
    DXext.DisplayStaticImageSolid 1, (Level \ 10) * 48, 528, 28, 48, 346, 488, True
    DXext.DisplayStaticImageSolid 1, (Level Mod 10) * 48, 528, 28, 48, 376, 488, True
  End If
  
  If Score >= 1000000 Then Score = Score Mod 1000000
  
  DXext.DisplayStaticImageSolid 1, (Score \ 100000) * 48, 528, 28, 48, 526, 488, True
  Score = Score Mod 100000
  DXext.DisplayStaticImageSolid 1, (Score \ 10000) * 48, 528, 28, 48, 556, 488, True
  Score = Score Mod 10000
  DXext.DisplayStaticImageSolid 1, (Score \ 1000) * 48, 528, 28, 48, 586, 488, True
  Score = Score Mod 1000
  DXext.DisplayStaticImageSolid 1, (Score \ 100) * 48, 528, 28, 48, 616, 488, True
  Score = Score Mod 100
  DXext.DisplayStaticImageSolid 1, (Score \ 10) * 48, 528, 28, 48, 646, 488, True
  DXext.DisplayStaticImageSolid 1, (Score Mod 10) * 48, 528, 28, 48, 676, 488, True
  
  If dual = True Then DXext.SyncronizeBuffers
End Sub

Public Sub Display_Screen2(Optional ByVal highlight1 As Long = -1, Optional ByVal highlight2 As Long = -1)
  Static arrows As Long
  
  If highlight1 = 1 Or highlight2 = 1 Then
    m_Title2(2).SourceOffsetY = 336
  Else
    m_Title2(2).SourceOffsetY = 288
  End If
  
  If highlight1 = 2 Or highlight2 = 2 Then
    m_Title2(3).SourceOffsetY = 432
  Else
    m_Title2(3).SourceOffsetY = 384
  End If
  
  If highlight1 = 3 Or highlight2 = 3 Then
    m_Title2(4).SourceOffsetY = 432
  Else
    m_Title2(4).SourceOffsetY = 384
  End If
  
  If highlight1 = 4 Or highlight2 = 4 Then
    m_Title2(5).SourceOffsetX = 240
  Else
    m_Title2(5).SourceOffsetX = 0
  End If
  
  If highlight1 = 5 Or highlight2 = 5 Then
    m_Title2(6).SourceOffsetY = 336
  Else
    m_Title2(6).SourceOffsetY = 288
  End If
  
  If highlight1 = 6 Or highlight2 = 6 Then
    m_Title2(7).SourceOffsetY = 48
  Else
    m_Title2(7).SourceOffsetY = 0
  End If
  
  If highlight1 = 7 Or highlight2 = 7 Then
    m_Title2(8).SourceOffsetY = 48
  Else
    m_Title2(8).SourceOffsetY = 0
  End If
  
  arrows = arrows + 1
  If arrows > 4 Then arrows = 1
  
  m_Title2(9).PosY_1000ths = 199000 + Difficulty * 48000
  
  Select Case arrows
    Case 1
      m_Title2(9).SourceOffsetX = 480
      m_Title2(9).SourceOffsetY = 576
    Case 2
      m_Title2(9).SourceOffsetX = 528
      m_Title2(9).SourceOffsetY = 576
    Case 3
      m_Title2(9).SourceOffsetX = 528
      m_Title2(9).SourceOffsetY = 624
    Case 4
      m_Title2(9).SourceOffsetX = 480
      m_Title2(9).SourceOffsetY = 624
  End Select
  
  m_Title2(2).PosX_1000ths = 144000 + Rnd() * 4000
  m_Title2(2).PosY_1000ths = 90000 + Rnd() * 4000
  m_Title2(3).PosX_1000ths = 144000 + Rnd() * 4000
  m_Title2(3).PosY_1000ths = 138000 + Rnd() * 4000
  m_Title2(4).PosX_1000ths = 240000 + Rnd() * 4000
  m_Title2(4).PosY_1000ths = 199000 + Rnd() * 4000
  m_Title2(5).PosX_1000ths = 216000 + Rnd() * 4000
  m_Title2(5).PosY_1000ths = 247000 + Rnd() * 4000
  m_Title2(6).PosX_1000ths = 240000 + Rnd() * 4000
  m_Title2(6).PosY_1000ths = 295000 + Rnd() * 4000
  m_Title2(7).PosX_1000ths = 246000 + Rnd() * 4000
  m_Title2(7).PosY_1000ths = 356000 + Rnd() * 4000
  m_Title2(8).PosX_1000ths = 261000 + Rnd() * 4000
  m_Title2(8).PosY_1000ths = 404000 + Rnd() * 4000
End Sub

Public Sub Display_Specials(ByVal Player As Long, ByVal PowerUp As Boolean, ByVal Speed As Boolean, _
      ByVal Bonus As Long, ByVal Wraith As Boolean, Optional ByVal dual As Boolean = True)
  
  If Player = 1 Then
    DXext.DisplayStaticImageSolid 1, 528, 96, 48, 48, 720, 6, True
  Else
    DXext.DisplayStaticImageSolid 1, 480, 96, 48, 48, 720, 6, True
  End If
  
  If Player = 2 Then
    DXext.DisplayStaticImageSolid 1, 528, 144, 48, 48, 720, 54, True
  Else
    DXext.DisplayStaticImageSolid 1, 480, 144, 48, 48, 720, 54, True
  End If
  
  If PowerUp Then
    DXext.DisplayStaticImageSolid 2, 480, 384, 48, 48, 720, 150, True
  Else
    DXext.DisplayStaticImageSolid 2, 576, 96, 48, 48, 720, 150, True
  End If
  
  If Speed Then
    DXext.DisplayStaticImageSolid 2, 432, 384, 48, 48, 720, 198, True
  Else
    DXext.DisplayStaticImageSolid 2, 576, 0, 48, 48, 720, 198, True
  End If
  
  If Bonus > 1 Then
    DXext.DisplayStaticImageSolid 2, 144 + Bonus * 48, 384, 48, 48, 720, 246, True
  Else
    DXext.DisplayStaticImageSolid 2, 576, 48, 48, 48, 720, 246, True
  End If
  
  If Wraith Then
    DXext.DisplayStaticImageSolid 2, 528, 384, 48, 48, 720, 294, True
  Else
    DXext.DisplayStaticImageSolid 2, 576, 144, 48, 48, 720, 294, True
  End If
  
  If dual Then DXext.SyncronizeBuffers
End Sub

Public Sub SetNormalGhost(ByVal GhostNum As Long)
  With m_Ghost(GhostNum)
    Select Case GhostNum
      Case 1
        .ActionSequence = Array(5, 19, 33, 47)
      Case 2
        .ActionSequence = Array(6, 20, 34, 48)
      Case 3
        .ActionSequence = Array(7, 21, 35, 49)
      Case 4
        .ActionSequence = Array(8, 22, 36, 50)
      Case 5
        .ActionSequence = Array(9, 23, 37, 51)
    End Select
    
    .UserLong2 = 0
    
    If GhostNum = 5 Then
      .ActionFrame = .ActionSequenceStart
    Else
      .ActionFrame = .ActionSequenceStart + GhostNum - 1
      GhostFastMode(GhostNum) = False
    End If
  End With
End Sub

'Reset Game
Public Sub GameMode_4(ByVal numPlayers As Long)
  PlayerActive = 1
  
  With PlayerData(1)
    .Lives = 3
    .MapLevel = 1
    .NextMan = 10000
    .Score = 0
  End With
  
  LoadMapLevel
  
  PlayerActive = 2
  
  With PlayerData(2)
    If numPlayers = 1 Then
      .Lives = 0
      .MapLevel = 0
    Else
      .Lives = 3
      .MapLevel = 1
    End If
    
    .NextMan = 10000
    .Score = 0
  End With
  
  LoadMapLevel
  
  PlayerActive = 1
  PrepPlayer
  GameCounter2 = 0
  
  GameMode = 5
  PlayerStartMusic
End Sub

'Wait for player go
Public Sub GameMode_5()
  Static toggle As Boolean
  
  UpdateAnim
  
  GameCounter2 = GameCounter2 + 1
  
  If GameCounter2 >= 120 Then
    Display_Specials PlayerActive, False, False, 1, False, True
    
    GameMode = 6
    
    m_SnackMan.UserLong1 = 0
    SnackManSound True
  Else
    If GameCounter = 0 Then
      If toggle Then
        Display_Specials PlayerActive, False, False, 1, False, True
        
        toggle = False
      Else
        Display_Specials 0, False, False, 1, False, True
        
        toggle = True
      End If
    End If
  End If
End Sub

'Level cleared - do short delay before proceeding
Public Sub GameMode_7()
  Static toggle As Boolean
  
  UpdateAnim
  
  GameCounter2 = GameCounter2 + 1
  
  With PlayerData(PlayerActive)
    If GameCounter2 >= 80 Then
      .MapLevel = .MapLevel + 1
      
      LoadMapLevel
      PrepPlayer
      
      Display_Score .Lives, .Score, .MapLevel, True
      
      GameCounter2 = 0
      GameMode = 5
      PlayerStartMusic
    Else
      If GameCounter = 0 Then
        If toggle Then
          Display_Score .Lives, .Score, .MapLevel, True
          
          toggle = False
        Else
          Display_Score .Lives, .Score, -1, True
          
          toggle = True
        End If
      End If
    End If
  End With
End Sub

'Player died - do short delay before proceeding
Public Sub GameMode_8()
  Dim loop1 As Long
  
  GameCounter = GameCounter + 1
  
  If GameCounter >= 7 Then
    GameCounter = 0
    
    DXext.BGMapBaseImageIndex = DXext.BGMapBaseImageIndex + 1
    If DXext.BGMapBaseImageIndex > 3 Then DXext.BGMapBaseImageIndex = 0
  
    With m_SnackMan
      .ActionFrame = .ActionFrame + 1
      
      If .ActionFrame > .ActionSequenceStop Then m_SnackMan.Visible = False
    End With
  End If
  
  GameCounter2 = GameCounter2 + 1
  
  If GameCounter2 >= 80 Then
    If PlayerActive = 1 Then
      If PlayerData(2).Lives > 0 Then PlayerActive = 2
    ElseIf PlayerData(1).Lives > 0 Then
      PlayerActive = 1
    End If
    
    If PlayerData(PlayerActive).Lives = 0 Then 'end game
      DXext.ClearDisplay
      
      PlayerActive = 2
      
      If DXext.AnimationListObject(200) Is Nothing Then Set m_Title2(1) = DXext.AnimationListObjectAdd(1, 480, 96, , 0, 576, , 200)
      
      DXext.AnimationListObjectDeactivateAll
      
      DXext.MapMode = NoBGorFGMap
      
      With m_Title2(1)
        .Visible = True
        .PosX_1000ths = 96000
        .PosY_1000ths = 176000
      End With
      
      AllSoundsOff
      GameOverSound
      GameMode = 9
    Else
      PrepPlayer
      
      With PlayerData(PlayerActive)
        Display_Score .Lives, .Score, .MapLevel, True
      End With
      
      GameCounter2 = 0
      GameMode = 5
      AllSoundsOff
      PlayerStartMusic
    End If
  End If
End Sub

'Main game routine
Public Sub GameMode_6()
  Dim temp1 As Long, temp2 As Double, loop1 As Long, loop2 As Long
  Dim pos_x As Long, pos_y As Long
  
  With m_SnackMan
    If .PosX_1000ths >= 0 And .PosX_1000ths < 672000 Then
      temp1 = Get_Direction
      
      If temp1 <> 0 Then LastDirection = temp1
      
      If LastDirection <> 0 Then 'change snackman direction
        If Test_Direction1(LastDirection) Then
          .UserLong1 = LastDirection
          
          LastDirection = 0
        End If
      End If
      
      If .UserLong1 <> .UserLong2 Then
        Select Case .UserLong1
          Case 1 'up
            .ActionSequence = Array(28, 29, 30, 31, 32, 31, 30, 29)
          Case 2 'right
            .ActionSequence = Array(0, 1, 2, 3, 4, 3, 2, 1)
          Case 3 'down
            .ActionSequence = Array(42, 43, 44, 45, 46, 45, 44, 43)
          Case Else 'left
            .ActionSequence = Array(14, 15, 16, 17, 18, 17, 16, 15)
        End Select
        
        If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
        
        .UserLong2 = .UserLong1
      End If
    End If
    
    If Test_Direction1(.UserLong1, True) Then
      If SpeedMode = True Then
        If SnackmanFastMode = False Then
          If (.PosX_1000ths - 4000) Mod 56000 = 0 And (.PosY_1000ths - 4000) Mod 56000 = 0 Then SnackmanFastMode = True
        End If
      Else
        If SnackmanFastMode = True Then
          If (.PosX_1000ths - 4000) Mod 56000 = 0 And (.PosY_1000ths - 4000) Mod 56000 = 0 Then SnackmanFastMode = False
        End If
      End If
      
      If SnackmanFastMode Then
        Select Case .UserLong1
          Case 1
            .PosY_1000ths = .PosY_1000ths - 7000
          Case 2
            .PosX_1000ths = .PosX_1000ths + 7000
            
            If .PosX_1000ths >= 672000 Then .PosX_1000ths = .PosX_1000ths - 728000
          Case 3
            .PosY_1000ths = .PosY_1000ths + 7000
          Case 4
            .PosX_1000ths = .PosX_1000ths - 7000
            
            If .PosX_1000ths <= -56000 Then .PosX_1000ths = .PosX_1000ths + 728000
        End Select
      Else
        Select Case .UserLong1
          Case 1
            .PosY_1000ths = .PosY_1000ths - 4000
          Case 2
            .PosX_1000ths = .PosX_1000ths + 4000
            
            If .PosX_1000ths >= 672000 Then .PosX_1000ths = .PosX_1000ths - 728000
          Case 3
            .PosY_1000ths = .PosY_1000ths + 4000
          Case 4
            .PosX_1000ths = .PosX_1000ths - 4000
            
            If .PosX_1000ths <= -56000 Then .PosX_1000ths = .PosX_1000ths + 728000
        End Select
      End If
    End If
    
    pos_x = .PosX_1000ths
    pos_y = .PosY_1000ths
    
    SnackManPan
  End With
  
  If PlayerData(PlayerActive).DotCount = 0 Then
    'all clear - advance to next level
    AllSoundsOff
    ClearedLevelSound
    
    For loop1 = 1 To 5
      m_Scores(loop1).Visible = False
    Next loop1
    
    With PlayerData(PlayerActive)
      .Score = .Score + .CompletionBonus * BonusMode
      
      GameCounter2 = 0
      GameMode = 7
    End With
  Else
    'check for other actions - time expires, etc.
    If GameCounter = 0 Then
      If TreatMode <> -1 Then 'check if no treat - maybe activate one
        TreatMode = TreatMode - 1
        
        If TreatMode <= 0 Then
          TreatMode = -1
          
          With PlayerData(PlayerActive)
            temp1 = (Rnd() * 15)
            
            If temp1 < 5 Then
              temp1 = 2
            Else
              temp1 = temp1 - 2
              If temp1 > .BestTreat + 1 Then temp1 = .BestTreat + 1
            End If
          End With
          
          If PlayerData(PlayerActive).MapLevel > 5 And Rnd() < 0.1 Then
            With m_Treat(1)
              .ObjectTypeID = temp1
              .PosX_1000ths = PlayerData(PlayerActive).SnackManStart(1) * 56000 + 4000
              .PosY_1000ths = PlayerData(PlayerActive).SnackManStart(2) * 56000 + 4000
              
              .UserLong1 = 0
              .UserLong2 = 0
              .UserLong3 = .PosY_1000ths
              .UserLong4 = PlayerData(PlayerActive).TreatTime
              
              .Visible = True
            End With
          Else
            With m_Treat(temp1)
              .ObjectTypeID = temp1
              .PosX_1000ths = PlayerData(PlayerActive).SnackManStart(1) * 56000 + 4000
              .PosY_1000ths = PlayerData(PlayerActive).SnackManStart(2) * 56000 + 4000
              
              .UserLong1 = 0
              .UserLong2 = 0
              .UserLong3 = .PosY_1000ths
              .UserLong4 = PlayerData(PlayerActive).TreatTime
              
              .Visible = True
            End With
          End If
        End If
      End If
      
      With m_Ghost(5) 'check wraith status
        If WraithMode Then
          .UserLong3 = .UserLong3 - 1
          
          If .UserLong3 = 8 Then
            .ActionSequence = Array(9, 25, 37, 53)
            
            If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
            WraithDeathSound True
          ElseIf .UserLong3 <= 0 Then
            .Visible = False
            WraithMode = False
            
            .UserLong3 = PlayerData(PlayerActive).WraithTimeOut
          End If
        Else
          .UserLong3 = .UserLong3 - 1
          
          If .UserLong3 <= 0 Then
            'activate the wraith
            WraithMode = True
            WraithDeathSound True
            
            .Visible = True
            SetNormalGhost 5
            
            If Rnd() < 0.5 Then
              .PosX_1000ths = PlayerData(PlayerActive).GhostHomes(2, 1) * 56000 + 4000
              .PosY_1000ths = PlayerData(PlayerActive).GhostHomes(2, 2) * 56000 + 4000
            Else
              .PosX_1000ths = PlayerData(PlayerActive).GhostHomes(3, 1) * 56000 + 4000
              .PosY_1000ths = PlayerData(PlayerActive).GhostHomes(3, 2) * 56000 + 4000
            End If
            
            .UserLong1 = -40
            .UserLong2 = 0
            .UserLong3 = PlayerData(PlayerActive).WraithTime
            .UserLong4 = .PosX_1000ths
            .UserLong5 = .PosY_1000ths
          End If
        End If
      End With
      
      If SpeedMode Then
        With m_SnackMan
          .UserLong3 = .UserLong3 - 1
          
          If .UserLong3 <= 0 Then SpeedMode = False
        End With
      End If
      
      If PowerUpMode Then
        With m_SnackMan
          .UserLong4 = .UserLong4 - 1
          
          If .UserLong4 <= 0 Then 'power-up mode ends
            PowerUpMode = False
            PowerUpModeSound False
            
            For loop1 = 1 To 4
              'reset blues to normal sequence (= 2 are vapour and need to return home still)
              If m_Ghost(loop1).UserLong2 = 1 Then SetNormalGhost loop1
            Next loop1
          ElseIf .UserLong4 = 16 Then
            For loop1 = 1 To 4
              'set blues to blink normal sequence (= 2 are vapour and need to return home still)
              With m_Ghost(loop1)
                If .UserLong2 = 1 Then
                  Select Case loop1
                    Case 1
                      .ActionSequence = Array(10, 24, 33, 47)
                    Case 2
                      .ActionSequence = Array(6, 24, 38, 48)
                    Case 3
                      .ActionSequence = Array(7, 21, 38, 52)
                    Case 4
                      .ActionSequence = Array(10, 22, 36, 52)
                    End Select
                    
                    If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart + loop1 - 1
                End If
              End With
            Next loop1
          End If
        End With
      End If
    End If
    
    'check for object collisions
    Dim pos_xt As Long, pos_yt As Long, pos_xtT As Long, pos_ytT As Long
    
    For loop1 = 1 To 5 'check ghost (move and collision)
      With m_Ghost(loop1)
        If .Visible Then
          pos_xt = .PosX_1000ths
          pos_yt = .PosY_1000ths
          
          If pos_xt < 4000 Or pos_xt > 620000 Or pos_yt < 4000 Or pos_yt > 396000 Then
            'skip move logic as ghost is out of bounds - ghost continues until it wraps
          ElseIf (pos_xt - 4000) Mod 56000 = 0 And (pos_yt - 4000) Mod 56000 = 0 Then
            Select Case .UserLong2 'select ghost target
              Case 0 'track
                pos_xtT = pos_x
                pos_ytT = pos_y
              Case 1 'run away
                If pos_x < 336000 Then
                  pos_xtT = 620000
                Else
                  pos_xtT = 4000
                End If
                
                If pos_y < 224000 Then
                  pos_ytT = 392000
                Else
                  pos_ytT = 4000
                End If
              Case 2 'go home
                GhostFastMode(loop1) = True
                
                pos_xtT = .UserLong4
                pos_ytT = .UserLong5
                
                If pos_xtT = pos_xt And pos_ytT = pos_yt Then
                  .UserLong1 = 0
                  
                  SetNormalGhost loop1
                End If
            End Select
            
            If .UserLong1 < 0 Then 'wraith is waiting to come out
              .UserLong1 = .UserLong1 + 1
              
              If .UserLong1 = 0 Then
                .UserLong1 = 1
                
                WraithAlertSound True
              End If
            ElseIf .UserLong1 = 0 Then 'ghost is in chamber - can he come out?
              If loop1 = 5 Then
                .UserLong1 = 1
              ElseIf Not PowerUpMode Then
                If Rnd() < 0.02 Then .UserLong1 = 1
              End If
            ElseIf .UserLong2 = 2 Then 'ghost is vapour
              If pos_xt = .UserLong4 And (pos_yt + 56000) = .UserLong5 Then
                .UserLong1 = 3
              Else
                maxLogicDepth = 10
                logicTargetX = (pos_xtT - 4000) \ 56000
                logicTargetY = (pos_ytT - 60000) \ 56000
                
                If LogicSearch(0, 0, (pos_xt - 4000) \ 56000, (pos_yt - 4000) \ 56000) < 9999 Then
                  .UserLong1 = SmartDirection
                Else 'use dumb logic until it gets closer
                  temp1 = .UserLong1
                
                  If Test_Direction4(temp1, pos_xt, pos_yt) = False Then 'forced direction change
                    If Rnd() < 0.5 Then
                      temp1 = temp1 + 1
                      If temp1 > 4 Then temp1 = 1
                      
                      If Not Test_Direction4(temp1, pos_xt, pos_yt) Then temp1 = temp1 - 2
                      If temp1 < 1 Then temp1 = temp1 + 4
                    Else
                      temp1 = temp1 - 1
                      If temp1 < 1 Then temp1 = 4
                      
                      If Not Test_Direction4(temp1, pos_xt, pos_yt) Then temp1 = temp1 + 2
                      If temp1 > 4 Then temp1 = temp1 - 4
                    End If
                  Else 'flexible direction change
                    temp2 = Rnd()
                      
                      If temp2 < 0.3 Then
                        temp1 = temp1 + 1
                        If temp1 > 4 Then temp1 = 1
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then
                          temp1 = temp1 - 1
                          If temp1 < 1 Then temp1 = 4
                        End If
                      ElseIf temp2 < 0.6 Then
                        temp1 = temp1 - 1
                        If temp1 < 1 Then temp1 = 4
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then
                          temp1 = temp1 + 1
                          If temp1 > 4 Then temp1 = 1
                        End If
                      End If
                  End If
                  
                  .UserLong1 = temp1
                End If
              End If
            Else 'not in chamber - out and about
              temp1 = .UserLong1
              
              If .UserLong2 = 0 Or loop1 = 5 Then 'if ghosts are blue then they are also smart
                loop2 = loop1
              Else
                loop2 = 3
              End If
              
              If Test_Direction4(temp1, pos_xt, pos_yt) = False Then
                Select Case loop2 'forced direction change
                  Case 1 'dumb ghost
                    maxLogicDepth = PlayerData(PlayerActive).GhostIntel(1)
                    logicTargetX = (pos_xtT - 4000) \ 56000
                    logicTargetY = (pos_ytT - 4000) \ 56000
                    
                    If LogicSearch(temp1, 0, (pos_xt - 4000) \ 56000, (pos_yt - 4000) \ 56000) < 9999 Then
                      temp1 = SmartDirection
                    ElseIf Rnd() < 0.5 Then
                      temp1 = temp1 + 1
                      If temp1 > 4 Then temp1 = 1
                      
                      If Not Test_Direction4(temp1, pos_xt, pos_yt) Then temp1 = temp1 - 2
                      If temp1 < 1 Then temp1 = temp1 + 4
                    Else
                      temp1 = temp1 - 1
                      If temp1 < 1 Then temp1 = 4
                      
                      If Not Test_Direction4(temp1, pos_xt, pos_yt) Then temp1 = temp1 + 2
                      If temp1 > 4 Then temp1 = temp1 - 4
                    End If
                  Case 2, 4 'semi intelligent ghost
                    maxLogicDepth = PlayerData(PlayerActive).GhostIntel(loop2)
                    logicTargetX = (pos_xtT - 4000) \ 56000
                    logicTargetY = (pos_ytT - 4000) \ 56000
                    
                    If LogicSearch(temp1, 0, (pos_xt - 4000) \ 56000, (pos_yt - 4000) \ 56000) < 9999 Then
                      temp1 = SmartDirection
                    ElseIf Rnd() < 0.5 Then
                      If Rnd() < 0.5 Then
                        temp1 = temp1 + 1
                        If temp1 > 4 Then temp1 = 1
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then temp1 = temp1 - 2
                        If temp1 < 1 Then temp1 = temp1 + 4
                      Else
                        temp1 = temp1 - 1
                        If temp1 < 1 Then temp1 = 4
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then temp1 = temp1 + 2
                        If temp1 > 4 Then temp1 = temp1 - 4
                      End If
                    Else
                      temp1 = SmartDecision2(temp1, pos_xt, pos_yt, pos_xtT, pos_ytT)
                    End If
                  Case 3 'brain ghost
                    maxLogicDepth = PlayerData(PlayerActive).GhostIntel(3)
                    logicTargetX = (pos_xtT - 4000) \ 56000
                    logicTargetY = (pos_ytT - 4000) \ 56000
                    
                    If LogicSearch(temp1, 0, (pos_xt - 4000) \ 56000, (pos_yt - 4000) \ 56000) < 9999 Then
                      temp1 = SmartDirection
                    Else
                      temp1 = SmartDecision2(temp1, pos_xt, pos_yt, pos_xtT, pos_ytT)
                    End If
                  Case Else 'wraith
                    Select Case temp1
                      Case 1, 3
                        If pos_xtT > pos_xt Then
                          temp1 = 2
                          
                          If Not Test_Direction4(temp1, pos_xt, pos_yt) Then WraithAlertSound True
                        Else
                          temp1 = 4
                          
                          If Not Test_Direction4(temp1, pos_xt, pos_yt) Then WraithAlertSound True
                        End If
                      Case Else
                        If pos_ytT > pos_yt Then
                          temp1 = 3
                          
                          If Not Test_Direction4(temp1, pos_xt, pos_yt) Then WraithAlertSound True
                        Else
                          temp1 = 1
                          
                          If Not Test_Direction4(temp1, pos_xt, pos_yt) Then WraithAlertSound True
                        End If
                    End Select
                End Select
              Else
                Select Case loop2 'flexible direction change
                  Case 1 'dumb ghost
                    maxLogicDepth = PlayerData(PlayerActive).GhostIntel(1)
                    logicTargetX = (pos_xtT - 4000) \ 56000
                    logicTargetY = (pos_ytT - 4000) \ 56000
                    
                    If LogicSearch(temp1, 0, (pos_xt - 4000) \ 56000, (pos_yt - 4000) \ 56000) < 9999 Then
                      temp1 = SmartDirection
                    Else
                      temp2 = Rnd()
                      
                      If temp2 < 0.3 Then
                        temp1 = temp1 + 1
                        If temp1 > 4 Then temp1 = 1
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then
                          temp1 = temp1 - 1
                          If temp1 < 1 Then temp1 = 4
                        End If
                      ElseIf temp2 < 0.6 Then
                        temp1 = temp1 - 1
                        If temp1 < 1 Then temp1 = 4
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then
                          temp1 = temp1 + 1
                          If temp1 > 4 Then temp1 = 1
                        End If
                      End If
                    End If
                  Case 2, 4 'semi intelligent ghost
                    maxLogicDepth = PlayerData(PlayerActive).GhostIntel(loop2)
                    logicTargetX = (pos_xtT - 4000) \ 56000
                    logicTargetY = (pos_ytT - 4000) \ 56000
                    
                    If LogicSearch(temp1, 0, (pos_xt - 4000) \ 56000, (pos_yt - 4000) \ 56000) < 9999 Then
                      temp1 = SmartDirection
                    ElseIf Rnd() < 0.5 Then
                      temp2 = Rnd()
                      
                      If temp2 < 0.3 Then
                        temp1 = temp1 + 1
                        If temp1 > 4 Then temp1 = 1
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then
                          temp1 = temp1 - 1
                          If temp1 < 1 Then temp1 = 4
                        End If
                      ElseIf temp2 < 0.6 Then
                        temp1 = temp1 - 1
                        If temp1 < 1 Then temp1 = 4
                        
                        If Not Test_Direction4(temp1, pos_xt, pos_yt) Then
                          temp1 = temp1 + 1
                          If temp1 > 4 Then temp1 = 1
                        End If
                      End If
                    Else
                      temp1 = SmartDecision(temp1, pos_xt, pos_yt, pos_xtT, pos_ytT)
                    End If
                  Case 3, 4 'brain ghost or wraith
                    maxLogicDepth = PlayerData(PlayerActive).GhostIntel(loop2)
                    logicTargetX = (pos_xtT - 4000) \ 56000
                    logicTargetY = (pos_ytT - 4000) \ 56000
                    
                    If LogicSearch(temp1, 0, (pos_xt - 4000) \ 56000, (pos_yt - 4000) \ 56000) < 9999 Then
                      temp1 = SmartDirection
                    Else
                      temp1 = SmartDecision(temp1, pos_xt, pos_yt, pos_xtT, pos_ytT)
                    End If
                End Select
              End If
              
              .UserLong1 = temp1
            End If
          End If
          
          If GhostFastMode(loop1) = True Then
            Select Case .UserLong1
              Case 1
                .PosY_1000ths = .PosY_1000ths - 5600
                
                If .PosY_1000ths <= -56000 Then .PosY_1000ths = .PosY_1000ths + 496000
              Case 2
                .PosX_1000ths = .PosX_1000ths + 5600
                
                If .PosX_1000ths >= 672000 Then .PosX_1000ths = .PosX_1000ths - 728000
              Case 3
                .PosY_1000ths = .PosY_1000ths + 5600
                
                If .PosY_1000ths >= 448000 Then .PosY_1000ths = .PosY_1000ths - 504000
              Case 4
                .PosX_1000ths = .PosX_1000ths - 5600
                
                If .PosX_1000ths <= -56000 Then .PosX_1000ths = .PosX_1000ths + 728000
            End Select
          Else
            Select Case .UserLong1
              Case 1
                .PosY_1000ths = .PosY_1000ths - PlayerData(PlayerActive).GhostBaseSpeed(loop1)
                
                If .PosY_1000ths <= -56000 Then .PosY_1000ths = .PosY_1000ths + 496000
              Case 2
                .PosX_1000ths = .PosX_1000ths + PlayerData(PlayerActive).GhostBaseSpeed(loop1)
                
                If .PosX_1000ths >= 672000 Then .PosX_1000ths = .PosX_1000ths - 728000
              Case 3
                .PosY_1000ths = .PosY_1000ths + PlayerData(PlayerActive).GhostBaseSpeed(loop1)
                
                If .PosY_1000ths >= 448000 Then .PosY_1000ths = .PosY_1000ths - 504000
              Case 4
                .PosX_1000ths = .PosX_1000ths - PlayerData(PlayerActive).GhostBaseSpeed(loop1)
                
                If .PosX_1000ths <= -56000 Then .PosX_1000ths = .PosX_1000ths + 728000
            End Select
          End If
          
          pos_xt = .PosX_1000ths
          pos_yt = .PosY_1000ths
          
          If pos_x > pos_xt - 24000 And pos_x < pos_xt + 24000 Then
            If pos_y > pos_yt - 24000 And pos_y < pos_yt + 24000 Then 'collided with ghost
              Select Case .UserLong2
                Case 0 'normal ghost
                  With m_SnackMan
                    .ActionSequence = Array(28, 29, 30, 31, 32, 13, 27, 41, 55, 69, 83)
                    .ActionFrame = .ActionSequenceStart - 1
                  End With
                  
                  For loop2 = 1 To 5
                    m_Ghost(loop2).Visible = False
                    m_Scores(loop2).Visible = False
                  Next loop2
                  
                  For loop2 = 1 To 12
                    m_Treat(loop2).Visible = False
                  Next loop2
                  
                  With PlayerData(PlayerActive)
                    .Lives = .Lives - 1
                  End With
                  
                  AllSoundsOff
                  SnackManDeathSound True
                  
                  StoreMap
                  
                  GameMode = 8
                  GameCounter2 = 0
                  
                  Exit Sub
                Case 1 'blue ghost
                  EatGhostCount = EatGhostCount + 1
                  
                  .ActionSequence = Array(11, 25, 39, 53)
                  If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart + loop2 - 1
                  
                  .UserLong2 = 2
                  
                  m_Scores(loop1).PosX_1000ths = .PosX_1000ths
                  m_Scores(loop1).PosY_1000ths = .PosY_1000ths
                  
                  GhostFastMode(loop1) = False
                  
                  Select Case EatGhostCount
                    Case 1
                      PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 100 * BonusMode
                      
                      With m_Scores(loop1)
                        .ActionSequence = Array(57, 57, 57, 57, 71, 71, 71, 71, 85, 85, 85, 85, 99, 99, 99, 99)
                        .ActionFrame = .ActionSequenceStart
                        .Visible = True
                      End With
                    Case 2
                      PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 200 * BonusMode
                      
                      With m_Scores(loop1)
                        .ActionSequence = Array(58, 58, 58, 58, 72, 72, 72, 72, 86, 86, 86, 86, 100, 100, 100, 100)
                        .ActionFrame = .ActionSequenceStart
                        .Visible = True
                      End With
                    Case 3
                      PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 400 * BonusMode
                      
                      With m_Scores(loop1)
                        .ActionSequence = Array(60, 60, 60, 60, 74, 74, 74, 74, 88, 88, 88, 88, 102, 102, 102, 102)
                        .ActionFrame = .ActionSequenceStart
                        .Visible = True
                      End With
                    Case Else
                      PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 1000 * BonusMode
                      
                      With m_Scores(loop1)
                        .ActionSequence = Array(62, 62, 62, 62, 76, 76, 76, 76, 90, 90, 90, 90, 104, 104, 104, 104)
                        .ActionFrame = .ActionSequenceStart
                        .Visible = True
                      End With
                  End Select
                  
                  EatGhostSound loop1, True
                  
                  If EatGhostCount = 4 Then
                    BonusMode = BonusMode + 1
                    BonusValueSound True
                  End If
              End Select
            End If
          End If
        End If
      End With
    Next loop1
    
    For loop1 = 1 To 4 'check power-ups
      With m_PowerUp(loop1)
        If .Visible Then
          pos_xt = .PosX_1000ths
          pos_yt = .PosY_1000ths
          
          If pos_x > pos_xt - 24000 And pos_x < pos_xt + 24000 Then
            If pos_y > pos_yt - 24000 And pos_y < pos_yt + 24000 Then
              .Visible = False
              PlayerData(PlayerActive).SuperDot(loop1, 1) = -1
              
              PowerUpMode = True
              EatGhostCount = 0
              m_SnackMan.UserLong4 = PlayerData(PlayerActive).SuperDotTime
              
              For loop2 = 1 To 4
                'set normals to blue sequence
                With m_Ghost(loop2)
                  If .UserLong2 <> 2 Then
                    .ActionSequence = Array(10, 24, 38, 52)
                    
                    .UserLong2 = 1
                    
                    If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart + loop2 - 1
                  End If
                End With
              Next loop2
              
              EatPowerUpSound True
              PowerUpModeSound True
              
              With PlayerData(PlayerActive)
                .Score = .Score + 100 * BonusMode
                
                .DotCount = .DotCount - 1
              End With
            End If
          End If
        End If
      End With
    Next loop1
    
    For loop1 = 1 To 12 'check treats (collision and move)
      With m_Treat(loop1)
        If .Visible Then
          If GameCounter = 0 Then
            .UserLong4 = .UserLong4 - 1
            
            .UserLong2 = .UserLong2 + 1
            
            If .UserLong2 > 5 Then
              .UserLong2 = 0
              
              TreatBounceSound loop1, True
            End If
          End If
          
          If .UserLong4 <= 0 Then
            TreatMode = PlayerData(PlayerActive).TreatTimeOut
            .Visible = False
          Else 'move to new position
            If .UserLong1 = 0 Then
              If Rnd() < 0.5 Then
                .UserLong1 = 4
              Else
                .UserLong1 = 2
              End If
            End If
            
            If .PosX_1000ths < 4000 Or .PosX_1000ths > 620000 Then
              'skip move logic as treat is out of bounds (going through side path)
              Select Case .UserLong1
                Case 1
                  .UserLong3 = .UserLong3 - 2800
                Case 2
                  .PosX_1000ths = .PosX_1000ths + 2800
                  
                  If .PosX_1000ths >= 672000 Then .PosX_1000ths = .PosX_1000ths - 728000
                Case 3
                  .UserLong3 = .UserLong3 + 2800
                Case 4
                  .PosX_1000ths = .PosX_1000ths - 2800
                  
                  If .PosX_1000ths <= -56000 Then .PosX_1000ths = .PosX_1000ths + 728000
              End Select
            ElseIf (.UserLong3 - 4000) Mod 56000 = 0 And (.PosX_1000ths - 4000) Mod 56000 = 0 Then
              temp2 = Rnd() * 10 + 1
              
              If temp2 > 7 Then
                temp1 = .UserLong1 + 1
                If temp1 > 4 Then temp1 = 1
                
                If Test_Direction4(temp1, .PosX_1000ths, .UserLong3) Then .UserLong1 = temp1
              ElseIf temp2 > 4 Then
                temp1 = .UserLong1 - 1
                If temp1 < 1 Then temp1 = 4
                
                If Test_Direction4(temp1, .PosX_1000ths, .UserLong3) Then .UserLong1 = temp1
              End If
              
              If Test_Direction4(.UserLong1, .PosX_1000ths, .UserLong3) Then
                Select Case .UserLong1
                  Case 1
                    .UserLong3 = .UserLong3 - 2800
                  Case 2
                    .PosX_1000ths = .PosX_1000ths + 2800
                    
                    If .PosX_1000ths >= 672000 Then .PosX_1000ths = .PosX_1000ths - 728000
                  Case 3
                    .UserLong3 = .UserLong3 + 2800
                  Case 4
                    .PosX_1000ths = .PosX_1000ths - 2800
                    
                    If .PosX_1000ths <= -56000 Then .PosX_1000ths = .PosX_1000ths + 728000
                End Select
              End If
            Else
              Select Case .UserLong1
                Case 1
                  .UserLong3 = .UserLong3 - 2800
                Case 2
                  .PosX_1000ths = .PosX_1000ths + 2800
                  
                  If .PosX_1000ths >= 672000 Then .PosX_1000ths = .PosX_1000ths - 728000
                Case 3
                  .UserLong3 = .UserLong3 + 2800
                Case 4
                  .PosX_1000ths = .PosX_1000ths - 2800
                  
                  If .PosX_1000ths <= -56000 Then .PosX_1000ths = .PosX_1000ths + 728000
              End Select
            End If
            
            If .UserLong2 < 4 Then
              .PosY_1000ths = .UserLong3 - .UserLong2 * 6000 + 9000
            Else
              .PosY_1000ths = .UserLong3 - (6 - .UserLong2) * 6000 + 9000
            End If
            
            pos_xt = .PosX_1000ths
            pos_yt = .UserLong3
            
            If pos_x > pos_xt - 12000 And pos_x < pos_xt + 12000 Then
              If pos_y > pos_yt - 12000 And pos_y < pos_yt + 12000 Then
                .Visible = False
                
                EatTreatSound True
                TreatMode = PlayerData(PlayerActive).TreatTimeOut
                
                m_Scores(5).PosX_1000ths = .PosX_1000ths
                m_Scores(5).PosY_1000ths = .PosY_1000ths
                m_Scores(5).Visible = True
                
                Select Case .ObjectTypeID
                  Case 2 'speed
                    m_SnackMan.UserLong3 = PlayerData(PlayerActive).SpeedTime
                    SpeedMode = True
                    
                    With m_Scores(5)
                      .ActionSequence = Array(112, 112, 112, 112, 113, 113, 113, 113, 114, 114, 114, 114, 115, 115, 115, 115)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 3 '50
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 50 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(56, 56, 56, 56, 70, 70, 70, 70, 84, 84, 84, 84, 98, 98, 98, 98)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 4 '100
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 100 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(57, 57, 57, 57, 71, 71, 71, 71, 85, 85, 85, 85, 99, 99, 99, 99)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 5 '200
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 200 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(58, 58, 58, 58, 72, 72, 72, 72, 86, 86, 86, 86, 100, 100, 100, 100)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 6 '300
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 300 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(59, 59, 59, 59, 73, 73, 73, 73, 87, 87, 87, 87, 101, 101, 101, 101)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 7 '400
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 400 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(60, 60, 60, 60, 74, 74, 74, 74, 88, 88, 88, 88, 102, 102, 102, 102)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 8 '500
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 500 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(61, 61, 61, 61, 75, 75, 75, 75, 89, 89, 89, 89, 103, 103, 103, 103)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 9 '1000
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 1000 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(62, 62, 62, 62, 76, 76, 76, 76, 90, 90, 90, 90, 104, 104, 104, 104)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 10 '2000
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 2000 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(63, 63, 63, 63, 77, 77, 77, 77, 91, 91, 91, 91, 105, 105, 105, 105)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case 11 '4000
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 4000 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(64, 64, 64, 64, 78, 78, 78, 78, 92, 92, 92, 92, 106, 106, 106, 106)
                      .ActionFrame = .ActionSequenceStart
                    End With
                  Case Else '8000
                    PlayerData(PlayerActive).Score = PlayerData(PlayerActive).Score + 8000 * BonusMode
                    
                    With m_Scores(5)
                      .ActionSequence = Array(65, 65, 65, 65, 79, 79, 79, 79, 93, 93, 93, 93, 107, 107, 107, 107)
                      .ActionFrame = .ActionSequenceStart
                    End With
                End Select
              End If
            End If
          End If
        End If
      End With
    Next loop1
    
    For loop1 = 1 To 5
      With m_Scores(loop1)
        If .Visible Then
          If GameCounter = 0 Then .ActionFrame = .ActionFrame + 1
          
          If .ActionFrame > .ActionSequenceStop Then
            .Visible = False
          Else
            .PosX_1000ths = .PosX_1000ths + 1000
            .PosY_1000ths = .PosY_1000ths - 1000
          End If
        End If
      End With
    Next loop1
  End If
  
  With PlayerData(PlayerActive)
    If .Score >= .NextMan Then
      .Lives = .Lives + 1
      
      Select Case .NextMan
        Case 10000
          .NextMan = 20000
        Case 20000
          .NextMan = 30000
        Case 30000
          .NextMan = 40000
        Case 40000
          .NextMan = 55000
        Case Else
          .NextMan = .NextMan + 20000
      End Select
    End If
  End With
  
  With PlayerData(PlayerActive)
    Display_Score .Lives, .Score, .MapLevel, False
  End With
  
  Display_Specials PlayerActive, PowerUpMode, SpeedMode, BonusMode, WraithMode, False
  
  UpdateAnim
End Sub

Private Sub UpdateAnim()
  Dim loop1 As Long
  
  GameCounter = GameCounter + 1
  
  If GameCounter >= 5 Then
    GameCounter = 0
    
    DXext.BGMapBaseImageIndex = DXext.BGMapBaseImageIndex + 1
    If DXext.BGMapBaseImageIndex > 3 Then DXext.BGMapBaseImageIndex = 0
      
    For loop1 = 1 To 5
      With m_Ghost(loop1)
        .ActionFrame = .ActionFrame + 1
          
        If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
      End With
    Next loop1
    
    For loop1 = 1 To 4
      With m_PowerUp(loop1)
        .ActionFrame = .ActionFrame + 1
        
        If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
      End With
    Next loop1
  End If
  
  With m_SnackMan
    .ActionFrame = .ActionFrame + 1
    
    If .ActionFrame > .ActionSequenceStop Then .ActionFrame = .ActionSequenceStart
  End With
End Sub

'Game over
Public Sub GameMode_9()
  Dim loop1 As Long
  
  GameCounter2 = GameCounter2 + 1
  If GameCounter2 > 80 Then
    GameCounter2 = 0
    
    If PlayerActive = 1 Then
      PlayerActive = 2
    Else
      PlayerActive = 1
    End If
    
    With PlayerData(PlayerActive)
      Display_Score 0, .Score, .MapLevel, True
      Display_Specials PlayerActive, False, False, 1, False, True
    End With
  End If
  
  If dx_KeyboardState.Key(DIK_SPACE) <> 0 Then
    GameMode = 2
  ElseIf dx_KeyboardState.Key(DIK_F1) <> 0 Then
    'start 1 player game
    
    GameMode_4 1
  ElseIf dx_KeyboardState.Key(DIK_F2) <> 0 Then
    'start 2 player game
    
    GameMode_4 2
  ElseIf dx_JoystickState.buttons(0) <> 0 Then
    GameMode = 2
  ElseIf dx_JoystickState.buttons(1) <> 0 Then
    GameMode_4 1
  ElseIf dx_JoystickState.buttons(2) <> 0 Then
    GameMode_4 2
  End If
End Sub

Public Function LogicSearch(ByVal SourceDirection As Long, ByVal logicDepth As Long, ByVal SourceX As Long, ByVal SourceY As Long) As Long
  Dim results2 As Long, results3 As Long, results4 As Long
  Dim ElementValue As Long
  
  logicDepth = logicDepth + 1
  
  If logicDepth > maxLogicDepth Then
    LogicSearch = 9999
  ElseIf SourceX = logicTargetX And SourceY = logicTargetY Then
    If logicDepth = 1 Then
      LogicSearch = 10000
    Else
      LogicSearch = logicDepth
    End If
  Else
    ElementValue = DXext.BGMapElement(SourceY, SourceX)
    
    If ElementValue < 96 Then ElementValue = ElementValue Mod 12
    
    If SourceY > 0 And SourceDirection <> 3 Then
      Select Case ElementValue
        Case 1, 3, 4, 6, 7, 10, 11
          LogicSearch = LogicSearch(1, logicDepth, SourceX, SourceY - 1)
        Case Else
          LogicSearch = 9999
      End Select
    Else
      LogicSearch = 9999
    End If
    
    If SourceY < 7 And SourceDirection <> 1 Then
      Select Case ElementValue
        Case 1, 4, 5, 6, 8, 9, 11
          results3 = LogicSearch(3, logicDepth, SourceX, SourceY + 1)
        Case Else
          results3 = 9999
      End Select
    Else
      results3 = 9999
    End If
    
    If SourceX > 0 And SourceDirection <> 2 Then
      Select Case ElementValue
        Case 2, 3, 5, 6, 9, 10, 11
          results4 = LogicSearch(4, logicDepth, SourceX - 1, SourceY)
        Case Else
          results4 = 9999
      End Select
    Else
      results4 = 9999
    End If
    
    If SourceX < 11 And SourceDirection <> 4 Then
      Select Case ElementValue
        Case 2, 3, 4, 5, 7, 8, 11
          results2 = LogicSearch(2, logicDepth, SourceX + 1, SourceY)
        Case Else
          results2 = 9999
      End Select
    Else
      results2 = 9999
    End If
    
    SmartDirection = 1
    
    If results2 < LogicSearch Then
      LogicSearch = results2
      SmartDirection = 2
    End If
    
    If results3 < LogicSearch Then
      LogicSearch = results3
      SmartDirection = 3
    End If
    
    If results4 < LogicSearch Then
      LogicSearch = results4
      SmartDirection = 4
    End If
  End If
End Function

Public Function SmartDecision(ByVal originalDirection As Long, ByVal SourceX As Long, ByVal SourceY As Long, ByVal TargetX As Long, ByVal TargetY As Long) As Long
  SmartDecision = originalDirection
  
  Select Case originalDirection
    Case 1, 3
      If TargetX > SourceX Then
        If TargetY = SourceY Or Rnd() < 0.5 Then
          If Test_Direction4(2, SourceX, SourceY) Then SmartDecision = 2
        End If
      ElseIf TargetX < SourceX Then
        If TargetY = SourceY Or Rnd() < 0.5 Then
          If Test_Direction4(4, SourceX, SourceY) Then SmartDecision = 4
        End If
      End If
    Case Else
      If TargetY > SourceY Then
        If TargetX = SourceX Or Rnd() < 0.5 Then
          If Test_Direction4(3, SourceX, SourceY) Then SmartDecision = 3
        End If
      ElseIf TargetY < SourceY Then
        If TargetX = SourceX Or Rnd() < 0.5 Then
          If Test_Direction4(1, SourceX, SourceY) Then SmartDecision = 1
        End If
      End If
  End Select
End Function

Public Function SmartDecision2(ByVal originalDirection As Long, ByVal SourceX As Long, ByVal SourceY As Long, ByVal TargetX As Long, ByVal TargetY As Long) As Long
  Select Case originalDirection
    Case 1, 3
      If TargetX > SourceX Then
        SmartDecision2 = 2
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 4
      ElseIf TargetX < SourceX Then
        SmartDecision2 = 4
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 2
      ElseIf Rnd() < 0.5 Then
        SmartDecision2 = 2
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 4
      Else
        SmartDecision2 = 4
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 2
      End If
    Case Else
      If TargetY > SourceY Then
        SmartDecision2 = 3
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 1
      ElseIf TargetY < SourceY Then
        SmartDecision2 = 1
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 3
      ElseIf Rnd() < 0.5 Then
        SmartDecision2 = 3
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 1
      Else
        SmartDecision2 = 1
        
        If Not Test_Direction4(SmartDecision2, SourceX, SourceY) Then SmartDecision2 = 3
      End If
  End Select
End Function

Public Sub Display_Border()
  Dim loop1 As Long
  
  For loop1 = 1 To 13
    DXext.DisplayStaticImageSolid 1, 504, 480, 48, 48, 6 + loop1 * 48, 6, True
    DXext.DisplayStaticImageSolid 1, 504, 528, 48, 48, 6 + loop1 * 48, 438, True
  Next loop1
  
  For loop1 = 1 To 8
    DXext.DisplayStaticImageSolid 1, 480, 504, 48, 48, 6, 6 + loop1 * 48, True
    DXext.DisplayStaticImageSolid 1, 528, 504, 48, 48, 662, 6 + loop1 * 48, True
  Next loop1
  
  DXext.DisplayStaticImageSolid 1, 480, 480, 48, 48, 6, 6, True
  DXext.DisplayStaticImageSolid 1, 528, 480, 48, 48, 662, 6, True
  DXext.DisplayStaticImageSolid 1, 480, 528, 48, 48, 6, 438, True
  DXext.DisplayStaticImageSolid 1, 528, 528, 48, 48, 662, 438, True
  
  DXext.DisplayStaticImageSolid 1, 384, 96, 96, 48, 0, 488, True 'lives
  DXext.DisplayStaticImageSolid 2, 48, 0, 48, 48, 96, 488, True  'a snackman
  
  DXext.DisplayStaticImageSolid 1, 0, 672, 96, 48, 246, 488, True 'level
  
  DXext.DisplayStaticImageSolid 1, 384, 144, 96, 48, 430, 488, True 'score
    
  DXext.SyncronizeBuffers
End Sub

Private Sub mnuController_Click(Index As Integer)
  mnuController(activeController).Checked = False
  activeController = Index
  mnuController(activeController).Checked = True
End Sub

Private Sub mnuExit_Click()
  VerifyExit = False
  
  Unload Me
End Sub

Private Sub mnuRestart_Click()
  GameMode = 2
  
  mnuResume_Click
End Sub

Public Sub mnuResume_Click()
  mnuGame.Visible = False
  mnuDisplay.Visible = False
  mnuControllerTitle.Visible = False
  Me.Width = 11985
  Me.Height = 8745
  
    If ScreenSizeMode >= 1 Then
      If ScreenSizeMode = 1 Then
        DXext.Init_DXDrawScreen frmMain, 800, 600, 16, 60, , True
      ElseIf ScreenSizeMode = 2 Then
        DXext.Init_DXDrawScreen frmMain, 800, 600, 32, 60, , True
      Else
        DXext.Init_DXDrawScreen frmMain, 800, 600, 8, 60, , True
      End If
      
      TimeLoop = DelayTillTime(50, , True)
      
      DXext.ClearDisplay True
      
      Display_Border
      Display_Score 0, 0, 0, True
      Display_Specials 0, False, False, 1, False, True
    ElseIf Not DXext.TestDisplayValid() Then
      DXext.Init_DXDrawWindow frmMain, picWindow, , True
      
      TimeLoop = DelayTillTime(50, , True)
      
      DXext.ClearDisplay True
      
      Display_Border
      Display_Score 0, 0, 0, True
      Display_Specials 0, False, False, 1, False, True
    End If
  
  Reaquire_DXInput
  
  UnPauseAllSound
  
  TimeLoop = DelayTillTime(0, , True)
  PauseMode = False
End Sub

Private Sub mnuScreenMode_Click(Index As Integer)
  mnuScreenmode(ScreenSizeMode).Checked = False
  ScreenSizeMode = Index
  mnuScreenmode(ScreenSizeMode).Checked = True
End Sub

Public Sub Reaquire_DXInput()
  Init_DXInput Me, True, activeController
  
  Set_JoystickRange 10000, True, True
  Set_JoystickDeadZoneSat 1000, 8000, True, True
End Sub
