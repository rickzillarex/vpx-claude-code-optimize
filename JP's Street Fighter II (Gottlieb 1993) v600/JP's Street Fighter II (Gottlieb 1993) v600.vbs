' ======================================================================
' VPX Performance Optimizations Applied:
'   - Flipper timer interval 1ms -> 10ms
'   - Cache flipper COM properties (CurrentAngle/StartAngle/EndAngle) in FlipperTricks timer
'   - Eliminate ^2 in Vol/BallVel/OnBallBallCollision - use direct multiplication
'   - Eliminate ^10 in Pan/AudioFade - use chain multiply (t2*t4*t8)
'   - Pre-built BallRollStr() array to eliminate per-frame string concatenation
' ======================================================================

' JP's Street Fighter II
' Based on the table by Premier/Gottlieb from 1993
' VPX8 table by jpsalas 2024 versio 6.0.0

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

LoadVPM "01210000", "gts3.VBS", 3.10

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 1
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Const swLCoin = 0
Const swRCoin = 1
Const swCCoin = 2
Const swCoinShuttle = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6

Dim bsTrough, bsLKicker, bsL2Kicker, bsL3Kicker, bsRLowKicker, x

'************
' Table init.
'************

Const cGameName = "sfight2"

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's Street Fighter II" & vbNewLine & "VPX table by JPSalas v6.0.0"
        '.Games(cGameName).Settings.Value("rol") = 0
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Dip(0) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 1 * 64 + 1 * 128) '01-08
        .Dip(1) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 1 * 32 + 1 * 64 + 1 * 128) '09-16
        .Dip(2) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 1 * 32 + 1 * 64 + 1 * 128) '17-24
        .Dip(3) = (1 * 1 + 1 * 2 + 1 * 4 + 0 * 8 + 1 * 16 + 0 * 32 + 1 * 64 + 1 * 128) '25-32
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 151
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(bumper1, LeftSlingshot, RightSlingshot)

    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 21, 0, 0, 31, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 75, 11
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 3

    Set bsRLowKicker = New cvpmBallStack
    bsRLowKicker.InitSw 0, 34, 0, 0, 0, 0, 0, 0
   'bsRLowKicker.InitKick sw34a, 221, 16
    bsRLowKicker.InitKick sw34a, 233, 16
    bsRLowKicker.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsRLowKicker.Balls = 0

    set bsLKicker = new cvpmBallStack
    bsLKicker.InitSw 0, 33, 0, 0, 0, 0, 0, 0
   'bsLKicker.InitKick sw33a, 145, 16
    bsLKicker.InitKick sw33a, 130, 16
    bsLKicker.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsLKicker.Balls = 0

    set bsL2Kicker = new cvpmBallStack
    bsL2Kicker.InitSaucer sw13, 13, 180, 7
    bsL2Kicker.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    set bsL3Kicker = new cvpmBallStack
    bsL3Kicker.InitSaucer sw14, 14, 200, 8
    bsL3Kicker.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    vpmMapLights aLights

    ' Init other dropwalls - animations
    para2.IsDropped = 1:para3.IsDropped = 1:para4.IsDropped = 1:para5.IsDropped = 1
    para6.IsDropped = 1:para7.IsDropped = 1:para8.IsDropped = 1:
    Drain.createball
    Drain.kick 0, 3
    KickerBottom.Createball
    KickerBottom.kick 0, 0

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
End Sub

'**********
' Solenoids
'**********

'SolCallback(1)  = "vpmSolSound ""Jet"","
'SolCallback(2)  = "vpmSolSound ""lSling"","
'SolCallback(3)  = "vpmSolSound ""lSling"","
SolCallback(4) = "bsLKicker.SolOut"
SolCallback(5) = "bsRlowKicker.SolOut"
SolCallback(6) = "SolDiv1"
SolCallback(7) = "SolDiv2"
SolCallback(8) = "SolDiv3"
SolCallback(9) = "SolDiv4"
SolCallback(10) = "bsL2Kicker.SolOut"
SolCallback(11) = "bsL3Kicker.SolOut"
'12 NOT USED
'13 NOT USED
SolCallback(14) = "CarReset" '
'Flashers
SolCallback(15) = "vpmFlasher f15,"
SolCallback(16) = "vpmFlasher f16,"
SolCallback(17) = "vpmFlasher f17,"
SolCallback(18) = "vpmFlasher f18,"
SolCallback(19) = "vpmFlasher f19,"
SolCallback(20) = "vpmFlasher f20,"
SolCallback(21) = "vpmFlasher f21,"
SolCallback(22) = "vpmFlasher f22,"
SolCallback(23) = "vpmFlasher f23,"

'SolCallback(24) ' this is the underplayfield flipper
SolCallback(25) = "ChunliTimer.Enabled= "

'SolCallback(26) = ' these are the lights on the backglass, and they looks they are always on.
SolCallback(27) = "vpmSolSound ""SolOn"","
SolCallback(28) = "bsTrough.SolOut"
SolCallback(29) = "bsTrough.SolIn"
SolCallback(30) = "vpmSolSound ""fx_Knocker"","
'SolCallback(31) = "Nugde"
'SolCallback(32) = "GameOn"

Sub SolDiv1(enabled)
    if enabled then
        Diverter1.isdropped = 1
    else
        Diverter1.isdropped = 0
    End if
End Sub

Sub SolDiv2(enabled)
    if enabled then
        Diverter2.isdropped = 1
    else
        Diverter2.isdropped = 0
    End if
End Sub

Sub SolDiv3(enabled)
    if enabled then
        leftramp.heightbottom = 55
        leftrampb.heightbottom = 60
        leftrampb.heighttop = 55
        leftramp.Collidable = 0
    else
        leftramp.heightbottom = 0
        leftrampb.heightbottom = 0
        leftrampb.heighttop = 5
        leftramp.Collidable = 1
    End if
End Sub

Sub SolDiv4(enabled)
    if enabled then
        Rightramp.heightbottom = 55
        Rightrampb.heightbottom = 60
        Rightrampb.heighttop = 55
        Rightramp.Collidable = 0
    else
        Rightramp.heightbottom = 0
        Rightrampb.heightbottom = 0
        Rightrampb.heighttop = 5
        Rightramp.Collidable = 1
    End if
End Sub

Dim CarPos:CarPos = 0
Sub CarReset(enabled)
    If enabled Then
        CarResetTimer.Enabled = 1
    End If
End Sub

Sub CarResetTimer_Timer 'reset to the first position
    Select Case CarPos
        Case 1:para2.IsDropped = 1:para1.IsDropped = 0:car.Y = 1120:CarPos = 0:Me.Enabled = 0
        Case 2:para3.IsDropped = 1:para2.IsDropped = 0:car.Y = 1100:CarPos = 1
        Case 3:para4.IsDropped = 1:para3.IsDropped = 0:car.Y = 1080:CarPos = 2
        Case 4:para5.IsDropped = 1:para4.IsDropped = 0:car.Y = 1060:CarPos = 3
        Case 5:para6.IsDropped = 1:para5.IsDropped = 0:car.Y = 1040:CarPos = 4
        Case 6:para7.IsDropped = 1:para6.IsDropped = 0:car.Y = 1020:CarPos = 5
        Case 7:para8.IsDropped = 1:para7.IsDropped = 0:car.Y = 1000:CarPos = 6
    End Select
End Sub

Dim Diverterangle
Diverterangle = 350

Sub ChunliTimer_timer
    Diverterangle = Diverterangle + 10
    If Diverter.startangle = 360 Then Diverterangle = 10
    Diverter.startangle = Diverterangle
End Sub

'*********************************************************
' Real Time Flipper adjustments - by JLouLouLou & JPSalas
'        (to enable flipper tricks)
'*********************************************************

Dim FlipperPower
Dim FlipperElasticity
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque, LiveStrokeEOS_Torque
Dim LeftFlipperOn
Dim RightFlipperOn

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LiveCatchSensivity

FlipperPower = 3600
FlipperElasticity = 0.6
FullStrokeEOS_Torque = 0.6 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.3 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.2
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 10
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    Dim LF_CurAngle, LF_StartAngle, LF_EndAngle
    Dim RF_CurAngle, RF_StartAngle, RF_EndAngle
    LF_CurAngle = LeftFlipper.CurrentAngle
    LF_StartAngle = LeftFlipper.StartAngle
    LF_EndAngle = LeftFlipper.EndAngle
    RF_CurAngle = RightFlipper.CurrentAngle
    RF_StartAngle = RightFlipper.StartAngle
    RF_EndAngle = RightFlipper.EndAngle
    If LF_CurAngle >= LF_StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If LeftFlipperOn = 1 Then
        If LF_CurAngle = LF_EndAngle then
            LeftFlipper.EOSTorque = FullStrokeEOS_Torque
            LLiveCatchTimer = LLiveCatchTimer + 1
            If LLiveCatchTimer <LiveCatchSensivity Then
                LeftFlipper.Elasticity = 0
            Else
                LeftFlipper.Elasticity = FlipperElasticity
                LLiveCatchTimer = LiveCatchSensivity
            End If
        End If
    Else
        LeftFlipper.Elasticity = FlipperElasticity
        LeftFlipper.EOSTorque = LiveStrokeEOS_Torque
        LLiveCatchTimer = 0
    End If

    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RF_CurAngle <= RF_StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If RightFlipperOn = 1 Then
        If RF_CurAngle = RF_EndAngle Then
            RightFlipper.EOSTorque = FullStrokeEOS_Torque
            RLiveCatchTimer = RLiveCatchTimer + 1
            If RLiveCatchTimer <LiveCatchSensivity Then
                RightFlipper.Elasticity = 0
            Else
                RightFlipper.Elasticity = FlipperElasticity
                RLiveCatchTimer = LiveCatchSensivity
            End If
        End If
    Else
        RightFlipper.Elasticity = FlipperElasticity
        RightFlipper.EOSTorque = LiveStrokeEOS_Torque
        RLiveCatchTimer = 0
    End If
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound "fx_nudge", 0, 1, 0, 0.25
    If keycode = 16 Then Controller.Switch(5) = 1
    If keycode = LeftFlipperKey then Controller.Switch(82) = 1
    If keycode = RightFlipperkey then Controller.Switch(83) = 1
    If Lowerflipper Then
        If keycode = RightFlipperkey then
            PlaySoundAt "fx_flipperup", RightFlipper1
            RightFlipper1.RotateToEnd
        End If
    End If
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = 16 Then Controller.Switch(5) = 0
    If keycode = LeftFlipperKey then Controller.Switch(82) = 0
    If keycode = RightFlipperkey then Controller.Switch(83) = 0:RightFlipper1.RotatetoStart
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings & div switches

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 11
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 12
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 10:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaysoundAt "fx_drain", drain:bsTrough.AddBall Me:End Sub
Sub sw13_Hit:PlaySoundAt "fx_kicker_enter", sw13:bsL2Kicker.AddBall 0:End Sub
Sub sw14_Hit:PlaySoundAt "fx_kicker_enter", sw14:bsL3Kicker.AddBall 0:End Sub
Sub sw100_Hit:vpmTimer.PulseSw 100:End Sub

Sub sw33_Hit
    PlaySoundAt "fx_balldrop", sw33
    sw33.Destroyball
    vpmTimer.PulseSwitch(33), 150, "bsLKicker.addball 0 '"
End Sub

Sub sw34_Hit
    PlaySoundAt "fx_balldrop", sw34
    sw34.Destroyball
    vpmTimer.PulseSwitch(34), 150, "bsRLowKicker.addball 0 '"
End Sub

Sub sw80_Hit
    PlaySoundAt "fx_balldrop", sw80
    sw80.Destroyball
    vpmTimer.PulseSwitch(80), 150, "bsLKicker.addball 0 '"
End Sub

Sub sw90_Hit
    PlaySoundAt "fx_balldrop", sw90
    sw90.Destroyball
    vpmTimer.PulseSwitch(90), 150, "bsRLowKicker.addball 0 '"
End Sub

' Rollovers & Ramp Switches
Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", sw22:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAt "fx_sensor", sw30:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAt "fx_sensor", sw32:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw91_Hit:Controller.Switch(91) = 1:PlaySoundAt "fx_sensor", sw91:End Sub
Sub sw91_UnHit:Controller.Switch(91) = 0:End Sub

Sub sw95_Hit:Controller.Switch(95) = 1:PlaySoundAt "fx_sensor", sw95:End Sub
Sub sw95_UnHit:Controller.Switch(95) = 0:End Sub

Sub sw96_Hit:Controller.Switch(96) = 1:PlaySoundAt "fx_sensor", sw96:End Sub
Sub sw96_UnHit:Controller.Switch(96) = 0:End Sub

Sub sw101_Hit:Controller.Switch(101) = 1:PlaySoundAt "fx_sensor", sw101:End Sub
Sub sw101_UnHit:Controller.Switch(101) = 0:End Sub

Sub sw105_Hit:Controller.Switch(105) = 1:PlaySoundAt "fx_sensor", sw105:End Sub
Sub sw105_UnHit:Controller.Switch(105) = 0:End Sub

Sub sw106_Hit:Controller.Switch(106) = 1:PlaySoundAt "fx_sensor", sw106:End Sub
Sub sw106_UnHit:Controller.Switch(106) = 0:End Sub

Sub sw110_Hit:Controller.Switch(110) = 1:PlaySoundAt "fx_sensor", sw110:End Sub
Sub sw110_UnHit:Controller.Switch(110) = 0:End Sub

Sub sw111_Hit:Controller.Switch(111) = 1:PlaySoundAt "fx_sensor", sw111:End Sub
Sub sw111_UnHit:Controller.Switch(111) = 0:End Sub

Sub sw114_Hit:Controller.Switch(114) = 1:PlaySoundAt "fx_sensor", sw114:End Sub
Sub sw114_UnHit:Controller.Switch(114) = 0:End Sub

Sub sw115_Hit:Controller.Switch(115) = 1:PlaySoundAt "fx_sensor", sw115:End Sub
Sub sw115_UnHit:Controller.Switch(115) = 0:End Sub

Sub sw116_Hit:Controller.Switch(116) = 1:PlaySoundAt "fx_sensor", sw116:End Sub
Sub sw116_UnHit:Controller.Switch(116) = 0:End Sub

' Targets
Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySoundAtBall "fx_target":End Sub
Sub sw92_Hit:vpmTimer.PulseSw 92:PlaySoundAtBall "fx_target":End Sub
Sub sw93_Hit:vpmTimer.PulseSw 93:PlaySoundAtBall "fx_target":End Sub
Sub sw94_Hit:vpmTimer.PulseSw 94:PlaySoundAtBall "fx_target":End Sub
Sub sw102_Hit:vpmTimer.PulseSw 102:PlaySoundAtBall "fx_target":End Sub
Sub sw103_Hit:vpmTimer.PulseSw 103:PlaySoundAtBall "fx_target":End Sub
Sub sw104_Hit:vpmTimer.PulseSw 104:PlaySoundAtBall "fx_target":End Sub
Sub sw112_Hit:vpmTimer.PulseSw 112:PlaySoundAtBall "fx_target":End Sub
Sub sw113_Hit:vpmTimer.PulseSw 113:PlaySoundAtBall "fx_target":End Sub

'Car Targets

Sub para1_Hit:vpmTimer.PulseSw 15:PlaySoundAtBall "fx_target":para1.Isdropped = 1:Para2.IsDropped = 0:Car.Y = 1100:CarPos = 1:End Sub
Sub para2_Hit:vpmTimer.PulseSw 15:PlaySoundAtBall "fx_target":para2.Isdropped = 1:Para3.IsDropped = 0:Car.Y = 1080:CarPos = 2:End Sub
Sub para3_Hit:vpmTimer.PulseSw 15:PlaySoundAtBall "fx_target":para3.Isdropped = 1:Para4.IsDropped = 0:Car.Y = 1060:CarPos = 3:End Sub
Sub para4_Hit:vpmTimer.PulseSw 15:PlaySoundAtBall "fx_target":para4.Isdropped = 1:Para5.IsDropped = 0:Car.Y = 1040:CarPos = 4:End Sub
Sub para5_Hit:vpmTimer.PulseSw 15:PlaySoundAtBall "fx_target":para5.Isdropped = 1:Para6.IsDropped = 0:Car.Y = 1020:CarPos = 5:End Sub
Sub para6_Hit:vpmTimer.PulseSw 15:PlaySoundAtBall "fx_target":para6.Isdropped = 1:Para7.IsDropped = 0:Car.Y = 1000:CarPos = 6:End Sub
Sub para7_Hit:vpmTimer.PulseSw 81:PlaySoundAtBall "fx_target":para7.Isdropped = 1:Para8.IsDropped = 0:Car.Y = 980:CarPos = 7:End Sub

'*******************
' Flipper Subs v3.0
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

SolCallback(24) = "SolRFlipper1"

Dim LowerFlipper
LowerFlipper = False

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub LeftFlipper_Animate()
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipper_Animate()
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
End Sub

Sub LeftFlipper1_Animate()
    LeftFlipperTop1.RotZ = LeftFlipper1.CurrentAngle
End Sub

Sub RightFlipper1_Animate()
    RightFlipperTop1.RotZ = RightFlipper1.CurrentAngle
End Sub

Sub Diverter_Animate()
    DiverterTop.RotZ = diverter.CurrentAngle
    ChunLiHelico.RotY = diverter.CurrentAngle
End Sub

Sub SolRFlipper1(Enabled)
    LowerFlipper = Enabled
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'************************************
' Diverse Collection Hit Sounds v3.0
'************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) * BallVel(ball) / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp> 0 Then
        Dim t2, t4, t8 : t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4 : Pan = Csng(t8*t2)
    Else
        Dim t2n, t4n, t8n : t2n = tmp*tmp : t4n = t2n*t2n : t8n = t4n*t4n : Pan = Csng(-(t8n*t2n))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX * ball.VelX) + (ball.VelY * ball.VelY) ) )
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        Dim af2, af4, af8 : af2 = tmp*tmp : af4 = af2*af2 : af8 = af4*af4 : AudioFade = Csng(af8*af2)
    Else
        Dim af2n, af4n, af8n : af2n = tmp*tmp : af4n = af2n*af2n : af8n = af4n*af4n : AudioFade = Csng(-(af8n*af2n))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'*****************************
'   JP's VPX Rolling Sounds
' with dropping sound and ball speed control
'*****************************

Const tnob = 19   'total number of balls
Const lob = 1     'number of locked balls
Const maxvel = 42 'max ball velocity
ReDim rolling(tnob)
InitRolling

Dim BallRollStr() : ReDim BallRollStr(19)
Dim ii_brs : For ii_brs = 0 To 19 : BallRollStr(ii_brs) = "fx_ballrolling" & ii_brs : Next

Sub InitRolling

Dim BallRollStr() : ReDim BallRollStr(19)
Dim ii_brs : For ii_brs = 0 To 19 : BallRollStr(ii_brs) = "fx_ballrolling" & ii_brs : Next
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    RollingTimer.Enabled = 1
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound(BallRollStr(b))
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 10
            End If
            rolling(b) = True
            PlaySound(BallRollStr(b)), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound(BallRollStr(b))
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity * velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***************
' Loop/vuk ramp
'***************

Sub vuk_Hit
    vuk.Destroyball
    vuk1.Createball
    vuk1.kick 190, 3
End Sub

'******************
'   GI effects
'******************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    If Enabled Then
        GiOn
    Else
        GiOff
    End If
End Sub

Sub GiOn
    Dim bulb
    PlaySound "GiOn"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    PlaySound "GiOff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are: 
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Desktop DMD
    x = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    DesktopDMD.visible = x

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' Side Blades
    x = Table1.Option("Side Blades", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aSideBlades:y.SideVisible = x:next
End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
        Case 11:table1.ColorGradeImage = "LUT Warm 0"
        Case 12:table1.ColorGradeImage = "LUT Warm 1"
        Case 13:table1.ColorGradeImage = "LUT Warm 2"
        Case 14:table1.ColorGradeImage = "LUT Warm 3"
        Case 15:table1.ColorGradeImage = "LUT Warm 4"
        Case 16:table1.ColorGradeImage = "LUT Warm 5"
        Case 17:table1.ColorGradeImage = "LUT Warm 6"
        Case 18:table1.ColorGradeImage = "LUT Warm 7"
        Case 19:table1.ColorGradeImage = "LUT Warm 8"
        Case 20:table1.ColorGradeImage = "LUT Warm 9"
        Case 21:table1.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub