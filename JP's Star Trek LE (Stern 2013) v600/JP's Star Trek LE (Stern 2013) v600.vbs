' ======================================================================
' VPX Performance Optimizations Applied:
'   - Flipper timer interval 1ms -> 10ms
'   - Cache flipper COM properties (CurrentAngle/StartAngle/EndAngle) in FlipperTricks timer
'   - Eliminate ^2 in Vol/BallVel/OnBallBallCollision - use direct multiplication
'   - Eliminate ^10 in Pan/AudioFade - use chain multiply (t2*t4*t8)
'   - Pre-built BallRollStr() array to eliminate per-frame string concatenation
'   - Cache chgLamp(ii,0)/(ii,1) into locals cIdx/cVal in LampTimer
' ======================================================================

' JP's Star Trek - Limited Edition (Stern 2013)
' based on the promo pictures + some extra decorations.
' IPD No. 6045 / December, 2013 / 4 Players

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, dtBank, bsSaucer, vMagnet, vMagnet2, plungerIM, x

Const cGameName = "st_162h"

Dim DesktopMode, UseVPMColoredDMD
If Table1.ShowDT Then
    UseVPMColoredDMD = True
    DesktopMode = True
    For Each x in aReels
        x.visible = 1
    Next
Else
    UseVPMColoredDMD = False
    DesktopMode = False
    For Each x in aReels
        x.visible = 0
    Next
End If

Const UseVPMModSol=True 'use modulated flashers

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

LoadVPM "01550000", "SAM.VBS", 3.26

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Set GICallback = GetRef("GIUpdate")

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Star Trek LE - Stern 2013" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
        .Games(cGameName).Settings.Value("rol") = 0
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swtilt
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 21, 20, 19, 18, 0, 0, 0
        .InitKick BallRelease, 90, 10
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
    End With

    ' Saucer
    Set bsSaucer = New cvpmSaucer
    With bsSaucer
        .InitKicker sw10, 10, 0, 54, 1.56
    '.InitAltKick 180, 25, 25 'down
    '.InitExitVariance 5, 5
    End With

    ' Drop target
    Set dtBank = New cvpmDropTarget
    With dtBank
        .InitDrop sw11, 11
        .InitSnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBank"
    End With

    ' Magnet
    Set vMagnet = New cvpmMagnet
    with vMagnet
        .InitMagnet VengMagnet, 30
        .GrabCenter = True
        .solenoid = 3
        .CreateEvents "vMagnet"
    End With

    ' Latch Magnet
    Set vMagnet2 = New cvpmMagnet
    with vMagnet2
        .InitMagnet LatchMagnet, 20
        .GrabCenter = True
        '.solenoid = 56
        .CreateEvents "vMagnet2"
    End With

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 62 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw23, IMPowerSetting, IMTime
        .Random 0.3
        .switch 23
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init Kickback
    KickBack.Pullback

    RealTime.Enabled = 1
End Sub

Sub table1_Paused:Controller.Pause = True:End Sub
Sub table1_unPaused:Controller.Pause = False:End Sub
Sub table1_exit:Controller.Pause = False:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If Keycode = LockBarKey then Controller.Switch(71) = 1
    If keycode = RightMagnaSave Then Controller.Switch(71) = 1
    If Keycode = LeftFlipperKey then Controller.Switch(84) = 1
    If Keycode = RightFlipperKey then Controller.Switch(86) = 1:Controller.Switch(82) = 1
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If Keycode = LockBarKey then Controller.Switch(71) = 0
    If keycode = RightMagnaSave Then Controller.Switch(71) = 0
    If Keycode = LeftFlipperKey then Controller.Switch(84) = 0
    If Keycode = RightFlipperKey then Controller.Switch(86) = 0:Controller.Switch(82) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    DOF 101, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 26
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
    DOF 102, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 27
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

' Scoring rubbers

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 30:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw10_Hit:PlaysoundAt "fx_kicker_enter", sw10:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAt "fx_sensor", sw52:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw51_Hit:Controller.Switch(51) = 1:PlaySoundAt "fx_sensor", sw51:End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

Sub sw1_Hit:Controller.Switch(1) = 1:PlaySoundAt "fx_sensor", sw1:End Sub
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub

Sub sw2_Hit:Controller.Switch(2) = 1:PlaySoundAt "fx_sensor", sw2:End Sub
Sub sw2_UnHit:Controller.Switch(2) = 0:End Sub

Sub sw4_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", sw4:End Sub
Sub sw4_UnHit:Controller.Switch(4) = 0:End Sub

'optos

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

'ramps

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor", sw14:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

'Spinners

Sub sw12_Spin():vpmTimer.PulseSw 5:PlaySoundAt "fx_spinner", sw12:End Sub

'Targets
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw49_Hit:vpmTimer.PulseSw 49:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw50_Hit:vpmTimer.PulseSw 50:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw7_Hit:vpmTimer.PulseSw 7:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw9_Hit:vpmTimer.PulseSw 9:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "SolTrough"
SolCallback(2) = "SolAutofire"
'3 magnet
SolCallback(4) = "dtBank.SolDropUp"
SolCallback(5) = "dtBank.SolHit 1,"
SolCallback(6) = "bsSaucer.SolOut"
SolCallback(7) = "VengeanceKickBack"
'8 shaker
'9 left pop
'10 right pop
'11 bottom pop
'13 left slingshot
'14 right slingshot
'15 left flipper
'16 right flipper
SolCallback(22) = "SolLaserMotor"
SolCallback(24) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(51) = "SolLgate"
SolCallback(52) = "SolRgate"
SolCallback(53) = "SolVengShake"
SolCallback(54) = "Solkickback"
'SolCallback(55) = "ScoopRot" 'I didn't add this to my table, the vukk always kick the ball upwards.
SolCallback(56) = "SolLatch"

'Flashers
SolModCallback(17) = "Flasher17"
SolModCallback(18) = "Flasher18"
SolModCallback(19) = "Flasher19"
SolModCallback(20) = "Flasher20"
SolModCallback(21) = "Flasher21"
SolModCallback(23) = "Flasher23"
SolModCallback(25) = "Flasher25"
SolModCallback(26) = "Flasher26"
SolModCallback(27) = "Flasher27"
SolModCallback(28) = "Flasher28"
SolModCallback(29) = "Flasher29"
SolModCallback(30) = "Flasher30"
SolModCallback(31) = "Flasher31"
SolModCallback(32) = "Flasher32"
SolModCallback(59) = "Flasher59"
'backdrop flashers
SolModCallback(60) = "Flasher60"
SolModCallback(61) = "Flasher61"
SolModCallback(62) = "Flasher62"
SolModCallback(63) = "Flasher63"
SolModCallback(64) = "Flasher64"

Sub Flasher17(m):m = m /255:f17.State = m:End Sub
Sub Flasher18(m):m = m /255:f18.State = m:End Sub
Sub Flasher19(m):m = m /255:f19l.State = m:End Sub
Sub Flasher20(m):m = m /255:f20l.State = m:End Sub
Sub Flasher21(m):m = m /255:f21l.State = m:End Sub
Sub Flasher23(m):m = m /255:f23l.State = m:End Sub
Sub Flasher25(m):m = m /255:f25.State = m:f25a.State = m:End Sub
Sub Flasher26(m):m = m /255:f26l.State = m:End Sub
Sub Flasher27(m):m = m /255:f27l.State = m:End Sub
Sub Flasher28(m):m = m /255:f28l.State = m:End Sub
Sub Flasher29(m):m = m /255:f29.State = m:End Sub
Sub Flasher30(m):m = m /255:f30l.State = m:End Sub
Sub Flasher31(m):m = m /255:f31l.State = m:End Sub
Sub Flasher32(m):m = m /255:f32.State = m:End Sub
Sub Flasher59(m):m = m /255:f41.State = m:End Sub
Sub Flasher60(m):m = m /255:f42.State = m:f42a.State = m:End Sub
Sub Flasher61(m):m = m /255:f43.State = m:f43a.State = m:End Sub
Sub Flasher62(m):m = m /255:f44.State = m:End Sub
Sub Flasher63(m):m = m /255:f45.State = m:End Sub
Sub Flasher64(m):m = m /255:f46.State = m:End Sub


' Solenoid subs

Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 22
    End If
End Sub

Sub SolAutofire(Enabled)
    If Enabled Then
        PlaySoundAt "fx_kicker", plunger
        PlungerIM.AutoFire
    End If
End Sub

Sub SolLgate(Enabled)
    LeftGate.Open = Enabled
End Sub

Sub SolRgate(Enabled)
    RightGate.Open = Enabled
End Sub

Sub SolLaserMotor(Enabled)
    LaserTimer.Enabled = Enabled
    f22.Visible = Enabled
End Sub

Sub Solkickback(Enabled)
    If Enabled Then
        PlaySoundAt "fx_plunger", kickback
        Kickback.Fire
    Else
        Kickback.Pullback
    End If
End Sub

Sub VengeanceKickBack(enabled)
    VengKicker.Enabled = enabled
End Sub

Sub VengKicker_Hit():vpmtimer.addtimer 300, "VengKicker.kick 176, 35:PlaySoundAt ""fx_kicker"", VengKicker'":End Sub

'*********************
' Vengeance animation
'*********************

Dim cBall
VengInit

Sub VengInit
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub

Sub SolVengShake(Enabled)
    If Enabled Then
        VengShake
    End If
End Sub

Sub VengShake
    cball.velx = -2 + 2 * RND(1)
    cball.vely = -10 + 2 * RND(1)
End Sub

Sub VengUpdate
    Dim a, b
    a = ckicker.y - cball.y
    b = cball.x - ckicker.x
    Vengeance.rotx = - a/4.5
    Vengeance.roty = b
	l56.X = 481 + b *2
	l57.X = 344 + b *2
	l57a.X = 615 + b *2
	l57.Height = 235 + b*1.7
	l57a.Height = 235 - b*1.7
If DesktopMode Then
	l56.Y = 614 +a *1.8
	l57.Y = 490 +a *1.2
	l57a.Y = 483 +a *1.2
Else
	l56.Y = 614 +a /1.3
	l57.Y = 490 +a /1.6
	l57a.Y = 483 +a /1.6
End If
End Sub

Sub LatchSw_Hit
    Controller.Switch(53) = 1
End Sub

Sub LatchSw_UnHit
    Controller.Switch(53) = 0
End Sub

Sub SolLatch(Enabled)
    If Enabled then
        vMagnet2.MagnetON = NOT vMagnet2.MagnetON

    End If
End Sub

'*******************
' Flipper Subs v3.0
'*******************

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
SolCallback(12) = "SolRFlipper1"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), LeftFlipper
        LeftFlipper.RotateToStart
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

Sub SolRFlipper1(Enabled)
    If Enabled Then
        RightFlipper1.RotateToEnd
    Else
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
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
    If LF_CurAngle >= LF_StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If
 
'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
	If LeftFlipperOn = 1 Then
		If LF_CurAngle = LF_EndAngle then
			LeftFlipper.EOSTorque = FullStrokeEOS_Torque
			LLiveCatchTimer = LLiveCatchTimer + 1
			If LLiveCatchTimer < LiveCatchSensivity Then
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
    If RF_CurAngle <= RF_StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If
 
'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
 	If RightFlipperOn = 1 Then
		If RF_CurAngle = RF_EndAngle Then
			RightFlipper.EOSTorque = FullStrokeEOS_Torque
			RLiveCatchTimer = RLiveCatchTimer + 1
			If RLiveCatchTimer < LiveCatchSensivity Then
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

'*****************
'   Gi Lights
'*****************

Sub GIUpdate(no, Value)
    Select Case no
        Case 0
            For each x in aGiLights
                x.State = ABS(Value)
            Next
            For each x in aGiFlashers
                x.Visible = Value
            Next
    End Select
End Sub

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, 1
    Next
End Sub

Sub ChangeGIIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGILights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' 11 colors: red, orange, amber, yellow...
'******************************************

'colors
Const red = 5
Const orange = 4
Const amber = 6
Const yellow = 3
Const darkgreen = 7
Const green = 2
Const blue = 1
Const darkblue = 8
Const purple = 9
Const white = 11
Const teal = 10

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Select Case col
        Case red
            n.color = RGB(18, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case orange
            n.color = RGB(18, 3, 0)
            n.colorfull = RGB(255, 64, 0)
        Case amber
            n.color = RGB(193, 49, 0)
            n.colorfull = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(18, 18, 0)
            n.colorfull = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 8, 0)
            n.colorfull = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 16, 0)
            n.colorfull = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 18, 18)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 8, 8)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(64, 0, 96)
            n.colorfull = RGB(128, 0, 192)
        Case white
            n.color = RGB(192, 192, 192)
            n.colorfull = RGB(255, 255, 255)
        Case teal
            n.color = RGB(1, 64, 62)
            n.colorfull = RGB(2, 128, 126)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

'**************************************************************
'     JP's Flasher Fading for VPX and Vpinmame v3.0
'       (Based on Pacdude's Fading Light System)
' This is a fast fading for the Flashers in vpinmame tables
'  just 4 steps, like in Pacdude's original script.
' Included the new Modulated flashers & Lights for WPC & Stern
'**************************************************************

Dim LampState(600), FadingState(600), FlashLevel(600)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            Dim cIdx, cVal : cIdx = chgLamp(ii, 0) : cVal = chgLamp(ii, 1)
            LampState(cIdx) = cVal       'keep the real state in an array
            FadingState(cIdx) = cVal + 3 'fading step
            If cIdx = 105 Then GiUpdate 0, cVal
			Select Case cIdx					'Check if GI light needs to be changed
				Case 9:	If cVal = 1 then ChangeGi blue
				Case 15:If cVal = 1 then ChangeGi red
				Case 10:If cVal = 1 then ChangeGi yellow
				Case 14:If cVal = 1 then ChangeGi teal
				Case 16:If cVal = 1 then ChangeGi darkblue
				Case 11:If cVal = 1 then ChangeGi green
			End Select
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps()
    SetRGBLamp l1, Lampstate(84), Lampstate(85), Lampstate(86)
    SetRGBLamp l2, Lampstate(81), Lampstate(87), Lampstate(88)
    SetRGBLamp l3, Lampstate(83), Lampstate(82), Lampstate(89)
    SetRGBLamp l4, Lampstate(92), Lampstate(91), Lampstate(90)
    SetRGBLamp l5, Lampstate(93), Lampstate(94), Lampstate(95)
    SetRGBLamp l6, Lampstate(96), Lampstate(97), Lampstate(98)
    SetRGBLamp l7, Lampstate(99), Lampstate(100), Lampstate(101)
    SetRGBLamp l8, Lampstate(102), Lampstate(103), Lampstate(104)
    SetRGBLamp l9, Lampstate(113), Lampstate(114), Lampstate(115)
    SetRGBLamp l10, Lampstate(116), Lampstate(117), Lampstate(118)
    SetRGBLamp l11, Lampstate(119), Lampstate(120), Lampstate(121)
    SetRGBLamp l12, Lampstate(122), Lampstate(123), Lampstate(124)
    SetRGBLamp l13, Lampstate(125), Lampstate(126), Lampstate(127)
    SetRGBLamp l14, Lampstate(128), Lampstate(129), Lampstate(130)
    SetRGBLamp l15, Lampstate(131), Lampstate(133), Lampstate(132)
    SetRGBLamp l16, Lampstate(134), Lampstate(136), Lampstate(135)
    SetRGBLamp l17, Lampstate(146), Lampstate(147), Lampstate(148)
    SetRGBLamp l18, Lampstate(149), Lampstate(150), Lampstate(151)
    SetRGBLamp l19, Lampstate(152), Lampstate(153), Lampstate(154)
    SetRGBLamp l20, Lampstate(155), Lampstate(156), Lampstate(157)
    SetRGBLamp l21, Lampstate(158), Lampstate(159), Lampstate(160)
    SetRGBLamp l22, Lampstate(161), Lampstate(162), Lampstate(163)
    SetRGBLamp l23, Lampstate(164), Lampstate(165), Lampstate(166)
    SetRGBLamp l24, Lampstate(167), Lampstate(168), Lampstate(169)
    SetRGBLamp l25, Lampstate(170), Lampstate(171), Lampstate(172)
    SetRGBLamp l26, Lampstate(179), Lampstate(177), Lampstate(178)
    SetRGBLamp l27, Lampstate(182), Lampstate(180), Lampstate(181)
    SetRGBLamp l28, Lampstate(185), Lampstate(183), Lampstate(184)
    SetRGBLamp l29, Lampstate(186), Lampstate(187), Lampstate(188)
    SetRGBLamp l30, Lampstate(189), Lampstate(190), Lampstate(191)
    SetRGBLamp l31, Lampstate(192), Lampstate(193), Lampstate(194)
    SetRGBLamp l32, Lampstate(195), Lampstate(197), Lampstate(196)
    SetRGBLamp l33, Lampstate(198), Lampstate(200), Lampstate(199)
    SetRGBLamp l34, Lampstate(214), Lampstate(215), Lampstate(216)
    SetRGBLamp l35, Lampstate(211), Lampstate(217), Lampstate(218)
    SetRGBLamp l36, Lampstate(213), Lampstate(212), Lampstate(219)
    SetRGBLamp l37, Lampstate(222), Lampstate(221), Lampstate(220)
    SetRGBLamp l38, Lampstate(223), Lampstate(224), Lampstate(225)
    SetRGBLamp l39, Lampstate(226), Lampstate(227), Lampstate(228)
    SetRGBLamp l40, Lampstate(229), Lampstate(230), Lampstate(231)
    SetRGBLamp l41, Lampstate(232), Lampstate(233), Lampstate(234)
    SetRGBLamp l42, Lampstate(235), Lampstate(236), Lampstate(237)
    SetRGBLamp l43, Lampstate(238), Lampstate(239), Lampstate(240)
    SetRGBLamp l44, Lampstate(241), Lampstate(242), Lampstate(243)
    SetRGBLamp l45, Lampstate(244), Lampstate(245), Lampstate(246)
    SetRGBLamp l46, Lampstate(247), Lampstate(248), Lampstate(249)
    SetRGBLamp l47, Lampstate(250), Lampstate(251), Lampstate(252)
    SetRGBLamp l48, Lampstate(253), Lampstate(254), Lampstate(255)
    SetRGBLamp l49, Lampstate(256), Lampstate(257), Lampstate(258)
    SetRGBLamp l52, Lampstate(278), Lampstate(279), Lampstate(280)
    SetRGBLamp l53, Lampstate(281), Lampstate(283), Lampstate(282)
    SetRGBLamp l54, Lampstate(284), Lampstate(286), Lampstate(285)
    SetRGBLamp l55, Lampstate(287), Lampstate(289), Lampstate(288)
    SetRGBLamp l58, Lampstate(308), Lampstate(309), Lampstate(310)
    SetRGBLamp l59, Lampstate(311), Lampstate(312), Lampstate(313)
    SetRGBLamp l60, Lampstate(314), Lampstate(315), Lampstate(316)
    SetRGBLamp l61, Lampstate(317), Lampstate(319), Lampstate(318)
    SetRGBLamp l62, Lampstate(320), Lampstate(322), Lampstate(321)
    SetRGBLamp l63a, Lampstate(300), Lampstate(302), Lampstate(301)
    SetRGBLamp l63b, Lampstate(300), Lampstate(302), Lampstate(301)
    SetRGBLamp l63c, Lampstate(300), Lampstate(302), Lampstate(301)
    SetRGBLamp l63d, Lampstate(300), Lampstate(302), Lampstate(301)
    SetRGBLamp l63e, Lampstate(300), Lampstate(302), Lampstate(301)
    SetRGBLamp l63f, Lampstate(300), Lampstate(302), Lampstate(301)
    SetRGBLamp l63g, Lampstate(300), Lampstate(302), Lampstate(301)
    SetRGBLamp l64a, Lampstate(303), Lampstate(305), Lampstate(304)
    SetRGBLamp l64b, Lampstate(303), Lampstate(305), Lampstate(304)
    SetRGBLamp l64c, Lampstate(303), Lampstate(305), Lampstate(304)
    SetRGBLamp l64d, Lampstate(303), Lampstate(305), Lampstate(304)
    SetRGBLamp l64e, Lampstate(303), Lampstate(305), Lampstate(304)
    SetRGBLamp l64f, Lampstate(303), Lampstate(305), Lampstate(304)
    SetRGBLamp l64g, Lampstate(303), Lampstate(305), Lampstate(304)

    l50.Opacity = LampState(276) * 10
    EnterpriseL.Opacity = LampState(277) * 5
    l56.Opacity = LampState(290) * 10
    l57.Opacity = LampState(291) * 10
    l57a.Opacity = LampState(291) * 10

    Flash 295, l70
    Flash 292, l71
    Flash 293, l72
    Flash 294, l73
    Flash 299, l74
    Flash 296, l75
    Flash 297, l76
    Flash 298, l77

    Flashm 50, l78a
    Flash 50, l78
    Flashm 51, l79a
    Flash 51, l79
    Flashm 52, l80a
    Flash 52, l80
    Flashm 53, l81a
    Flash 53, l81
    Flashm 54, l82a
    Flash 54, l82
    Flashm 55, l83a
    Flash 55, l83
    Flashm 56, l84a
    Flash 56, l84
    Flashm 57, l85a
    Flash 57, l85
    Flashm 58, l86a
    Flash 58, l86
    Flashm 59, l87a
    Flash 59, l87
    Flashm 60, l88a
    Flash 60, l88
    Flashm 61, l89a
    Flash 61, l89
    Flashm 62, l90a
    Flash 62, l90
    Flashm 63, l91a
    Flash 63, l91
    Flashm 64, l92a
    Flash 64, l92
    Flashm 65, l93a
    Flash 65, l93
    Flashm 66, l94a
    Flash 66, l94
    Flashm 67, l95a
    Flash 67, l95
    Flashm 68, l96a
    Flash 68, l96
    Flashm 69, l97a
    Flash 69, l97
    Flashm 70, l98a
    Flash 70, l98
    Flashm 71, l99a
    Flash 71, l99
    Flashm 72, l100a
    Flash 72, l100


End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 40 ' flasher fading speed
    LampTimer.Enabled = 1
    For x = 0 to 600
        LampState(x) = 0
        FadingState(x) = 3 ' used to track the fading state
        FlashLevel(x) = 0
    Next
End Sub

Sub SetRGBLamp(MyLight, R, G, B)
    If TypeName(MyLight) = "Light" Then
        MyLight.Color = RGB(R / 10, G / 10, B / 10)
        MyLight.ColorFull = RGB(R, G, B)
        MyLight.State = 1
    ElseIf TypeName(MyLight) = "Flasher" Then
        MyLight.Color = RGB(R, G, B)
    End If
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingState(nr) = abs(value) + 3
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingState(nr)
        Case 4:object.state = 1:FadingState(nr) = 0
        Case 3:object.state = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:object.state = 1
        Case 3:object.state = 0
    End Select
End Sub

' Flashers: 4 is on,3,2,1 fade steps. 0 is off

Sub Flash(nr, object)
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1:FadingState(nr) = 0
        Case 3:Object.IntensityScale = 0.66:FadingState(nr) = 2
        Case 2:Object.IntensityScale = 0.33:FadingState(nr) = 1
        Case 1:Object.IntensityScale = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1
        Case 3:Object.IntensityScale = 0.66
        Case 2:Object.IntensityScale = 0.33
        Case 1:Object.IntensityScale = 0
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub Reel(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1:FadingState(nr) = 0
        Case 3:object.SetValue 2:FadingState(nr) = 2
        Case 2:object.SetValue 3:FadingState(nr) = 1
        Case 1:object.SetValue 0:FadingState(nr) = 0
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1
        Case 3:object.SetValue 2
        Case 2:object.SetValue 3
        Case 1:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message:FadingState(nr) = 0
        Case 3:object.Text = "":FadingState(nr) = 0
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message
        Case 3:object.Text = ""
    End Select
End Sub

' Modulated Subs for the WPC tables

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub LampMod(nr, object)          ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
    Object.State = 1             'in case it was off
End Sub

Sub FlashMod(nr, object)         'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Walls and mostly Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingState(nr)
        Case 4:object.image = a:FadingState(nr) = 0 'fading to off...
        Case 3:object.image = b:FadingState(nr) = 2
        Case 2:object.image = c:FadingState(nr) = 1
        Case 1:object.image = d:FadingState(nr) = 0
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingState(nr)
        Case 4:object.image = a
        Case 3:object.image = b
        Case 2:object.image = c
        Case 1:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingState(nr)
        Case 4:object.image = a:FadingState(nr) = 0 'off
        Case 3:object.image = b:FadingState(nr) = 0 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingState(nr)
        Case 4:object.image = a
        Case 3:object.image = b
    End Select
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
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
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
    If tmp > 0 Then
        Dim t2, t4, t8 : t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4 : Pan = Csng(t8*t2)
    Else
        Dim t2n, t4n, t8n : t2n = tmp*tmp : t4n = t2n*t2n : t8n = t4n*t4n : Pan = Csng(-(t8n*t2n))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX * ball.VelX) + (ball.VelY * ball.VelY)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
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

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v4.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 1     'number of locked balls
Const maxvel = 40 'max ball velocity
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
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound(BallRollStr(b))
        aBallShadow(b).Y = 3000
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z -Ballsize/2

        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 5
            End If
            rolling(b) = True
            PlaySound(BallRollStr(b)), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound(BallRollStr(b))
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity * velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'************************************
' Game timer for real time updates
'************************************

Sub RealTime_Timer
    RollingUpdate
    VengUpdate
    LeftflipperTop.Rotz = LeftFlipper.CurrentAngle
    RightflipperTop.Rotz = RightFlipper.CurrentAngle
    RightflipperTop1.Rotz = RightFlipper1.CurrentAngle
End Sub

'*********************
' Asteroids timer
'*********************

Dim RotAngle
RotAngle = 0

Sub AsteroidsTimer_Timer
    RotAngle = (RotAngle + 1)MOD 360
    Asteroid1.RotZ = RotAngle
    Asteroid2.RotZ = RotAngle
End Sub

'****************
' Laser Timer
'****************

Dim LaserAngle
LaserAngle = 0

Sub LaserTimer_Timer
    LaserAngle = (LaserAngle + 2)MOD 360
    f22.RotZ = LaserAngle
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