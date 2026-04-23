' ======================================================================
' VPX Performance Optimizations Applied:
'   - Flipper timer interval 1ms -> 10ms
'   - Cache flipper COM properties (CurrentAngle/StartAngle/EndAngle) in FlipperTricks timer
'   - Eliminate ^2 in Vol/BallVel/OnBallBallCollision - use direct multiplication
'   - Eliminate ^10 in Pan/AudioFade - use chain multiply (t2*t4*t8)
'   - Pre-built BallRollStr() array to eliminate per-frame string concatenation
' ======================================================================

'JP's Viper Night Drivin' (Sega 1998)
'based on
'Viper Night Drivin' / IPD No. 4359 / 1998 / 6 Players
'Sega Pinball, Incorporated, of Chicago, Illinois,
'VPX8 table, jpsalas 2026, v1.0.0

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    DesktopDMD.Visible = 1
    VarHidden = 1 'hide the vpinmame dmd
Else
    UseVPMColoredDMD = False
    DesktopDMD.Visible = 0
    VarHidden = 0
End If

Dim UseVPMModSol
UseVPMModSol = True                                  'this table needs vpinmame 3.7

LoadVPM "03060000", "sega.vbs", 3.26

If VPinMAMEDriverVer <3.61 Then UseVPMModSol = False 'in case you have an older vpinmame

'********************
'Standard definitions
'********************

Const cGameName = "viprsega"
Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Dim plungerIM, bsTrough, bsSaucer, bsLVUK, bsRVUK, x

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's Sega Viper Night Drivin'" & vbNewLine & "VPX8 table by JPSalas v1.0.0"
        .Games(cGameName).Settings.Value("sound") = 1
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003, LeftSlingshot, RightSlingshot)

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 56 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 15, 14, 13, 12, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
        .IsTrough = 1
    End With

    ' Saucers

    Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw44, 44, 225, 3
    bsSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    Set bsLVUK = New cvpmBallStack
    bsLVUK.InitSaucer sw45, 45, 205, 12
    bsLVUK.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    Set bsRVUK = New cvpmBallStack
    bsRVUK.InitSaucer sw46, 46, 160, 12
    bsRVUK.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    UVLight 1
    vpmtimer.addtimer 1500, "GiOn '"

    'init posts - walls
    LeftPost.IsDropped = 1
    RightPost.IsDropped = 1
    CenterPost.IsDropped = 1
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = LeftMagnaSave Then Controller.Switch(20) = 1  'UK Only
    If KeyCode = RightMagnaSave Then Controller.Switch(21) = 1 'UK Only
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If KeyCode = PlungerKey Then Controller.Switch(53) = 1
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = LeftMagnaSave Then Controller.Switch(20) = 0  'UK Only
    If KeyCode = RightMagnaSave Then Controller.Switch(21) = 0 'UK Only
    If KeyCode = PlungerKey Then Controller.Switch(53) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SolAutofire"
SolCallback(3) = "bsLVUK.SolOut"
SolCallback(4) = "bsRVUK.SolOut"
SolCallback(5) = "bsSaucer.SolOut"
SolCallback(6) = "SolPostLeft" 'UK Only
'8 European Token Dispenser
'SolCallback(9)="vpmSolSound ""Jet3"","
'SolCallback(10)="vpmSolSound ""Jet3"","
'SolCallback(11)="vpmSolSound ""Jet3"","
'SolCallback(12)="vpmSolSound ""lsling"","
'SolCallback(13)="vpmSolSound ""lsling"","
SolCallback(14) = "SolPost"
SolCallback(17) = "SolRacoonLeft"
SolCallback(18) = "SolRacoonRight"
SolCallback(20) = "SolRampDiv"
SolCallback(21) = "SolOrbitDiv"
SolCallback(22) = "SolPostRight" 'UK only
'SolCallback() = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(23) = "UVLight" 'UV light
'24 Coin Meter

' Flashers
If UseVPMModSol Then
    ' Modulated flashers

    SolModCallback(25) = "FlasherF1" 'F1
    SolModCallback(26) = "FlasherF2" 'F2
    SolModCallback(27) = "FlasherF3" 'F3
    SolModCallback(28) = "FlasherF4" 'F4
    SolModCallback(29) = "FlasherF5" 'F5
    SolModCallback(30) = "FlasherF6" 'F6
    '31 'not used
    SolModCallback(32) = "FlasherF8" 'F8 pop bumpers
    For each x in aFlasers:x.Fader = 0:Next
Else
    'normal flashers
    SolCallback(25) = "vpmFlasher Array(F1,F1a,F1b,F1d),"             'F1 X2
    SolCallback(26) = "vpmFlasher Array(F2,F2a,F2b,F2c),"             'F2 X2
    SolCallback(27) = "vpmFlasher Array(F3,F3a,F3b,F3c),"             'F3 X2
    SolCallback(28) = "vpmFlasher Array(F4,F4a,F4b,F4c,F4d,F4e,F4f)," 'F4 X3
    SolCallback(29) = "vpmFlasher Array(F5,F5a,F5b,F5c,F5d,F5e,F5f)," 'F5 X3
    SolCallback(30) = "vpmFlasher F6,"                                'F6 X2
    SolCallback(32) = "vpmFlasher Array(F8,F8a,F8b),"                 'F8 X3
    For each x in aFlasers:x.Fader = 2:Next
End If

Sub UVLight(Enabled) 'in this table I simply added pink gi lights :)
    If Enabled Then
        For each x in aUVGi:x.State = 0:Next
    Else
        For each x in aUVGi:x.State = 1:Next
    End If
End Sub

Sub FlasherF1(m):m = m / 255:F1.State = m:F1a.State = m:F1b.State = m:F1d.State = m:End Sub
Sub FlasherF2(m):m = m / 255:F2.State = m:F2a.State = m:F2b.State = m:F2c.State = m:End Sub
Sub FlasherF3(m):m = m / 255:F3.State = m:F3a.State = m:F3b.State = m:F3c.State = m:End Sub
Sub FlasherF4(m):m = m / 255:F4.State = m:F4a.State = m:F4b.State = m:F4c.State = m:F4d.State = m:F4e.State = m:F4f.State = m:End Sub
Sub FlasherF5(m):m = m / 255:F5.State = m:F5a.State = m:F5b.State = m:F5c.State = m:F5d.State = m:F5e.State = m:F5f.State = m:End Sub
Sub FlasherF6(m):m = m / 255:F6.State = m:End Sub
Sub FlasherF8(m):m = m / 255:F8.State = m:F8a.State = m:F8b.State = m:End Sub


'Solenoid subs

Sub solAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolPostLeft(Enabled) 'UK Only
    If Enabled Then
        LeftPost.IsDropped = 0
        PlaysoundAt SoundFX("fx_SolenoidOn", DOFContactors), Primitive001
    Else
        LeftPost.IsDropped = 1
        PlaysoundAt SoundFX("fx_SolenoidOff", DOFContactors), Primitive001
    End If
End Sub

Sub SolPostRight(Enabled) 'UK Only
    If Enabled Then
        RightPost.IsDropped = 0
        PlaysoundAt SoundFX("fx_SolenoidOn", DOFContactors), Primitive013
    Else
        RightPost.IsDropped = 1
        PlaySoundAt SoundFX("fx_SolenoidOff", DOFContactors), Primitive013
    End If
End Sub

Sub CenterPostF_Animate:CenterPostP.Z = CenterPostF.CurrentAngle:End Sub

Sub SolPost(Enabled)
    If Enabled Then
        CenterPost.IsDropped = 0
        CenterPostF.RotateToEnd
        PlaysoundAt SoundFX("FX_SolenoidOn", DOFContactors), CenterPostF
    Else
        CenterPost.IsDropped = 1
        CenterPostF.RotateToStart
        PlaysoundAt SoundFX("FX_SolenoidOff", DOFContactors), CenterPostF
    End If
End Sub

Sub SolRacoonLeft(Enabled)
    If Enabled Then
        RacoonLeftF.RotateToEnd
        PlaySoundAt "fx_SolenoidOn2", RacoonLeft
    Else
        RacoonLeftF.RotateToStart
        PlaySoundAt "fx_SolenoidOff2", RacoonLeft
    End If
End Sub

Sub RacoonLeftF_Animate:RacoonLeft.Z = 60 + RacoonLeftF.CurrentAngle:End Sub
Sub RacoonRightF_Animate:RacoonRight.Z = 60 + RacoonRightF.CurrentAngle:End Sub

Sub SolRacoonRight(Enabled)
    If Enabled Then
        RacoonRight.TransZ = 20
        PlaySoundAt "fx_SolenoidOn2", RacoonRight
    Else
        RacoonRight.TransZ = 0
        PlaySoundAt "fx_SolenoidOff2", RacoonRight
    End If
End Sub

Sub SolRampDiv(Enabled)
    If Enabled Then
        TopDiverter.RotateToEnd
        PlaySoundAt SoundFX("fx_SolenoidOn2", DOFContactors), TopDiverter
    Else
        TopDiverter.RotateToStart
        PlaySoundAt SoundFX("fx_SolenoidOff2", DOFContactors), TopDiverter
    End If
End Sub

Sub SolOrbitDiv(Enabled)
    If Enabled Then
        OrbitDiverter.IsDropped = 0
        PlaySoundAt SoundFX("fx_SolenoidOn", DOFContactors), Primitive012
    Else
        OrbitDiverter.IsDropped = 1
        PlaySoundAt SoundFX("fx_SolenoidOff", DOFContactors), Primitive012
    End If
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Dim GiIntensity
GiIntensity = 1   'can be used For the LUT changing to increase the GI lights when the table is darker

Sub ChangeGi(col) 'changes the gi color
    Dim bulb
    For each bulb in aGILights
        SetLightColor bulb, col, -1
    Next
End Sub

Sub ChangeGIIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGILights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

Sub GIUpdate(no, Enabled)
    ' debug.print no
    If Enabled Then
        GiOn
    Else
        GiOff
    End If
End Sub

Sub GiOn
    Dim bulb
    PlaySound "fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    PlaySound "fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'******************************************
' Change light color - simulate color leds
' changes the light color and state
' 11 colors: red, orange, amber, yellow...
'******************************************

'colors
Const amber = 0
Const yellow = 1
Const green = 2
Const darkgreen = 3
Const blue = 4
Const darkblue = 5
Const purple = 6
Const red = 7

Const orange = 8
'Const amber = 9
Const teal = 10
Const white = 11

Sub SetLightColor(n, col, stat) 'stat 0 = off, 1 = on, 2 = blink, -1= no change
    Select Case col
        Case red
            n.color = RGB(255, 0, 0)
            n.colorfull = RGB(255, 0, 0)
        Case orange
            n.color = RGB(255, 64, 0)
            n.colorfull = RGB(255, 64, 0)
        Case amber
            n.color = RGB(255, 153, 0)
            n.colorfull = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(255, 255, 0)
            n.colorfull = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 64, 0)
            n.colorfull = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 128, 0)
            n.colorfull = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 255, 255)
            n.colorfull = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 64, 64)
            n.colorfull = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 192)
            n.colorfull = RGB(128, 0, 192)
        Case teal
            n.color = RGB(2, 128, 126)
            n.colorfull = RGB(2, 128, 126)
        Case white
            n.color = RGB(255, 252, 224)
            n.colorfull = RGB(255, 252, 224)
    End Select
    If stat <> -1 Then
        n.State = 0
        n.State = stat
    End If
End Sub

Sub SetFlashColor(n, col, stat) 'Flashers are linked to lights in VPX8
    Select Case col
        Case red
            n.color = RGB(255, 0, 0)
        Case orange
            n.color = RGB(255, 64, 0)
        Case amber
            n.color = RGB(255, 153, 0)
        Case yellow
            n.color = RGB(255, 255, 0)
        Case darkgreen
            n.color = RGB(0, 64, 0)
        Case green
            n.color = RGB(0, 128, 0)
        Case blue
            n.color = RGB(0, 255, 255)
        Case darkblue
            n.color = RGB(0, 64, 64)
        Case purple
            n.color = RGB(128, 0, 192)
        Case white
            n.color = RGB(255, 252, 224)
        Case teal
            n.color = RGB(2, 128, 126)
    End Select
End Sub


'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 59
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 62
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Scoring rubbers

Sub sw22_Slingshot:vpmTimer.pulseSw 22:End Sub
Sub sw23_Slingshot:vpmTimer.pulseSw 23:End Sub

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 50:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 49:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 51:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw44_Hit:PlaysoundAt "fx_kicker_enter", sw44:bsSaucer.AddBall 0:End Sub
Sub sw45_Hit:PlaysoundAt "fx_kicker_enter", sw45:bsLVUK.AddBall 0:End Sub
Sub sw46_Hit:PlaysoundAt "fx_kicker_enter", sw46:bsRVUK.AddBall 0:End Sub

' Ramp Gates

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_gate", sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySoundAt "fx_gate", sw19:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_gate", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_gate", sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_gate", sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_gate", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

'Optical Trigger

Sub sw17_hit 'Top Jump Metal Ramp
    Controller.Switch(17) = 1
    If activeball.Vely <-15 Then
        activeball.Vely = -15
    End If
End Sub
Sub sw17_unhit:Controller.Switch(17) = 0:End Sub

' Rollovers

Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw41:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw42:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub


'Targets
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers top animations

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
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

'*********************************
' Diverse Collection Hit Sounds
'*********************************

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
Sub aTargets_Hit(idx):PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall):End Sub
Sub aRollovers_Hit(idx):PlaySoundAt "fx_sensor", aRollovers(idx):End Sub

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

'***********************************************
'   JP's VP10 Rolling Sounds
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
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
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound(BallRollStr(b))
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 2
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

' Extra triggers

Sub Trigger001_Hit
    If activeball.Vely> 8 Then
        activeball.Vely = 8
        PlaySoundAt "fx_rampdrop", Trigger001
    End If
End Sub

Sub Trigger002_Hit
    activeball.Vely = 2
    activeball.Velx = 0
    PlaySoundAt "fx_rampdrop", Trigger002
End Sub

Sub Trigger003_Hit
    If activeball.Vely <-15 Then
        activeball.Vely = -15
    End If
End Sub

'*****************************
'    Change RAMP colors
'*****************************

Dim RampColor

Sub UpdateRampColor
    Dim x
    Select Case RampColor
        Case 0:x = RGB(0, 64, 255) 'blue
        Case 1:x = RGB(96, 96, 96) 'White
        Case 2:x = RGB(0, 128, 32) 'Green
        Case 3:x = RGB(128, 0, 0)  'Red
    End Select
    MaterialColor "Plastic Transp Ramps Red", x
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

    ' Color Ramps
    RampColor = Table1.Option("Color Ramps", 0, 3, 1, 3, 0, Array("Blue", "White", "Green", "Red") )
    UpdateRampColor

    'Left CAR color
    x = Table1.Option("Left CAR Color", 0, 4, 1, 0, 0, Array("Blue", "White", "Yellow", "Red", "Red & White") )
    Select Case x
        Case 0:LeftCar.Image = "ViperBody_Blue"
        Case 1:LeftCar.Image = "ViperBody_White"
        Case 2:LeftCar.Image = "ViperBody_Yellow"
        Case 3:LeftCar.Image = "ViperBody_Red"
        Case 4:LeftCar.Image = "ViperBody_RedWhite"
    End Select

    'Right CAR color
    y = Table1.Option("Right CAR Color", 0, 4, 1, 2, 0, Array("Blue", "White", "Yellow", "Red", "Red & White") )
    Select Case y
        Case 0:RightCar.Image = "ViperBody_Blue"
        Case 1:RightCar.Image = "ViperBody_White"
        Case 2:RightCar.Image = "ViperBody_Yellow"
        Case 3:RightCar.Image = "ViperBody_Red"
        Case 4:RightCar.Image = "ViperBody_RedWhite"
    End Select
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