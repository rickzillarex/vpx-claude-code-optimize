' The Lord of the Rings / IPD No. 4858 / October, 2003 / 4 Players
' VPX8 version 5.5.0 by jpsalas
' DOF commands by arngrim
'
' *********************************************************************
' VPX Performance Optimizations Applied:
'  1. Flipper tricks timer 1ms -> 10ms
'  2. Cache flipper COM properties in FlipperTricks timer
'  3. Eliminated ^10 in Pan/AudioFade with chain multiply
'  4. Eliminated ^2 in Vol, BallVel, OnBallBallCollision
'  5. Inline BallVel/Vol/Pitch in RollingUpdate loop
'  6. Pre-built rolling sound string array
'  7. Early-exit ball loop when no balls
'  8. Cache chgLamp locals in LampTimer_Timer
'  9. TableWidth/TableHeight already cached (no change needed)
' *********************************************************************

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
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

Const UseVPMModSol = False  'it doesn't work very well, as the ball is not ejected to the plunger, but it may work on later vpinmame versions. If so you need to change the 8 lights in the layer "Flasher Lights" to LED type.

LoadVPM "01560000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Dim x, i, j, k 'used in loops
Dim bsTrough, bsL, bsR, bsTR, bsTL, vlLock, mRingMagnet, plungerIM

'************
' Table init.
'************

' choose the ROM
Const cGameName = "lotr" 'USA
'Const cGameName = "lotr_fr" 'France
'Const cGameName = "lotr_gr" 'Germany
'Const cGameName = "lotr_it" 'Italy
'Const cGameName = "lotr_sp" 'Spain

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1 'ensure the sound is on
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "The Lord of the Rings - Stern 2003" & vbNewLine & "VPX table by JPSalas v5.5.0"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description

        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 90, 3
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
    End With

    ' Left vuk
    Set bsL = New cvpmBallStack
    With bsL
        .InitSaucer sw9, 9, 0, 30
        .KickZ = 1.56
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Right vuk
    Set bsR = New cvpmBallStack
    With bsR
        .InitSaucer sw30, 30, 0, 30
        .KickZ = 1.56
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Top Right Saucer
    Set bsTR = New cvpmBallStack
    With bsTR
        .InitSaucer sw46, 46, 270, 3
        .KickForceVar = 2
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_popper", DOFContactors)
    End With

    ' Top Left vuk
    Set bsTL = New cvpmBallStack
    With bsTL
        .InitSw 0, 41, 0, 0, 0, 0, 0, 0
        .InitKick sw41, 270, 6
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
    End With

    ' Visible Lock - implements post ball lock
    Set vlLock = New cvpmVLock
    With vlLock
        .InitVLock Array(sw19, sw18, sw17), Array(sw19k, sw18k, sw17k), Array(19, 18, 17)
        .InitSnd "fx_sensor", "fx_sensor"
        .CreateEvents "vlLock"
    End With

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 60 ' Plunger Power
    Const IMTime = 0.7        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd "fx_kicker", "fx_kicker"
        .CreateEvents "plungerIM"
    End With

    Set mRingMagnet = New cvpmMagnet
    With mRingMagnet
        .InitMagnet sw47a, 60
        .GrabCenter = True
        .solenoid = 6 'Ring Magnet
        .CreateEvents "mRingMagnet"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    RealTime.Enabled = 1

    ' Div Init
    OrbitPin.IsDropped = 1
    SolBalrog 0

    LoadLUT
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = LeftMagnaSave Then bLutActive = True: SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
    If KeyDownHandler(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False: HideLUT
    If KeyUpHandler(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", plunger:Plunger.Fire
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

Sub DiverterFlipper_Animate: Diverter.RotZ = DiverterFlipper.CurrentAngle: End Sub
Sub sw28_Animate: Balrog.RotZ = sw28.CurrentAngle: End Sub
Sub TowerFlipper_Animate: Tower.RotX = TowerFlipper.CurrentAngle: End Sub

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels 
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

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

Dim GiIntensity
GiIntensity = 1   'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

' New LUT postit
Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea="PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
SetLUTLine ""
LUBack.imagea="PostitBL"
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
    vpmTimer.PulseSw 59
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
    vpmTimer.PulseSw 62
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
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw9_Hit:PlaySoundAt "fx_kicker_enter", sw9:bsL.AddBall 0:End Sub
Sub sw30_Hit:PlaySoundAt "fx_kicker_enter", sw30:bsR.AddBall 0:End Sub
Sub sw41a_Hit:PlaySoundAt "fx_hole_enter", sw41a:bsTL.AddBall Me:End Sub
Sub sw41b_Hit:PlaySoundAt "fx_hole_enter", sw41b:bsTL.AddBall Me:End Sub
Sub sw41c_Hit:PlaySoundAt "fx_hole_enter", sw41c:bsTL.AddBall Me:End Sub
Sub sw41d_Hit:PlaySoundAt "fx_hole_enter", sw41d:bsTL.AddBall Me:End Sub
Sub sw46_Hit:PlaySoundAt "fx_kicker_enter", sw46:bsTR.AddBall 0:End Sub

' Rollovers & Ramp Switches
Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_Unhit:Controller.Switch(60) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAt "fx_sensor", sw21:End Sub
Sub sw21_Unhit:Controller.Switch(21) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_Unhit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_Unhit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub
Sub sw45_Unhit:Controller.Switch(45) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub
Sub sw37_Unhit:Controller.Switch(37) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub

Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "fx_sensor", sw39:End Sub
Sub sw39_Unhit:Controller.Switch(39) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_Unhit:Controller.Switch(33) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:End Sub
Sub sw34_Unhit:Controller.Switch(34) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub
Sub sw35_Unhit:Controller.Switch(35) = 0:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub
Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAt "fx_sensor", sw40:End Sub
Sub sw40_Unhit:Controller.Switch(40) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw42:End Sub
Sub sw42_Unhit:Controller.Switch(42) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_Unhit:Controller.Switch(24) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", sw22:End Sub
Sub sw22_Unhit:Controller.Switch(22) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:ActiveBall.VelX = 3:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_Unhit:Controller.Switch(47) = 0:End Sub

' Targets
Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw28_Collide(parm):vpmTimer.PulseSw 28:PlaySound "fx_plasticHit", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall):End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw52_Spin():vpmTimer.PulseSw 52:PlaySoundAt "fx_spinner", sw52:End Sub

' lock post

Sub lock1_Hit:lock.Isdropped = 1:End Sub
Sub lock1_UnHit:lock.Isdropped = 0:End Sub

'*********
'Solenoids
'*********

SolCallBack(1) = "SolRelease"
SolCallBack(2) = "Auto_Plunger"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsTL.SolOut"
SolCallback(5) = "bsR.SolOut"
' SolCallback(6) = "SolRingMagnet"
SolCallback(7) = "SolTower"
SolCallback(8) = "SolDiv"
' 9 left bumper
'10 right bumper
'11 bottom bumper
'12 not used
SolCallback(13) = "SolOrbit"
'15 left flipper
'16 right flipper
'17 left slingshot
'18 right slingshot
SolCallBack(19) = "bsTR.SolOut"
'20 balrog motor relay
SolCallBack(21) = "SolLockRelease"
SolCallBack(22) = "SolBalrog"
SolCallBack(24) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
'28 not used
' Flashers
SolCallback(32) = "SetLamp 132," 'Balrog

If UseVPMModSol Then
SolModCallback(14) = "Flasher14"
SolModCallback(23) = "Flasher23"
SolModCallback(25) = "Flasher25"
SolModCallback(26) = "Flasher26"
SolModCallback(27) = "Flasher27"
SolModCallback(29) = "Flasher29"
SolModCallback(30) = "Flasher30"
SolModCallback(31) = "Flasher31"
Else
SolCallback(14) = "vpmFlasher Array(f14,f14a),"
SolCallback(23) = "vpmFlasher Array(f23,f23a),"
SolCallback(25) = "vpmFlasher Array(f25a002,f25a001,f25c002,f25c001,f25a3,f25c),"
SolCallback(26) = "vpmFlasher f26l,"
SolCallback(27) = "vpmFlasher f27l,"
SolCallback(29) = "vpmFlasher f29c,"
SolCallback(30) = "vpmFlasher f30,"
SolCallback(31) = "vpmFlasher f31l,"
End If

Sub Flasher14(m): m = m /255: f14.State = m: f14a.State = m: End Sub
Sub Flasher23(m): m = m /255: f23.State = m: f23a.State = m: End Sub
Sub Flasher25(m): m = m /255: f25a002.State = m: f25a001.State = m: f25c002.State = m: f25c001.State = m: f25a3.State = m: f25c.State = m: End Sub
Sub Flasher26(m): m = m /255: f26l.State = m: End Sub
Sub Flasher27(m): m = m /255: f27l.State = m: End Sub
Sub Flasher29(m): m = m /255: f29c.State = m: End Sub
Sub Flasher30(m): m = m /255: f30.State = m: End Sub
Sub Flasher31(m): m = m /255: f31l.State = m: End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolLockRelease(enabled)
    vlLock.SolExit enabled
End Sub

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls > 0 Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

Sub SolOrbit(Enabled):OrbitPin.IsDropped = NOT Enabled:End Sub

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

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

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
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

FlipperPower = 5000
FlipperElasticity = 0.8
FullStrokeEOS_Torque = 0.3 	' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2	' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 10
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
    Dim LFCurAngle, RFCurAngle
    LFCurAngle = LeftFlipper.CurrentAngle
    RFCurAngle = RightFlipper.CurrentAngle

'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LFCurAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
	If LeftFlipperOn = 1 Then
		If LFCurAngle = LeftFlipper.EndAngle then
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
    If RFCurAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
 	If RightFlipperOn = 1 Then
		If RFCurAngle = RightFlipper.EndAngle Then
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

'************************
'    Tower animation
'************************

Sub SolTower(Enabled)
    If Enabled Then
        TowerFlipper.RotateToEnd
    Else
        TowerFlipper.RotateToStart
    End If
End Sub

'************************
'   Diverter animation
'************************

Sub SolDiv(Enabled)
    If Enabled Then
        DiverterFlipper.RotateToEnd
    Else
        DiverterFlipper.RotateToStart
    End If
End Sub

'************************
'    Balrog Animation
'************************

Dim BalrogDir
BalrogDir = 0

Sub SolBalrog(Enabled)
    If Enabled Then
        If BalrogDir = 0 Then
            Controller.Switch(32) = 0
            Controller.Switch(31) = 1
            sw28.RotateToStart
            BalrogDir = 1
        Else
            Controller.Switch(31) = 0
            Controller.Switch(32) = 1
            sw28.RotateToEnd
            BalrogDir = 0
        End If
    End If
End Sub

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array 
'**********************************************************

Dim LampState(200), FadingStep(200), FlashLevel(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii, idx, val
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            idx = chgLamp(ii, 0) : val = chgLamp(ii, 1)
            LampState(idx) = val            'keep the real state in an array
            FadingStep(idx) = val
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps
    Lamp 1, l1
    Lamp 2, l2
    Lamp 3, l3
    Lamp 4, l4
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10
    Lamp 11, l11
    Lamp 12, l12
    Lamp 13, l13
    Lamp 14, l14
    Lamp 15, l15
    Lamp 16, l16
    Lamp 17, l17
    Lamp 18, l18
    Lamp 19, l19
    Lamp 20, l20
    Lamp 21, l21
    Lamp 22, l22
    Lampm 23, l23
    FadeObj 23, globe, "globe_on", "globe_a", "globe_b", "globe"

    Lamp 24, l24
    Lamp 25, l25
    Lamp 26, l26
    Lamp 27, l27
    Lamp 28, l28
    Lamp 29, l29
    Lamp 30, l30
    Lamp 31, l31
    Lamp 32, l32
    Lamp 33, l33
    Lamp 34, l34
    Lamp 35, l35
    Lamp 36, l36
    Lamp 37, l37
    Lamp 38, l38
    Lamp 39, l39
    Lamp 40, l40
    Lamp 41, l41
    Lamp 42, l42
    Lamp 43, l43
    Lamp 44, l44
    Lamp 45, l45
    Lamp 46, l46
    Lamp 47, l47
    Lamp 48, l48
    Lamp 49, l49
    Lamp 50, l50
    Lamp 51, l51
    Lamp 52, l52
    Lamp 53, l53
    Lamp 54, l54
    Lamp 55, l55
    Lamp 56, l56
    Lamp 57, l57
    Lamp 58, l58
    Lamp 59, l59
    Lamp 60, l60
    Lamp 61, l61b
    Lamp 62, l62b
    Lamp 63, l63b
    Lamp 64, l64
    Lamp 65, l65
    Lamp 66, l66
    Lamp 67, l67
    Lamp 68, l68
    Lamp 69, l69
    Lamp 70, l70
    Lamp 71, l71
    Lamp 72, l72
    Lamp 73, l73a
    Lamp 74, l74a
    Lamp 75, l75a
    Lamp 76, l76a
    Lamp 77, l77a
    Lamp 78, l78a
    'Lamp 79, l79
    'Lamp 80, l80
    Lamp 81, l81
    Lamp 82, l82
    Lamp 83, l83
    Lamp 84, l84
    Lamp 85, l85
    Lamp 86, l86
    Lamp 87, l87
    Lamp 88, l88
    Lamp 89, l89
    Lamp 90, l90
    Lamp 91, l91
    Lamp 92, l92
    Lamp 93, l93
    Lamp 94, l94
    Lamp 95, l95
    Lamp 96, l96
    Lamp 97, l97
    Lamp 98, l98
    Lamp 99, l99

    'flashers
    FadeObj 132, balrog, "balrog_on", "balrog_a", "balrog_b", "balrog"
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 10
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingStep(x) = 0
        FlashLevel(x) = 0
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingStep(nr) = abs(value)
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.state = 1:FadingStep(nr) = -1
        Case 0:object.state = 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingStep(nr)
        Case 1:object.state = 1
        Case 0:object.state = 0
    End Select
End Sub

' Flashers:  0 starts the fading until it is off

Sub Flash(nr, object)
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1:FadingStep(nr) = -1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
                FadingStep(nr) = -1
            End If
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
            End If
    End Select
End Sub

' Desktop Objects: Reels & texts

' Reels - 4 steps fading
Sub Reel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 2:FadingStep(nr) = 2
        Case 2:object.SetValue 3:FadingStep(nr) = 3
        Case 3:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 2
        Case 2:object.SetValue 3
        Case 3:object.SetValue 0
    End Select
End Sub

' Reels non fading
Sub NfReel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub NfReelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message:FadingStep(nr) = -1
        Case 0:object.Text = "":FadingStep(nr) = -1
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message
        Case 0:object.Text = ""
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

'Walls, flashers, ramps and Primitives used as 4 step fading images
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = 2
        Case 2:object.image = c:FadingStep(nr) = 3
        Case 3:object.image = d:FadingStep(nr) = -1
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
        Case 2:object.image = c
        Case 3:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = -1
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
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
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Dim bvx, bvy
    bvx = ball.VelX : bvy = ball.VelY
    Vol = Csng((bvx*bvx + bvy*bvy) / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp, tmp2
    tmp = ball.x * 2 / TableWidth-1
    If tmp > 0 Then
        tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        Pan = Csng(tmp2)
    Else
        tmp = -tmp : tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        Pan = Csng(-tmp2)
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    Dim bvx, bvy
    bvx = ball.VelX : bvy = ball.VelY
    BallVel = SQR(bvx*bvx + bvy*bvy)
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp, tmp2
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        AudioFade = Csng(tmp2)
    Else
        tmp = -tmp : tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        AudioFade = Csng(-tmp2)
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
Const lob = 0     'number of locked balls
Const maxvel = 47 'max ball velocity
ReDim rolling(tnob)
' Pre-built rolling sound string array
ReDim BallRollStr(tnob)
Dim iii
For iii = 0 To tnob
    BallRollStr(iii) = "fx_ballrolling" & iii
Next
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    Dim bvx, bvy, bvel, bz
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
        aBallShadow(b).Y = 3000
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        bvx = BOT(b).VelX : bvy = BOT(b).VelY : bz = BOT(b).z
        bvel = SQR(bvx*bvx + bvy*bvy)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z - 25

        If bvel > 1 Then
            If bz < 30 Then
                ballpitch = bvel * 20
                ballvol = Csng((bvx*bvx + bvy*bvy) / 2000)
            Else
                ballpitch = bvel * 20 + 50000 'increase the pitch on a ramp
                ballvol = Csng((bvx*bvx + bvy*bvy) / 2000) * 3
            End If
            rolling(b) = True
            PlaySound BallRollStr(b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and bz <55 and bz > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, bvel * 20, 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If bvx AND bvy <> 0 Then
            speedfactorx = ABS(maxvel / bvx)
            speedfactory = ABS(maxvel / bvy)
            If speedfactorx <1 Then
                BOT(b).VelX = bvx * speedfactorx
                BOT(b).VelY = bvy * speedfactorx
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
    PlaySound("fx_collide"), 0, Csng(velocity * velocity / 2000), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
    For each x in aGiFlashers
        x.visible = Enabled
    Next
End Sub