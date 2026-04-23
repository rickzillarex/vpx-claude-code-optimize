' JP's IronMan 2 Armored Adventures v4.3
' Based on The Stern table from 2007

Option Explicit

Randomize

Const BallSize = 50
Const BallMass = 1

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "sam.vbs", 3.10

'********************
'Standard definitions
'********************

Const cGameName = "im_186ve" 'vault edition rom
'Const cGameName = "im_186" 'iron man 2 rom

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "fx_coin"

'Variables
Dim bsTrough, PlungerIM, Mag1, Mag2, x

'************
' Table init.
'************

Sub Table1_Init
    vpminit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Iron-Man 2 (Stern 2007)"
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

    On Error Goto 0

    Controller.Switch(53) = 1 'sandman down

    'Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd "ballrelease", "Solenoid"
    bsTrough.Balls = 4

    ' Magnets
    Set mag1 = New cvpmMagnet
    With mag1
        .InitMagnet Magnet1, 30
        .GrabCenter = False
        .solenoid = 3
        .CreateEvents "mag1"
    End With

    Set mag2 = New cvpmMagnet
    With mag2
        .InitMagnet Magnet2, 30
        .GrabCenter = False
        .solenoid = 4
        .CreateEvents "mag2"
    End With

    'Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    'Impulse Plunger
    Const IMPowerSetting = 55 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Switch 23
        .Random 1.5
        .InitExitSnd "fx_popper", "fx_popper"
        .CreateEvents "plungerIM"
    End With

    ' walls
    ClaneUpPost.Isdropped = true
    UpPost.Isdropped = true
    mongerframe.isdropped = 1
    sw5.isdropped = 1
    sw4.isdropped = 1
    sw6.isdropped = 1

    'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    'Fast Flips
    On Error Resume Next
    InitVpmFFlipsSAM
    If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5"
    On Error Goto 0

    RealTime.Enabled = 1

    'Load LUT
    LoadLUT
End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:MongerShake2
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:MongerShake2
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "fx_nudge", 0, 1, 0, 0.25:MongerShake2
    If keycode = LeftMagnaSave Then bLutActive = True: SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
    If vpmKeyDown(Keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False: HideLUT
    If vpmKeyUp(Keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

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

'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
' SolCallback(3) Monger Magnet
' SolCallback(4) Whiplash Magnet
SolCallback(5) = "WMKick"
SolCallback(6) = "OrbitPost"
'SolCallback(7) = "Gate3.open ="
'SolCallback(8) = "Gate6.open ="

SolCallback(12) = "ClanePost"

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"

SolCallback(19) = "SolMonger"

'Flashers
SolCallback(20) = "SetLamp 120,"
SolCallback(21) = "SetLamp 121,"
SolCallback(22) = "SetLamp 122,"
SolCallback(23) = "SetLamp 123,"
SolCallback(25) = "SetLamp 125,"
SolCallback(26) = "SetLamp 126,"
SolCallback(27) = "SetLamp 127,"
SolCallback(28) = "SetLamp 128,"
SolCallback(29) = "SetLamp 129,"
SolCallback(30) = "SetLamp 130,"
SolCallback(31) = "SetLamp 131,"
SolCallback(32) = "SetLamp 132,"

'************************
' Shake Whiplash when hit
'************************

Dim WhiplashPos

Sub ShakeWhiplash
    WhiplashPos = 12
    WhiplashShakeTimer.Enabled = 1
End Sub

Sub WhiplashShakeTimer_Timer
    Whiplash.TransX = WhiplashPos
    If WhiplashPos = 0 Then WhiplashShakeTimer.Enabled = 0:Exit Sub
    If WhiplashPos < 0 Then
        WhiplashPos = ABS(WhiplashPos)- 1
    Else
        WhiplashPos = - WhiplashPos + 1
    End If
End Sub

'****************
' Shake IronMan
'****************

Dim IronManPos

Sub ShakeIronMan
    IronManPos = 12
    IronManShakeTimer.Enabled = 1
End Sub

Sub IronManShakeTimer_Timer
    IronMan.TransX = IronManPos
    If IronManPos = 0 Then IronManShakeTimer.Enabled = 0:Exit Sub
    If IronManPos < 0 Then
        IronManPos = ABS(IronManPos)- 1
    Else
        IronManPos = - IronManPos + 1
    End If
End Sub

Sub iManShakeTrigger_Hit:ShakeIronMan:End Sub

'******************
'Solenoid Functions
'******************

Sub solTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 22
    End If
End Sub

Sub solAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub ClanePost(Enabled)
    If Enabled Then
        ClaneUpPost.Isdropped = false
    Else
        ClaneUpPost.Isdropped = true
    End If
End Sub

Sub OrbitPost(Enabled)
    If Enabled Then
        UpPost.Isdropped = false
    Else
        UpPost.Isdropped = true
    End If
End Sub

Sub WMKick(enabled)
    If enabled Then
        PlaySoundAt "ballhit", sw10
        sw10.Kick 180, 30
        controller.switch(10) = false
    End If
End Sub

'Drains
Sub drain_Hit():PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 26
    LeftSlingShot.TimerEnabled = 1
'ShakeLeftSpider
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
    vpmTimer.PulseSw 27
    RightSlingShot.TimerEnabled = 1
'ShakeRightSpider
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 30:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

'Lower Lanes
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

'Upper Lanes
Sub sw7_Hit:Controller.Switch(7) = 1:PlaySoundAt "fx_sensor", sw7:End Sub
Sub sw7_UnHit:Controller.Switch(7) = 0:End Sub
Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "fx_sensor", sw39:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub

'Ramp switches
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw49_Hit:Controller.Switch(49) = 1:PlaySoundAt "fx_sensor", sw49:End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub

'Spinners
Sub sw11_Spin:vpmTimer.PulseSw 11:PlaySoundAt "fx_spinner", sw11:End Sub
Sub sw13_Spin:vpmTimer.PulseSw 13:PlaySoundAt "fx_spinner", sw13:End Sub
Sub sw14_Spin:vpmTimer.PulseSw 14:PlaySoundAt "fx_spinner", sw14:End Sub

'Targets
Sub sw33_Hit:vpmTimer.PulseSw 33:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):ShakeWhiplash:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):ShakeWhiplash:End Sub
Sub sw50_Hit:vpmTimer.PulseSw 50:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'kickers
Sub sw10_Hit():controller.switch(10) = true:PlaySoundAt "fx_kicker_enter", sw10:End Sub

'Monger
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtBall "fx_target":MongerShake:End Sub
Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySoundAtBall "fx_target":MongerShake:End Sub
Sub sw6_Hit:vpmTimer.PulseSw 6:PlaySoundAtBall "fx_target":MongerShake:End Sub

'*******************
' Flipper Subs Rev3
'*******************

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
    ' Cache flipper COM properties into locals (eliminates ~20 COM reads per tick)
    Dim Lca : Lca = LeftFlipper.CurrentAngle
    Dim Lsa : Lsa = LeftFlipper.StartAngle
    Dim Rca : Rca = RightFlipper.CurrentAngle
    Dim Rsa : Rsa = RightFlipper.StartAngle

'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If Lca >= Lsa - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
	If LeftFlipperOn = 1 Then
		If Lca = LeftFlipper.EndAngle then
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
    If Rca <= Rsa + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
 	If RightFlipperOn = 1 Then
		If Rca = RightFlipper.EndAngle Then
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

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array
'**********************************************************

ReDim LampState(200)
ReDim FadingStep(200)
ReDim FlashLevel(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii, cIdx, cVal
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            cIdx = chgLamp(ii, 0) : cVal = chgLamp(ii, 1)
            LampState(cIdx) = cVal
            FadingStep(cIdx) = cVal
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps()
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
    Lamp 23, l23
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
    Lampm 44, l44a
    Lamp 44, l44
    Lampm 45, l45a
    Lamp 45, l45
    Lampm 46, l46a
    Lamp 46, l46
    Lampm 47, l47a
    Lamp 47, l47
    Lampm 48, l48a
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
    Lampm 60, Bumper1L
    Flash 60, Flasher1
    Lampm 61, Bumper2L
    Flash 61, Flasher2
    Lampm 62, Bumper3L
    Flash 62, Flasher3
    Lamp 63, l63

    'Flashers
    Lamp 120, f20
    Flashm 121, f21a
    Flashm 121, f21big
    Flash 121, f21
    Lamp 122, f22
    Flashm 125, f25b
    Flash 125, f25
    Flashm 126, f26a
    Flashm 126, F26big
    Flash 126, f26
    Lampm 127, f27a
    Lampm 127, f27b
    Lamp 127, f27c
    Flash 123, f23
    If MongerPos = 198 Then
        Flash 128, f28
    Else
        f28.IntensityScale = 0
    End If
    Lampm 129, f29a
    Lamp 129, f29b
    Flash 130, f30
    Flashm 131, f31a
    Flashm 131, f31b
    Flashm 131, f31c
    Flashm 131, f31d
    Flashm 131, f31e
    Flashm 131, F31big
    Flash 131, f31
    Flashm 132, f32a
    Flashm 132, f32b
    Flashm 132, F32big
    Flash 132, f32
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
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

' Pre-computed inverse multipliers for AudioFade/Pan (eliminates division per call)
Dim InvTWHalf, InvTHHalf
InvTWHalf = 2 / TableWidth
InvTHHalf = 2 / TableHeight

' Pre-built rolling sound strings (eliminates string concat per ball per frame)
ReDim BallRollStr(19)
Dim brsI : For brsI = 0 To 19 : BallRollStr(brsI) = "fx_ballrolling" & brsI : Next

' Flipper bat RotZ tracking for guarded writes
Dim lastLFTopAngle : lastLFTopAngle = -9999
Dim lastRFTopAngle : lastRFTopAngle = -9999

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Dim bv : bv = BallVel(ball)
    Vol = Csng(bv * bv / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * InvTWHalf - 1
    If tmp > 0 Then
        Dim t2p, t4p, t8p : t2p = tmp*tmp : t4p = t2p*t2p : t8p = t4p*t4p
        Pan = Csng(t8p * t2p)
    Else
        Dim nt : nt = -tmp
        Dim t2n, t4n, t8n : t2n = nt*nt : t4n = t2n*t2n : t8n = t4n*t4n
        Pan = Csng(-(t8n * t2n))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    Dim vx, vy : vx = ball.VelX : vy = ball.VelY
    BallVel = SQR(vx * vx + vy * vy)
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * InvTHHalf - 1
    If tmp > 0 Then
        Dim t2p, t4p, t8p : t2p = tmp*tmp : t4p = t2p*t2p : t8p = t4p*t4p
        AudioFade = Csng(t8p * t2p)
    Else
        Dim nt : nt = -tmp
        Dim t2n, t4n, t8n : t2n = nt*nt : t4n = t2n*t2n : t8n = t4n*t4n
        AudioFade = Csng(-(t8n * t2n))
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
Const maxvel = 42 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls
    Dim ubBot : ubBot = UBound(BOT)

    ' stop the sound of deleted balls
    For b = ubBot + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
        aBallShadow(b).Y = 3000
    Next

    ' exit the sub if no balls on the table
    If ubBot = lob - 1 Then Exit Sub

    ' play the rolling sound for each ball and draw the shadow
    Dim bx, by, bz, bvx, bvy, bvz, bv
    For b = lob to ubBot
        ' Cache COM properties once per ball
        bx = BOT(b).X : by = BOT(b).Y : bz = BOT(b).Z
        bvx = BOT(b).VelX : bvy = BOT(b).VelY

        aBallShadow(b).X = bx
        aBallShadow(b).Y = by
        aBallShadow(b).Height = bz - 25 ' BallSize/2 = 25

        ' Inline BallVel: compute once, reuse for Vol and Pitch
        bv = SQR(bvx * bvx + bvy * bvy)

        If bv > 1 Then
            If bz < 30 Then
                ballpitch = bv * 20
                ballvol = Csng(bv * bv / 2000)
            Else
                ballpitch = bv * 20 + 50000
                ballvol = Csng(bv * bv / 2000) * 10
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
        bvz = BOT(b).VelZ
        If bvz < -1 Then
            If bz < 55 Then
                If bz > 27 Then
                    PlaySound "fx_balldrop", 0, ABS(bvz) / 17, Pan(BOT(b)), 0, bv * 20, 1, 0, AudioFade(BOT(b))
                End If
            End If
        End If

        ' jps ball speed control
        If bvx AND bvy <> 0 Then
            speedfactorx = ABS(maxvel / bvx)
            speedfactory = ABS(maxvel / bvy)
            If speedfactorx < 1 Then
                BOT(b).VelX = bvx * speedfactorx
                BOT(b).VelY = bvy * speedfactorx
            End If
            If speedfactory < 1 Then
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

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
    Dim aLF : aLF = LeftFlipper.CurrentAngle
    If aLF <> lastLFTopAngle Then lastLFTopAngle = aLF : LeftflipperTop.Rotz = aLF
    Dim aRF : aRF = RightFlipper.CurrentAngle
    If aRF <> lastRFTopAngle Then lastRFTopAngle = aRF : RightflipperTop.Rotz = aRF
End Sub

'*************************
' GI - needs vpinmame 3
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    If Enabled Then
        PlaySound "fx_GiOn"
    Else
        PlaySound "fx_Gioff"
    End If
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
End Sub

'******************
' Monger Animation
'******************

Dim MongerPos, MongerDir

' start with monger dropped.
MongerDir = -2:MongerPos = 0
Controller.Switch(1) = 0
Controller.Switch(3) = 1

Sub Solmonger(Enabled)
    If Enabled Then
        If MongerDir = 2 Then
            DropMonger
        Else
            RiseMonger
        End If
    End If
End Sub

Sub RiseMonger()
    PlaySound "fx_motor"
    MongerDir = 2
    Controller.Switch(1) = 0
    MongerTimer.Enabled = 1
End Sub

Sub DropMonger()
    PlaySound "fx_motor"
    MongerDir = -2
    Controller.Switch(3) = 0
    MongerTimer.Enabled = 1
End Sub

Sub MongerTimer_Timer
    MongerPos = MongerPos + MongerDir
    If MongerPos > 198 Then
        MongerPos = 198
        Me.Enabled = 0
        Controller.Switch(3) = 1
    Else
        If MongerPos < 0 Then
            MongerPos = 0
            Me.Enabled = 0
            Controller.Switch(1) = 1
        End If
    End If
    UpdateMonger
End Sub

Sub UpdateMonger
    MongerFrameP.TransZ = MongerPos
    MongerCage.TransZ = MongerPos
    Monger.TransZ = MongerPos
    If MongerPos > 140 Then
        sw4.IsDropped = 0
        sw5.IsDropped = 0
        sw6.IsDropped = 0
        mongerframe.IsDropped = 0
    End If
    If MongerPos < 20 Then
        sw4.IsDropped = 1
        sw5.IsDropped = 1
        sw6.IsDropped = 1
        mongerframe.IsDropped = 1
    End If
End Sub

'********************************************
' Monger Shake animations when hit or nudging
'********************************************

'captive ball for hit animations

Dim ccBall
Const cMod = .65 'percentage of hit power transfered to the 3 Bank of targets

InitCaptiveBall

Sub InitCaptiveBall
    Set ccBall = hball.CreateSizedBallWithMass(25, 1.6)
    hball.Kick 0, 0
End Sub

Sub MongerShake
    ccball.velx = activeball.velx * cMod
    ccball.vely = activeball.vely * cMod
    CaptiveTimer.enabled = True
    CaptiveTimer2.enabled = True
End Sub

Sub MongerShake2 'when nudging
    CaptiveTimer.enabled = True
    CaptiveTimer2.enabled = True
End Sub

Sub CaptiveTimer_Timer           'start animation
    Dim x, y
    x = (hball.x - ccball.x) / 4 'reduce the X axis movement
    y = (hball.y - ccball.y) / 2
    MongerFrameP.transy = x
    MongerFrameP.transx = - y
    MongerCage.transy = x
    MongerCage.transx = - y
    Monger.transy = x
    Monger.transx = - y
End Sub

Sub CaptiveTimer2_Timer 'stop animation
    MongerFrameP.transy = 0
    MongerFrameP.transx = 0
    MongerCage.transy = 0
    MongerCage.transx = 0
    Monger.transy = 0
    Monger.transx = 0
    CaptiveTimer.enabled = False
    CaptiveTimer2.enabled = False
End Sub