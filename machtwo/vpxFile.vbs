' *********************************************************************
' VPX Performance Optimizations Applied:
'  1. Flipper tricks timer 1ms -> 10ms
'  2. Cache flipper COM properties in FlipperTricks timer
'  3. Eliminated ^10 in Pan/AudioFade with chain multiply
'  4. Eliminated ^2 in Vol, BallVel, OnBallBallCollision
'  5. Inline BallVel/Vol/Pitch in RollingTimer loop
'  6. Pre-built rolling sound string array
'  7. Early-exit ball loop when no balls
'  8. Cache chgLamp locals in LampTimer_Timer
' *********************************************************************
' Spinball's Mach 2.0 Two / IPD No. 4617 / 1995

' Thanks to :
' destruk / TAB and tinyrodent for their old table. most vpinmame scripting are taken from their table.
' Dip menu by Inkochnito
' Wildman : directb2s
' JPSalas : your time for testing, various script and elements of spinball tables.
' Kiwi : your time for testing and advice.


'v1.1
'fix slingshot force
'fix ambient color



Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "spinball.vbs", 3.26

Const cGameName = "mach2"
'Const cGameName = "mach2a"

Dim VarHidden, UseVPMColoredDMD

If Table1.ShowDT = True Then
    UseVPMColoredDMD = True
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

If B2SOn = True Then VarHidden = 1

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SCoin = "fx_Coin"

Dim bsTrough,bsTunel,bsPicabolas,bsVUK,cbBall,plungerIM,x

'************************************************************************************************************************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Mach 2.0 Two - Spinball 1996" & vbNewLine & "VPX8 table by Fredobiwan v1.0"
        .Games(cGameName).Settings.Value("sound") = 1
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
'        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(sw66, sw76, RightSlingshot, LeftSlingshot)

	' Trough handler - standard, but no entry switch
	Set bsTrough = New cvpmBallStack
	bsTrough.InitSw 0,72,71,70,0,0,0,0
	bsTrough.InitKick BallRelease,90,8
	bsTrough.Balls = 3
	bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

	' Tunel Subterraneo - under playfield
	Set bsTunel = New cvpmBallStack
	bsTunel.InitSw 0,67,0,0,0,0,0,0
	bsTunel.InitKick sw67,282,35
	bsTunel.KickAngleVar = 10
	bsTunel.KickForceVar = 2
	bsTunel.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

	' Picabolas - top kicker
	Set bsPicabolas = New cvpmBallStack
	bsPicabolas.InitSaucer sw60,60,180,10
	bsPicabolas.KickAngleVar = 3
	bsPicabolas.KickForceVar = 3
	bsPicabolas.KickZ = 0.3
	bsPicabolas.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

	'Elevador - Bonus X VUK
	Set bsVUK = New cvpmBallStack
	bsVUK.InitSaucer sw50, 50, 110, 26
	bsVUK.KickZ = 1.5
	bsVUK.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    'Captive ball
	Set cbBall = New cvpmCaptiveBall
	cbBall.InitCaptive CapTrigger, CapWall, CapKicker, 360
	cbBall.NailedBalls = 0
	cbBall.ForceTrans = .8
	cbBall.MinForce = 2
	cbBall.CreateEvents "cbBall"
	cbBall.Start

    'Impulse Plunger
	Const IMPowerSetting = 55 ' Plunger Power
	Const IMTime = 0.1        ' Time in seconds for Full Plunge
	Set plungerIM = New cvpmImpulseP
	With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Switch 73
        .Random .2
        .InitExitSnd SoundFX("fx_solenoidon",DOFContactors), "fx_solenoidoff"
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"

	' Init object
	Post.IsDropped = 0
	BallSavePrim.TransY = 0
	LockRamp.TransZ = -82

	If Controller.Dip(0) = 2 Then
	'nvram was cleared - force dip initialization
	Controller.Dip(0) = 9
	Controller.Dip(1) = 16
	Controller.Dip(2) = 9
	myShowDips
	vpmTimer.AddTimer 1000, "vpmKeyUp keyReset'"
	End If

End Sub

'************************************************************************************************************************

'SolCallback(1) = "Sol1"'Credit Button Light
'SolCallback(2) = "Sol2"'Coin Lockout

'************************************************************************************************************************

' knocker

SolCallBack(3)= "vpmSolSound SoundFX(""Fx_knocker"",DOFKnocker),"

'************************************************************************************************************************

' trough

SolCallback(4) = "bsTrough.SolOut"

Sub Drain2_Hit:PlaySoundAt "fx_drain", Drain2:bsTrough.AddBall Me:End Sub

Sub Drain_Hit:PlaySoundAt "fx_rampdrop", Drain:End Sub

'************************************************************************************************************************

' game on

SolCallback(5)= "GameStart"

Sub GameStart(Enabled):vpmNudge.SolGameOn Enabled:End Sub

'************************************************************************************************************************

'SolCallBack(6) = left bumper
'SolCallBack(7) = right bumper

'************************************************************************************************************************

' plane

SolCallBack(8) = "SolPost"

Sub solPost(Enabled)
	If Enabled Then
		Post.IsDropped = 1
	Else
		Post.IsDropped = 0
	End If 
End Sub

'************************************************************************************************************************

' tunel

SolCallBack(9) = "bsTunel.SolOut"

' no switch under playfield

Sub sw67a_Hit:PlaySoundAt "fx_hole_enter", sw67a:Me.DestroyBall:vpmTimer.AddTimer 1500,"AddTunel":End Sub

' switch 57 under playfield

Sub sw67b_Hit:PlaySoundAt "fx_hole_enter", sw67b:Me.DestroyBall:vpmTimer.PulseSwitch 57,1000,"AddTunel":End Sub

Sub AddTunel(swNo):bsTunel.AddBall 0:End Sub

'************************************************************************************************************************

' rescue

SolCallback(10)  = "SolBallSave"

Sub SolBallSave(Enabled)
	If Enabled Then
		BallSave.Enabled = 1
	Else
		BallSave.Enabled = 0
		BallSavePrim.TransY = 0
	End If
End Sub

Sub BallSave_Hit:PlaysoundAt SoundFX("fx_solenoid",DOFContactors), BallSave:BallSave.kick 360,32:BallSave.Enabled = 0:BallSavePrim.TransY = 35:End Sub

'************************************************************************************************************************

' auto plunger

SolCallback(11) = "Auto_Plunger"

Sub Auto_Plunger(Enabled)
	If Enabled Then
	   PlungerIM.AutoFire
	End If
 End Sub

'************************************************************************************************************************

' left vuk

SolCallback(13) = "bsVUK.SolOut"

Sub sw50_Hit:PlaySoundAt "fx_kicker_enter", sw50:bsVUK.AddBall 0:End Sub

'************************************************************************************************************************

' top kicker

SolCallBack(14) = "bsPicabolas.SolOut"

Sub sw60_Hit:PlaySoundAt "fx_kicker_enter", sw60:vpmTimer.PulseSwitch 60,1000,"AddTime":End Sub

Sub AddTime(swNo):bsPicabolas.AddBall 0:End Sub

'************************************************************************************************************************

' flipper enable

SolCallback(25) = "vpmNudge.SolGameOn"

'************************************************************************************************************************

' keys

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = keyInsertCoin1 Or KeyCode = keyInsertCoin2 or KeyCode = keyInsertCoin3 Then Controller.Switch(30) = 1
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25:planeanimshake
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25:planeanimshake
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25:planeanimshake
    If Keycode = LeftFlipperKey then Controller.Switch(133) = 1
    If Keycode = RightFlipperKey then Controller.Switch(131) = 1
    If KeyCode = PlungerKey Then Controller.Switch(87) = 1
    If KeyCode = keyAdvance Then Controller.Switch(-8) = 1
    If KeyCode = keyHiScoreReset Then Controller.Switch(35) = 1
    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = keyInsertCoin1 Or KeyCode = keyInsertCoin2 or KeyCode = keyInsertCoin3 Then Controller.Switch(30) = 0
    If Keycode = LeftFlipperKey then Controller.Switch(133) = 0
    If Keycode = RightFlipperKey then Controller.Switch(131) = 0
    If KeyCode = PlungerKey Then Controller.Switch(87) = 0
    If KeyCode = keyAdvance Then Controller.Switch(-8) = 0
    If KeyCode = keyHiScoreReset Then Controller.Switch(35) = 0
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'************************************************************************************************************************

' flipper subs

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

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

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

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

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'************************************************************************************************************************

' slingshot animations

Dim RStep

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), SlingRight
    SlingRight.RotZ = 20
    RStep = 0
    vpmTimer.PulseSw 40
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:SlingRight.RotZ=15
        Case 2:SlingRight.RotZ=10
        Case 3:SlingRight.RotZ=5
        Case 4:SlingRight.RotZ=0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Dim LStep

Sub LeftSlingShot_Slingshot
	PlaySoundAt SoundFX("fx_slingshot", DOFContactors), SlingLeft
	SlingLeft.RotZ = -20
    LStep = 0
    vpmTimer.PulseSw 46
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:SlingLeft.RotZ=-15
        Case 2:SlingLeft.RotZ=-10
        Case 3:SlingLeft.RotZ=-5
        Case 4:SlingLeft.RotZ=0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'************************************************************************************************************************

' center ramp diverter is mechanical - no rom control

Sub SwRight_hit:DiverterOn.RotateToEnd:End Sub

Sub SwLeft_hit:DiverterOn.RotateToStart:End Sub

Sub DiverterOn_Animate:PrimDiverter.RotZ = DiverterOn.CurrentAngle:End Sub

'************************************************************************************************************************

' ramp lift

Sub Ramptimer_Timer
	If L71.state = 1 Then
		RampLift.RotX = -15
		LockRamp.TransZ = -110
		RampHelp.collidable = 1
	Else
		RampLift.RotX = 0
		LockRamp.TransZ = -82
		RampHelp.collidable = 0
	End If
End Sub

'************************************************************************************************************************

' light L93 motor avion

Sub Lighttimer_Timer
	If L93.state = 1 Then
	PlaySound SoundFX("fx_motor", DOFGear), 0, 1, -0.05
	WingsAnim.Enabled = 1
	Else
	Playsound ""
	End If
End Sub

' wings animation

Dim WingsStep

Sub WingsAnim_Timer
	Select Case WingsStep
		Case 1:WingLeft.ObjRotZ = 0:WingRight.ObjRotZ = 0
		Case 2:WingLeft.ObjRotZ = 3:WingRight.ObjRotZ = -3
		Case 3:WingLeft.ObjRotZ = 6:WingRight.ObjRotZ = -6
		Case 4:WingLeft.ObjRotZ = 9:WingRight.ObjRotZ = -9
		Case 5:WingLeft.ObjRotZ = 12:WingRight.ObjRotZ = -12
		Case 6:WingLeft.ObjRotZ = 15:WingRight.ObjRotZ = -15
		Case 7:WingLeft.ObjRotZ = 18:WingRight.ObjRotZ = -18
		Case 8:WingLeft.ObjRotZ = 21:WingRight.ObjRotZ = -21
		Case 9:WingLeft.ObjRotZ = 24:WingRight.ObjRotZ = -24
		Case 10:WingLeft.ObjRotZ = 27:WingRight.ObjRotZ = -27
		Case 11:WingLeft.ObjRotZ = 30:WingRight.ObjRotZ = -30
		Case 12:WingLeft.ObjRotZ = 27:WingRight.ObjRotZ = -27
		Case 13:WingLeft.ObjRotZ = 24:WingRight.ObjRotZ = -24
		Case 14:WingLeft.ObjRotZ = 21:WingRight.ObjRotZ = -21
		Case 15:WingLeft.ObjRotZ = 18:WingRight.ObjRotZ = -18
		Case 16:WingLeft.ObjRotZ = 15:WingRight.ObjRotZ = -15
		Case 17:WingLeft.ObjRotZ = 12:WingRight.ObjRotZ = -12
		Case 18:WingLeft.ObjRotZ = 9:WingRight.ObjRotZ = -9
		Case 19:WingLeft.ObjRotZ = 6:WingRight.ObjRotZ = -6
		Case 20:WingLeft.ObjRotZ = 3:WingRight.ObjRotZ = -3
		Case 21:WingLeft.ObjRotZ = 0:WingRight.ObjRotZ = 0:WingsStep = 0:WingsAnim.Enabled = 0
	End Select
    WingsStep = WingsStep + 1
End Sub

'************************************************************************************************************************

' targets

' left/right enter ramp
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' mega bumper
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' captive ball
Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' left
Sub sw51_Hit:vpmTimer.PulseSw 51:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' right
Sub sw54_Hit:vpmTimer.PulseSw 54:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw55_Hit:vpmTimer.PulseSw 55:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' top
Sub sw61_Hit:vpmTimer.PulseSw 61:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw62_Hit:vpmTimer.PulseSw 62:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw63_Hit:vpmTimer.PulseSw 63:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'************************************************************************************************************************

' bumpers

Sub sw66_Hit:vpmTimer.PulseSw 66:PlaySoundAt SoundFX("fx_bumper", DOFContactors), sw66:End Sub
Sub sw76_Hit:vpmTimer.PulseSw 76:PlaySoundAt SoundFX("fx_bumper", DOFContactors), sw76:End Sub

'************************************************************************************************************************

' hit 10 pts

Sub sw80_Hit:vpmTimer.PulseSw 80:End Sub

'************************************************************************************************************************

' hit 100 pts

Sub sw81_Hit:vpmTimer.PulseSw 81:End Sub

'************************************************************************************************************************

' switches ramp

' top

Sub sw83_Hit:vpmTimer.PulseSw 83:End Sub

' right wire ramp

Dim sw84Step

Sub sw84_hit:sw84Step = 0:sw84Anim.Enabled = 1:vpmTimer.PulseSw 84:PlaySoundAt "fx_sensor", sw84:End Sub

Sub sw84Anim_Timer
	Select Case sw84Step
		Case 1:sw84a.ObjRotZ = -62
		Case 2:sw84a.ObjRotZ = -64
		Case 3:sw84a.ObjRotZ = -66
		Case 4:sw84a.ObjRotZ = -68
		Case 5:sw84a.ObjRotZ = -70
		Case 6:sw84a.ObjRotZ = -68
		Case 7:sw84a.ObjRotZ = -66
		Case 8:sw84a.ObjRotZ = -64
		Case 9:sw84a.ObjRotZ = -62
		Case 10:sw84a.ObjRotZ = -62:sw84Step = 0:sw84Anim.Enabled = 0
	End Select
    sw84Step = sw84Step + 1
End Sub

' left wire ramp

Dim sw85Step

Sub sw85_hit:sw85Step = 0:sw85Anim.Enabled = 1:vpmTimer.PulseSw 85:PlaySoundAt "fx_sensor", sw85:End Sub

Sub sw85Anim_Timer
	Select Case sw85Step
		Case 1:sw85a.RotY = -26
		Case 2:sw85a.RotY = -24
		Case 3:sw85a.RotY = -22
		Case 4:sw85a.RotY = -20
		Case 5:sw85a.RotY = -18
		Case 6:sw85a.RotY = -20
		Case 7:sw85a.RotY = -22
		Case 8:sw85a.RotY = -24
		Case 9:sw85a.RotY = -26
		Case 10:sw85a.RotY = -26:sw85Step = 0:sw85Anim.Enabled = 0
	End Select
    sw85Step = sw85Step + 1
End Sub

'************************************************************************************************************************

' plane animation / shake

Dim PlaneShake

Sub TriggerAnim_hit:planeanimshake:End Sub
Sub Wall048_Hit:planeanimshake:End Sub

Sub planeanimshake
    PlaneShake = 3
    PlaneAnim.Enabled = 1
End Sub

Sub PlaneAnim_Timer
	Plane.ObjRotX = PlaneShake
	If PlaneShake <= 0.1 AND PlaneShake >= -0.1 Then Me.Enabled = 0:Exit Sub
	If PlaneShake < 0 Then
        PlaneShake = ABS(PlaneShake)- 0.1
	Else
        PlaneShake = - PlaneShake + 0.1
    End If
End Sub

'************************************************************************************************************************

' rollovers

Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw41:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:Lightsw42.State = 1:PlaySoundAt "fx_sensor", sw42:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:Lightsw42.State = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:PlaySoundAt "fx_sensor", sw64:End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAt "fx_sensor", sw65:End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:PlaySoundAt "fx_sensor", sw74:End Sub
Sub sw74_UnHit:Controller.Switch(74) = 0:End Sub

Sub sw75_Hit:Controller.Switch(75) = 1:PlaySoundAt "fx_sensor", sw75:End Sub
Sub sw75_UnHit:Controller.Switch(75) = 0:End Sub

Sub sw77_Hit:Controller.Switch(77) = 1:PlaySoundAt "fx_sensor", sw77:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw82_Hit:Controller.Switch(82) = 1:PlaySoundAt "fx_sensor", sw82:End Sub
Sub sw82_UnHit:Controller.Switch(82) = 0:End Sub

'************************************************************************************************************************

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

'************************************************************************************************************************

' JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array 

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

Sub UpdateLamps()

    Lamp 1, L1'lockK
    Lamp 2, L2'lockC
    Lamp 3, L3'lockO
    Lamp 4, L4'lockL

    Lamp 6, L6'machC
    Lamp 7, L7'machH
    Lamp 8, L8'ball launch
    Lamp 9, L9'machM

	Lamp 10, L10'machA
	Lamp 11, L11'machC
	Lamp 12, L12'machH

'----------------------------

'	used for Gi lights

'	sector1

'	Lamp 14, L14

'	sector2

	Lampm 15, L15
	Lampm 15, L15d
	Lampm 15, L15c
	Lampm 15, L15b
	Lamp 15, L15a

'	sector3

	Lampm 17, L17
	Lampm 17, L17c
	Lampm 17, L17b
	Lamp 17, L17a

'	sector4

	Lampm 18, L18
	Lampm 18, L18g
	Lampm 18, L18f
	Lampm 18, L18e
	Lampm 18, L18d
	Lampm 18, L18c
	Lampm 18, L18b
	Lamp 18, L18a

'----------------------------

	Lamp 19, L19'5 millions right

	Lamp 20, L20'5 millions left
	Lamp 21, L21'mega bumper
	Lamp 22, L22'prep 5 mil

	Lamp 23, L23'bumper left
	Lamp 24, L24'bumper right

    Lamp 25, L25'rescue

    Lamp 26, L26'ramp right 1
    Lamp 27, L27'ramp right 2

    Lamp 28, L28'ramp left 1
    Lamp 29, L29'ramp left 2

	Lamp 30, L30'machM
	Flash 31, L31'cabine
	Lamp 32, L32'machA

	Lamp 33, L33'red top
	Lamp 34, L34'green top
    Lamp 35, L35'yellow top

    Flashm 36, L36'red mid
    Lampm 36, L36a

    Flashm 37, L37'green mid
    Lampm 37, L37a

    Flashm 38, L38'yellow mid
    Lampm 38, L38a

'    Lamp 40, L40'credit

	Lamp 41, L41'red left
	Lamp 42, L42'green left
	Lamp 43, L43'yellow left

	Lamp 44, L44'red center
	Lamp 45, L45'green center
	Lamp 46, L46'yellow center

'-----------------------------
'	used for emblem lights

'	sector1 cab

	Lampm 47, L47
	Lampm 47, L47b
	Lamp 47, L47a

'	sector2 cab

'	Lamp 48, L48

'_____________________________

	Lamp 49, L49'score

    Lamp 50, L50'million time
    Lamp 51, L51'surprise
    Lamp 52, L52'highscore
    Lamp 53, L53'crazy switch

    Lampm 57, L57'new shot
    Lamp 57, L57a

    Lamp 58, L58'super jackpot
    Lamp 59, L59'jackpot

	Lamp 60, L60'extra ball

	Flash 61, L61'jackpot top kicker
	Flash 62, L62'extra top kicker
	Flash 63, L63'special top kicker

    Flash 64, L64'freeway

    Lamp 65, L65'center red
    Lamp 66, L66'right red
    Lamp 67, L67'left red

    Lamp 68, L68'center green
    Lamp 69, L69'right green
    Lamp 70, L70'left green

    Lamp 71, L71'ramp relay

    Flash 72, L72'extra ball captive ball

    Flash 73, L73'bonus x panel
    Flash 74, L74'jackpot panel
    Lamp 75, L75'bonus x playfield
    Lamp 76, L76'jackpot playfield

'    Lamp 77, L77'Coin Lockout

    Lamp 78, L78'center yellow
    Lamp 79, L79''right yellow
    Lamp 80, L80'left yellow

'	flashers

	Lamp 89, L89'unload plane

	Lamp 90, L90'load plane

	Lamp 91, L91'center ramp

	Lamp 92, L92'tunel

	Lamp 93, L93'motor avion

	'cabinet flashers / backglass ???

'	Lamp 94, L94
'	Lamp 95, L95
'	Lamp 96, L96

End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 20
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

'************************************************************************************************************************

' diverse collection hit sounds

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

'************************************************************************************************************************

' Supporting Ball & Sound Functions v4.0

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
    If tmp> 0 Then
        tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        Pan = Csng(tmp2)
    Else
        tmp = -tmp : tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        Pan = Csng(-tmp2)
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 200
End Function

Function BallVel(ball) 'Calculates the ball speed
    Dim bvx, bvy
    bvx = ball.VelX : bvy = ball.VelY
    BallVel = SQR(bvx*bvx + bvy*bvy)
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp, tmp2
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        AudioFade = Csng(tmp2)
    Else
        tmp = -tmp : tmp2 = tmp*tmp : tmp2 = tmp2*tmp2*tmp2 : tmp2 = tmp2*tmp2
        AudioFade = Csng(-tmp2)
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'************************************************************************************************************************

'   JP's VP10.8 Rolling Sounds

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 40 'max ball velocity
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
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    Dim bvx, bvy, bvel, bz
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        bvx = BOT(b).VelX : bvy = BOT(b).VelY : bz = BOT(b).z
        bvel = SQR(bvx*bvx + bvy*bvy)
        If bvel > 1 Then
            If bz < 30 Then
                ballpitch = bvel * 200
                ballvol = Csng((bvx*bvx + bvy*bvy) / 2000)
            Else
                ballpitch = bvel * 200 + 50000 'increase the pitch on a ramp
                ballvol = Csng((bvx*bvx + bvy*bvy) / 2000) * 2
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
        If BOT(b).VelZ < -1 and bz < 55 and bz > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, bvel * 200, 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
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

'************************************************************************************************************************

' Ball 2 Ball Collision Sound

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity * velocity / 2000), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'************************************************************************************************************************

' Real Time Flipper adjustments - by JLouLouLou & JPSalas
'        (to enable flipper tricks)

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
FullStrokeEOS_Torque = 0.3 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

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
    If LFCurAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower:End If

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
    If RFCurAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower:End If

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

'************************************************************************************************************************

' RealTime Updates

Sub LeftFlipper_Animate:leftLogo.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightLogo.RotZ = RightFlipper.CurrentAngle:End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'************************************************************************************************************************

' Dip menu by Inkochnito

' ctrDips - wrap vpmDips, provide auto-top position

Class ctrDips
	Dim vpmDips
	Dim iTop
	Sub Class_Initialize
		Set vpmDips = New cvpmDips
	End Sub
	Sub addFrame(aLeft, aTop, aWidth, aHeading, aMask, aNames)
		vpmDips.addFrame aLeft, iTop, aWidth, aHeading, aMask, aNames
		iTop = iTop + 17 + (7.5 * (UBound(aNames) + 1))
	End Sub
	Sub addChk(aLeft, aTop, aWidth, aNames)
		vpmDips.addChk aLeft, iTop, aWidth, aNames
		iTop = iTop + 16
	End Sub
	Public Sub addLabel(aLeft, aTop, aWidth, aHeight, aCaption)
		vpmDips.addLabel aLeft, iTop, aWidth, aHeight, aCaption
		iTop = iTop + 16
	End Sub
End Class

' DipVal - convert from Mach Two manual SL and Num to vpm bits

Function DipVal(sl,num)
	Dim bits
	If num < 5 Then
		' Weird - low nibble is in reverse order
		bits = 2 ^ (4 - num)
	Else
		bits = 2 ^ (num - 1)
	End If
	DipVal = bits * (&H100 ^ (sl - 1))
End Function

' myShowDips - custom dip function for Mach Two

Set vpmShowDips = GetRef("myShowDips")

Sub myShowDips
	Dim trDips : Set trDips = New ctrDips
	With trDips
		.vpmDips.AddForm 700,400,"Mach Two - DIP switches"
		.AddFrame 0,167,190,"Numero de Bolas (balls per game)", DipVal(2,1), Array("3 balls", 0 , "5 balls",DipVal(2,1))
		.AddFrame 0,76,190,"Valor Tanteo (initial replay)", DipVal(1,5) + DipVal(1,6) + DipVal(3,1), Array("100 million points", 0 , "150 million points",DipVal(1,5), "200 million points",DipVal(1,6), "250 million points", DipVal(1,5) + DipVal(1,6), "(automatically adjust)",DipVal(3,1))
		.AddFrame 0,0,190,"New Shot (ball saver)", DipVal(2,5) + DipVal(2,6), Array("multiball only", 0 , "10 seconds", DipVal(2,5), "12 seconds", DipVal(2,6), "15 seconds", DipVal(2,5) + DipVal(2,6))
		.AddFrame 0,213,190,"Subway Lock Difficulty",DipVal(3,4), Array("easy - advance L-O-C-K (see rules)", 0 , "difficult - must lock in plane first",DipVal(3,4))
		.AddFrame 0,0,190,"Especial en Avion",DipVal(3,6) + DipVal(3,7), Array("none",0, "2 balls in plane",DipVal(3,7), "3 balls in plane",DipVal(3,6))
		.AddLabel 50,310,300,20,"After hitting OK, press F3 to reset game with new settings."
		.iTop = 0
		.AddFrame 205,152,190,"Credits per coin",DipVal(1,1) + DipVal(1,2) + DipVal(1,3) + DipVal(1,4), Array("1 coin  - 1 credit",DipVal(1,1) + DipVal(1,4), "2 coins  - 3 credit",DipVal(1,1) + DipVal(1,3))
		.AddFrame 205,76,190,"Handicap (min. high score for credit)",DipVal(2,3) + DipVal(2,4), Array("250 million points", 0 , "300 million points",DipVal(2,4), "350 million points",DipVal(2,3), "400 million points",DipVal(2,3) + DipVal(2,4))
		.iTop = .iTop + 15
		.AddFrame 205,0,190,"Numero Vueltas (to lower left ramp)",DipVal(2,7) + DipVal(2,8), Array("1 center ramp shot", 0 , "3 center ramp shots",DipVal(2,7), "5 center ramp shots",DipVal(2,8), "7 center ramp shots",DipVal(2,7) + DipVal(2,8))
		.AddFrame 205,213,190,"Rescue Difficulty",DipVal(3,8), Array("easy - lit for ball launch and multiball", 0 , "difficult - lit for multiball only",DipVal(3,8))
		.vpmDips.ViewDips
	End With
End Sub
