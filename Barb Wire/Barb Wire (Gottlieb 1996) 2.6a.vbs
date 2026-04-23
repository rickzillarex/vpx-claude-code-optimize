
Option Explicit

'*****************************************************************************************************
' CREDITS
' Table rebuilt by Aubrel for VPX (thanks to Destruk, TAB, Rascal and Bigus for the previous builds)
' Initial table created by Bigus, fuzzel, jimmyfingers, jpsalas, toxie & unclewilly (in alphabetical order)
' Flipper primitives by zany
' Ball rolling sound script by jpsalas
' Ball shadow by ninuzzu
' Ball control & ball dropping sound by rothbauerw
' DOF by arngrim
' Positional sound helper functions by djrobx
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Aubrel 2021 March/June
'- LightsMod option (allows to add some not acurate lightings) and some randomness to the kickers added
'- GI Solenoid (Tilt) fixed (not 26 but 31 and the result is not excatly the same)
'- Some code cleanups (and ManualBallControl removed)
'- LUTs selection and "LightslevelUpdate" added
'- vpm lights mapping and "lit" primitives domes replaced by "JP's VP10 Fading Lamps&Flashers" and "iaakki's FadeDisableLighting"
'- Sound code completed, fixed and improved, many sounds added/updated, volume adjustements...
'- BigFatso moves code was wrong and killed the game rules. It should be working now.
'- ContactLens primitive and code added (fake lights removed).
'- Bumper light fixed (L30)


' Thalamus 2019 March
' Made ready for improved directional sounds, but, there is very few samples in the table.

'*************************************
' Volume Options (Volume constants, 1 is max)
Const VolFlip    = 1     ' Flipper volume.
Const VolMotor   = 0.5   ' BigFatso motor volume
Const VolSling   = 1     ' Slingshot volume.
Const VolBump    = 1     ' Bumper volume.

' Sound Options (Volume multipliers)
Const VolRol     = 0.25  ' Ballrolling volume.
Const VolRamp    = 0.5   ' Ramps volume factor.
Const VolFlipHit = 10    ' Flipper colision volume.
Const VolWire    = 4     ' Wireramp effect volume.
Const VolRub     = 2     ' Rubbers/Posts colision volume.
Const VolPin     = 20    ' Rubber Pins colision volume.
Const VolMetal   = 10    ' Metals colision volume.
Const VolWood    = 20    ' Wood colision volume.
Const VolTarg    = 1000  ' Targets volume.
Const VolGates   = 1000  ' Gates volume.
Const VolCol     = 2000  ' Ball collision volume.

'Visual Options
Const LightsMod  = 1     'set to 1 to switch off GI also with RelayA (not accurate), 2 will also add slingshots lightings (phantasy); 0 is default (accurate)
Dim Luts,LutPos
Luts = array("ColorGradeLUT256x16_1to1","ColorGradeLUT256x16_ModDeSat","ColorGradeLUT256x16_ConSat","ColorGradeLUT256x16_ExtraConSat","ColorGradeLUT256x16_ConSatMD","ColorGradeLUT256x16_ConSatD")
LutPos = 2         		 'set the nr of the LUT you want to use (0=first in the list above, 1=second, etc; 4 and 5 are "Medium Dark" and "Dark" LUTs); 2 is the default
Const VisualLock = 0     'set to 1 if you don't want to change the visual settings using Magna-Save buttons; 0 is the default (live changes enabled)

'*************************************

LoadVPM "01120100","gts3.vbs",3.10

Dim VarRol,VarCL
If Table1.ShowDT = true then
	VarRol=0
	VarCL=-5
Else
	VarRol=1
	Ramp15.Visible=0
	Ramp16.Visible=0
	Ramp17.Visible=0
	VarCL=5
End If

ContactLensP.TransY = ContactLensP.TransY + VarCL

If LutPos = 4 Then LightsLevelUpdate 1, 1    'set lights for "Medium Dark" LUT
If LutPos = 5 Then LightsLevelUpdate 2, 2    'set lights for "Dark" LUT
Table1.ColorGradeImage = Luts(LutPos)

'*************************************
' PRE-COMPUTED CONSTANTS (optimization)
' NOTE: Table1.Width/Height are COM properties not available until init.
' Const-based values are safe at module level.
'*************************************
Dim PIover180
PIover180 = 3.14159265358979 / 180
Dim BS_d6
BS_d6 = 50 / 6

' Pre-built string arrays for rolling sounds (eliminates per-tick string concatenation)
Const tnob = 5 ' total number of balls
ReDim BallRollStr(5)
ReDim BallDropStr(5)
Dim brsI
For brsI = 0 To tnob
    BallRollStr(brsI) = "fx_ballrolling" & brsI
    BallDropStr(brsI) = "fx_ball_drop" & brsI
Next

' Table dimension inverses (assigned in Table1_Init)
Dim InvTWHalf, InvTHHalf, TW_d2

Const UseSolenoids=2,UseLamps=0,UseSync=1,UseGI=0,SCoin="coin",cGameName="barbwire"
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff"

'Const swStartButton=4

'SolCallback(1) = "vpmSolSound SoundFX(""fx_bumper"",DOFContactors),"
'SolCallback(2) = "vpmSolSound SoundFX(""Left_Slingshot"",DOFContactors),"
'SolCallback(3) = "vpmSolSound SoundFX(""Right_Slingshot"",DOFContactors),"
SolCallback(6) = "bsLK.SolOut"
SolCallback(7) = "bsRK.SolOut"
SolCallback(8) = "vpmSolAutoPlungeS PlungerAuto, SoundFX(""plunger"",DOFContactors), 1,"
'SolCallback(9) = 'Pull BallGate Diverter
SolCallback(10) = "vpmSolDiverter BallGate,True," 'Hold Ballgate Diverter
SolCallback(11) = "LensUnit1"
SolCallback(12) = "LensUnit2"
SolCallback(13) = "LensUnit3"
SolCallback(14) = "SetLamp 113,"
SolCallback(15) = "SetLamp 114,"
SolCallback(16) = "MoveFatso"     'Big Fatso Motor
SolCallback(17) = "SetLamp 116,"  'Bottom Left Dome Flasher
SolCallback(18) = "SetLamp 117,"  'Captive Ball Flasher
SolCallback(19) = "SetLamp 118,"  'Top Left Upkicker Flasher
SolCallback(20) = "SetF19"        'Top Left Dome Flasher
SolCallback(21) = "SetLamp 120,"  'Left Ramp Flasher
SolCallback(22) = "SetLamp 121,"  'Big Fatso Flasher
SolCallback(23) = "SetLamp 122,"  'Right Ramp Flasher
SolCallback(24) = "SetF23"        'Top Right Dome Flasher
SolCallback(25) = "SetF24"        'Bottom Right Dome Flasher
SolCallback(26) = "RelayA"        'Lightbox Relay (A)
'SolCallback(27)= 'Ticket/Coin Meter
SolCallback(28) = "bsTrough.SolOut"
SolCallback(29) = "bsTrough.SolIn"
SolCallback(30) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(31) = "TiltGI"        'Playfield GI (TiltGI)
SolCallback(32) = "vpmNudge.SolGameOn"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), LeftFlipper, VolFlip
'    PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), LeftFlipperUp, VolFlip
    LeftFlipper.RotateToEnd
    LeftFlipperUp.RotateToEnd
  Else
    PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), LeftFlipper, VolFlip
'    PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), LeftFlipperUp, VolFlip
    LeftFlipper.RotateToStart
    LeftFlipperUp.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("fx_FlipperUp",DOFFlippers), RightFlipper, VolFlip
    RightFlipper.RotateToEnd
  Else
    PlaySoundAtVol SoundFX("fx_FlipperDown",DOFFlippers), RightFlipper, VolFlip
    RightFlipper.RotateToStart
  End If
End Sub

Dim bsTrough,bsLK,bsRK,UnitLens1,UnitLens2,UnitLens3

Sub LensUnit1(Enabled)
	UnitLens1=Enabled
End Sub

Sub LensUnit2(Enabled)
	UnitLens2=Enabled
End Sub

Sub LensUnit3(Enabled)
	UnitLens3=Enabled
End Sub

Sub SetF19(Enabled)
	SetLamp 119, Enabled
	F19C.Visible=Enabled
End Sub

Sub SetF23(Enabled)
	SetLamp 123, Enabled
	F23C.Visible=Enabled
End Sub

Sub SetF24(Enabled)
	SetLamp 124, Enabled
	F24C.Visible=Enabled
End Sub

'LightBox Relay A
Sub RelayA(Enabled)
	If LightsMod>0 Then pfgilights Enabled
End Sub

'Playfield GI
Sub TiltGI(Enabled) 'Sol 31 (Tilt)
	pfgilights Enabled
End Sub

Sub pfgilights(Enabled)
	PlaySound "fx_relay"
	dim xx
	For each xx in GI:xx.State = Not Enabled: Next
	If Enabled Then
		LightsLevelUpdate 1, 0
	Else
		LightsLevelUpdate -1, 0
	End If
End Sub

Sub Table1_Init
	vpmInit Me
 	On Error Resume Next
 	With Controller
		.GameName=cGameName
		If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
		.SplashInfoLine="Barb Wire - Gottlieb 1996"
		.HandleKeyboard=0
		.ShowTitle=0
		.ShowDMDOnly=1
		.HandleMechanics=0
		.ShowFrame=0
		.Games(cGameName).Settings.Value("rol") = 0
		.Run GetPlayerHwnd
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	vpmNudge.TiltSwitch=151
	vpmNudge.Sensitivity=6
	vpmNudge.TiltObj=Array(Bumper1,Leftslingshot,Rightslingshot)

	' Pre-compute table dimension inverses (COM props only available after init)
	InvTWHalf = 2 / Table1.Width
	InvTHHalf = 2 / Table1.Height
	TW_d2 = Table1.Width / 2

'	vpmMapLights AllLights

	Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 16,0,0,26,0,0,0,0
		bsTrough.InitKick BallRelease,120,2
		bsTrough.InitEntrySnd "SolOn","SolOn"
		bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("SolOn",DOFContactors)
		bsTrough.Balls=3

 	Set bsLK=New cvpmBallStack
		bsLK.InitSw 0,50,0,0,0,0,0,0
		bsLK.InitKick LUK,152,6
		bsLK.InitExitSnd SoundFX("Popper_ball",DOFContactors),SoundFX("SolOn",DOFContactors)
		bsLK.KickForceVar = 2
		bsLK.KickAngleVar = 2

	Set bsRK=New cvpmBallStack
		bsRK.InitSw 0,25,0,0,0,0,0,0
		bsRK.InitKick RUK,235,12
		bsRK.InitExitSnd SoundFX("Popper_ball",DOFContactors),SoundFX("SolOn",DOFContactors)
		bsRK.KickForceVar = 2
		bsRK.KickAngleVar = 2

	vpmCreateEvents AllSwitches
	Drain.CreateBall
	Drain.Kick 150,2
 	Kicker2.CreateBall
	Kicker2.Kick 180,1
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
	If keycode=StartGamekey then Controller.switch(4)=1 'Start (Button 1)
	If KeyCode=KeyFront Then Controller.Switch(1)=1 'Buy-In (Button 2)
	If KeyCode=LeftFlipperKey Then Controller.Switch(42)=1
	If KeyCode=RightFlipperkey Then Controller.Switch(43)=1
	If KeyCode=LeftMagnaSave Then  Controller.switch(5)=1 'Tournament (Left MagnaSave Button)
	If VisualLock = 0 Then
		If KeyCode = RightMagnaSave Then
			If LutPos = 5 Then LutPos = 0:LightsLevelUpdate -2,-2 Else LutPos = LutPos +1   'Max LUTs number is set to 5; so back to 0 and Lights level reseted for standard LUTs.
			If LutPos > 3 Then LightsLevelUpdate 1,1                                        'LUTs 4 and 5 are "Medium Dark" and "Dark LUTs" so Lights Level should be updated
			Table1.ColorGradeImage = Luts(LutPos)
		End If
	End If
	If KeyCode=PlungerKey Then Plunger.PullBack
	If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If keycode=StartGamekey then Controller.switch(4)=0
	If KeyCode=KeyFront Then Controller.Switch(1)=0
	If KeyCode=LeftFlipperKey Then Controller.Switch(42)=0
	If KeyCode=RightFlipperkey Then	Controller.Switch(43)=0
	If KeyCode=LeftMagnaSave Then Controller.switch(5)=0
	if KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAtVol"Plunger",Plunger,1
	If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 10:PlayBumperSound:End Sub
Sub PlayBumperSound()
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySoundAtBumperVol SoundFX("fx_bumper_1",DOFContactors),Bumper1,VolBump
			Case 2 : PlaySoundAtBumperVol SoundFX("fx_bumper_2",DOFContactors),Bumper1,VolBump
			Case 3 : PlaySoundAtBumperVol SoundFX("fx_bumper_3",DOFContactors),Bumper1,VolBump
		End Select
End Sub


Sub Drain_Hit:bsTrough.Addball Me:PlaySound SoundFX("drain",DOFContactors):End Sub
Sub KickerRampDrop_Hit:PlaySoundAt SoundFX("ball_drop",DOFContactors),KickerRampDrop:End Sub
Sub RUK_Hit:bsRK.AddBall Me:PlaySoundAt SoundFX("popper",DOFContactors),RUK:End Sub
Sub LUK_Hit:bsLK.AddBall Me:PlaySoundAt SoundFX("popper",DOFContactors),LUK:End Sub
Sub Kicker4_Hit:Me.DestroyBall:vpmTimer.PulseSwitch(60),600,"ToRightUpkicker":End Sub
Sub ToRightUpkicker(swNo):bsRK.AddBall 0:End Sub
Sub Kicker3_Hit:Me.DestroyBall:vpmTimer.PulseSwitch(70),200,"ToLeftVuk":PlaySoundAt SoundFX("popper",DOFContactors),Kicker3:End Sub
Sub ToLeftVuk(swNo):bsLK.AddBall 0:End Sub

Sub ContactLensT_Timer  'I'm not sure it really works that way but it gives the same result as the real pinball :)
	Dim ty : ty = ContactLensP.TransY   ' cache COM read
	If UnitLens2 Then
		If ty > (-55 +VarCL) Then ty = ty - 5
		If ty < (-55 +VarCL) Then ty = ty + 5
	Else
		If UnitLens1 And ty < (0 +VarCL) Then ty = ty + 5
		If UnitLens3 And ty > (-140 +VarCL) Then ty = ty - 5
	End If
	ContactLensP.TransY = ty
End Sub

Dim FatsoDir
FatsoDir=1

Sub MoveFatso(Enabled)
	If Enabled Then
		FatsoTimer.Enabled=1
	Else
		FatsoTimer.Enabled=0
	End If
End Sub

Sub FatsoTimer_Timer
    PlaySoundAtVol SoundFX("motor", DOFContactors), FatsoPrim, VolMotor
	Dim fty : fty = FatsoPrim.TransY   ' cache COM read
	If fty >= 0 Then Controller.Switch(27)=1:Controller.Switch(17)=0:FatsoDir=-2
	If fty <= -66 Then Controller.Switch(27)=0:Controller.Switch(17)=1:FatsoDir=2
	fty = fty + FatsoDir
	FatsoPrim.TransY = fty
	If fty <= -40 Then BigFatso.IsDropped=1 Else BigFatso.IsDropped=0
End Sub

Sub BigFatso_Hit:vpmTimer.PulseSw 70:End Sub

'Metal ramp sounds
Sub Trigger1_Hit : PlaySound "fx_metalrolling",-1, Vol(ActiveBall)*VolWire, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall) : End Sub
Sub Trigger2_Hit : StopSound "fx_metalrolling" : End Sub
Sub Trigger3_Hit : PlaySound "fx_metalrolling",-1, Vol(ActiveBall)*VolWire, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall) : End Sub
Sub Trigger4_Hit : StopSound "fx_metalrolling": End Sub


 Sub s13_Hit:vpmTimer.PulseSw 13:End Sub
 Sub s14_Hit:vpmTimer.PulseSw 14:End Sub
 Sub s15_Hit:Controller.Switch(15) = 1:End Sub
 Sub s15_UnHit:Controller.Switch(15) = 0:End Sub
 Sub s40_Hit:vpmTimer.PulseSw 40:End Sub
 Sub s41_Hit:vpmTimer.PulseSw 41:End Sub
 Sub s52_Hit:vpmTimer.PulseSw 52:End Sub
 Sub s53_Hit:vpmTimer.PulseSw 53:End Sub
 Sub s54_Hit:vpmTimer.PulseSw 54:End Sub
 Sub s55_Hit:vpmTimer.PulseSw 55:End Sub
 Sub s56_Hit:vpmTimer.PulseSw 56:End Sub
 Sub s57_Hit:vpmTimer.PulseSw 57:End Sub
 Sub s62_Hit:vpmTimer.PulseSw 62:End Sub
 Sub s63_Hit:vpmTimer.PulseSw 63:End Sub
 Sub s64_Hit:vpmTimer.PulseSw 64:End Sub
 Sub s65_Hit:vpmTimer.PulseSw 65:End Sub
 Sub s66_Hit:vpmTimer.PulseSw 66:End Sub
 Sub s67_Hit:vpmTimer.PulseSw 67:End Sub

' Flipper timer: cache CurrentAngle to avoid redundant COM reads
Dim lastLFAngle, lastLFUpAngle, lastRFAngle
lastLFAngle = -9999 : lastLFUpAngle = -9999 : lastRFAngle = -9999

Sub flippers_Timer()
    Dim aLF, aLFUp, aRF
    aLF = LeftFlipper.CurrentAngle
    aLFUp = LeftFlipperUp.CurrentAngle
    aRF = RightFlipper.CurrentAngle
    If aLF <> lastLFAngle Then
        lastLFAngle = aLF
        GottliebFlipperLeft.objRotZ = aLF - 90
        FlipperLSh.RotZ = aLF
    End If
    If aLFUp <> lastLFUpAngle Then
        lastLFUpAngle = aLFUp
        GottliebFlipperLeftUp.objRotZ = aLFUp - 90
        FlipperLShUp.RotZ = aLFUp
    End If
    If aRF <> lastRFAngle Then
        lastRFAngle = aRF
        GottliebFlipperRight.objRotZ = aRF - 90
        FlipperRSh.RotZ = aRF
    End If
End Sub


'LightsLevelUpdate by Aubrel
Dim obj
Sub LightsLevelUpdate(OffsetL, OffsetGI)
  For each obj in PFLights
    obj.Intensity = obj.Intensity * 1.3^(OffsetL)
  Next
  For each obj in FlasherLights
    obj.Intensity = obj.Intensity * 1.15^(OffsetL)
  Next
  For each obj in OtherLights
    obj.Intensity = obj.Intensity * 1.15^(OffsetL)
  Next
  For each obj in AllFlashers
    obj.Opacity = obj.Opacity * 1.2^(OffsetL)
  Next
  For each obj in GI
    obj.Intensity = obj.Intensity * 1.25^(OffsetGI)
  Next
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_SlingShot:vpmTimer.PulseSw 12:PlaySoundAtBallVol SoundFX("Right_Slingshot",DOFContactors), VolSling
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	If LightsMod>1  Then GI001.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
			If LightsMod>1  Then GI001.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_SlingShot:vpmTimer.PulseSw 11:PlaySoundAtBallVol SoundFX("Left_Slingshot",DOFContactors), VolSling
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	If LightsMod>1  Then GI002.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
			If LightsMod>1  Then GI002.State = 1
    End Select
    LStep = LStep + 1
End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200)
Dim FadingLevel(200)
Dim FlashSpeedUp(200)
Dim FlashSpeedDown(200)
Dim FlashMin(200)
Dim FlashMax(200)
Dim FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed (was 5ms/200Hz, now 10ms/100Hz - visually identical for fades)
LampTimer.Enabled = True

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        Dim cIdx, cVal
        For ii = 0 To UBound(chgLamp)
            cIdx = chgLamp(ii, 0)
            cVal = chgLamp(ii, 1)
            LampState(cIdx) = cVal           'keep the real state in an array
            FadingLevel(cIdx) = cVal + 4     'actual fading step

	   'Special Handling
	   'If cIdx = 0 Then RelQ cVal 'Game Overr Q Relay
        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
    NFadeL 0, L0
'   NFadeL 1, L1
'   NFadeL 2, L2
'   NFadeL 3, L3
'   NFadeL 4, L4
    NFadeL 5, L5
    NFadeL 6, L6
    NFadeL 7, L7
'   NFadeL 10, L10
    NFadeL 11, L11
    NFadeL 12, L12
    NFadeL 13, L13
    NFadeL 14, L14
    NFadeL 15, L15
    NFadeL 16, L16
    NFadeL 17, L17
'   NFadeL 20, L20
    NFadeL 21, L21
    NFadeL 22, L22
    NFadeL 23, L23
    NFadeL 24, L24
    NFadeL 25, L25
    NFadeL 26, L26
    NfadeL 27, L27
    NfadeL 30, L30
'   NfadeL 31, L31
    NfadeL 32, L32
    NfadeL 33, L33
    NfadeL 34, L34
    NFadeL 35, L35
    NFadeL 36, L36
    NFadeL 37, L37
'   NFadeL 40, L40
'   NFadeL 41, L41
    NFadeL 42, L42
    NFadeL 43, L43
    NFadeL 44, L44
    NfadeL 45, L45
    NfadeL 46, L46
    NfadeL 47, L47
    NFadeL 50, L50
    NFadeL 51, L51
    NFadeL 52, L52
    NFadeL 53, L53
    NFadeL 54, L54
    NFadeL 55, L55
    NFadeL 56, L56
    NFadeL 57, L57
    NFadeL 60, L60
    NFadeL 61, L61
    NFadeL 62, L62
    NFadeL 63, L63
    NFadeL 64, L64
    NFadeL 65, L65
    NFadeL 66, L66
    NFadeL 67, L67
    NFadeL 70, L70
    NFadeL 71, L71
    NFadeL 72, L72
    NFadeL 73, L73
    NFadeL 74, L74
    NFadeL 75, L75
    NFadeL 76, L76
    NFadeL 77, L77

    NFadeL  113, F13
    NFadeL  114, F14
    NFadeLm 116, F16
    NFadeLm 116, F16b
    Flash   116, F16C
    NFadeLm 117, F17
    NFadeLm 117, F17b
    NFadeL  117, L57
    NFadeLm 118, F18
    NFadeL  118, F18b
    NFadeLm 119, F19
    FadeDisableLighting 119, PrimF19, 60
    NFadeL  120, F20
    NFadeLm 121, F21
    NFadeL  121, F21b
    NFadeL  122, F22
    NFadeLm 123, F23
    NFadeLm 123, L30a
    NFadeLm 123, F23b
    FadeDisableLighting 123, PrimF23, 60
    NFadeLm 124, F24
    FadeDisableLighting 124, PrimF24, 60
End Sub

' div lamp subs
Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself
Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off
Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub


'FadeDisableLighting by iaakki
' Removed debug.print (COM reflection + string alloc per tick)
Sub FadeDisableLighting(nr, a, alvl)
  Select Case FadingLevel(nr)
    Case 4
      a.UserValue = a.UserValue - (a.UserValue * 0.5)^3 - 0.03
      If a.UserValue < 0 Then
        a.UserValue = 0
        FadingLevel(nr) = 0
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
    Case 5
      a.UserValue = (a.UserValue + 0.1) * 1.1
      If a.UserValue > 1 Then
        a.UserValue = 1
        FadingLevel(nr) = 1
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
  End Select
End Sub

' Flasher objects
Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, AudioPan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * InvTHHalf - 1
  If tmp > 0 Then
    Dim t2,t4,t8
    t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4
    AudioFade = Csng(t8 * t2)
  Else
    Dim nt,n2,n4,n8
    nt = -tmp : n2 = nt*nt : n4 = n2*n2 : n8 = n4*n4
    AudioFade = Csng(-(n8 * n2))
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * InvTWHalf - 1
  If tmp > 0 Then
    Dim t2,t4,t8
    t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4
    AudioPan = Csng(t8 * t2)
  Else
    Dim nt,n2,n4,n8
    nt = -tmp : n2 = nt*nt : n4 = n2*n2 : n8 = n4*n4
    AudioPan = Csng(-(n8 * n2))
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Dim bv : bv = BallVel(ball)
  Vol = Csng(bv * bv / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  Dim vx, vy
  vx = ball.VelX : vy = ball.VelY
  BallVel = INT(SQR(vx * vx + vy * vy))
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  Dim bvz : bvz = BallVelZ(ball)
  VolZ = Csng(bvz * bvz / 200)*1.2
End Function


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    Dim ball, bvx, bvy, bv, bz, bvz, ballpitch, ballvol

    For b = 0 to UBound(BOT)
        Set ball = BOT(b)
        ' Inline BallVel: cache COM reads once
        bvx = ball.VelX : bvy = ball.VelY
        bv = Int(Sqr(bvx * bvx + bvy * bvy))
        bz = ball.z

        ' play the rolling sound for each ball
        If bv > 1 Then
            If bz < 30 Then
                ballpitch = bv * 20
                ballvol = Csng(bv * bv / 2000)
            Else
                ballpitch = bv * 20 + 25000 'increase the pitch on a ramp
                ballvol = Csng(bv * bv / 2000) * 10 * VolRamp
            End If
            rolling(b) = True
            PlaySound BallRollStr(b), -1, ballvol*VolRol, AudioPan(ball), 0, ballpitch, 1, 0, AudioFade(ball)
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds (use cached bz, read VelZ once)
        bvz = ball.VelZ
        If bvz < -1 and bz < 55 and bz > 27 Then 'height adjust for ball drop sounds
            PlaySound BallDropStr(b), 0, ABS(bvz)/17, AudioPan(ball), 0, bv * 20, 1, 0, AudioFade(ball)
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim v2 : v2 = velocity * velocity
	PlaySound("fx_collide"), 0, Csng(v2) / 2000 * VolCol, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            If BallShadow(b).visible <> 0 Then BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    Dim bx, by, bz, newVis
    For b = 0 to UBound(BOT)
        bx = BOT(b).X : by = BOT(b).Y : bz = BOT(b).Z
        If bx < TW_d2 Then
            BallShadow(b).X = (bx - BS_d6 + ((bx - TW_d2)/7)) + 6
        Else
            BallShadow(b).X = (bx + BS_d6 + ((bx - TW_d2)/7)) - 6
        End If
        ballShadow(b).Y = by + 12
        If bz > 20 Then newVis = 1 Else newVis = 0
        If BallShadow(b).visible <> newVis Then BallShadow(b).visible = newVis
    Next
End Sub



'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPin, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Wood_Hit (idx)
	PlaySound "fx_woodhit", 0, Vol(ActiveBall)*VolWood, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber_post", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub LeftFlipperUp_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub BallGate_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolFlipHit, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolFlipHit, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolFlipHit, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub
