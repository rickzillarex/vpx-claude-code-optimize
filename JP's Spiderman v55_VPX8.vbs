' JP's Spider-Man
' Based on Stern's Spider-Man Vault Edition
' Playfield & plastics redrawn by me.
' VPX8 version by JPSalas 2024, version 5.5.0 (actually this table is from 2014, made while testing VPX beta)

Option Explicit
Randomize

'********************
'Standard definitions
'********************

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseVPMModSol = True
LoadVPM "01550000", "sam.vbs", 3.26

Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0


'Const cGameName = "sman_262" 'Spiderman rom
Const cGameName = "smanve_101" 'Spiderman VE rom
'Const cGameName = "smanve_101c" 'Spiderman VE color rom

'Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "fx_coin"

'Variables
Dim bsTrough, bsSandman, bsDocOck, DocMagnet, PlungerIM, x

'************
' Table init.
'************

Sub Table1_Init
    vpminit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's Spider-Man (Stern 2007)"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = VarHidden
        On Error Resume Next
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

    On Error Goto 0

    Controller.Switch(53) = 1 'sandman down

    'Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 4

    'Sandman VUK
    Set bsSandman = New cvpmBallStack
    bsSandman.InitSw 0, 59, 0, 0, 0, 0, 0, 0
    bsSandman.InitKick sw59a, 90, 35
    bsSandman.KickZ = 1.56
    bsSandman.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSandman.InitAddSnd "fx_hole_enter"

    'Doc Ock VUK
    Set bsDocOck = New cvpmBallStack
    bsDocOck.InitSw 0, 36, 0, 0, 0, 0, 0, 0
    bsDocOck.InitKick sw36a, 90, 32
    bsDocOck.KickZ = 1.56
    bsDocOck.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsDocOck.InitAddSnd "fx_hole_enter"

    'Doc Ock Magmet
    Set DocMagnet = New cvpmMagnet
    DocMagnet.InitMagnet DocOckMagnet, 5
    DocMagnet.Solenoid = 3
    DocMagnet.GrabCenter = True
    DocMagnet.CreateEvents "DocMagnet"

    'Loop Diverter
    diverter.IsDropped = 1

    'Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    'Impulse Plunger
    Const IMPowerSetting = 55 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Switch 23
        .Random 1.5
        .InitExitSnd "fx_plunger2", "fx_plunger"
        .CreateEvents "plungerIM"
    End With

vpmMapLights aLights

'Fast Flips
	On Error Resume Next 
	InitVpmFFlipsSAM
	If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5"
	On Error Goto 0
	
	LoadLUT
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

LeftFlipper.TimerInterval = 10   ' was 1ms (1000Hz), 10ms (100Hz) is visually identical
LeftFlipper.TimerEnabled = 1

' Pre-cached flipper start/end angles (constant COM properties read once)
Dim LFStartAngle, LFEndAngle, RFStartAngle, RFEndAngle
LFStartAngle = LeftFlipper.StartAngle
LFEndAngle = LeftFlipper.EndAngle
RFStartAngle = RightFlipper.StartAngle
RFEndAngle = RightFlipper.EndAngle

Sub LeftFlipper_Timer 'flipper's tricks timer
    Dim lcaL : lcaL = LeftFlipper.CurrentAngle
    Dim lcaR : lcaR = RightFlipper.CurrentAngle

'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If lcaL >= LFStartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque Else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
	If LeftFlipperOn = 1 Then
		If lcaL = LFEndAngle Then
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
    If lcaR <= RFStartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque Else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
 	If RightFlipperOn = 1 Then
		If lcaR = RFEndAngle Then
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

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = RightFlipperKey then Controller.Switch(90) = 1
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound "fx_nudge", 0, 1, 0, 0.25
    If keycode = LeftMagnaSave Then bLutActive = True: SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
    If vpmKeyDown(Keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal Keycode)
	If keycode = LeftMagnaSave Then bLutActive = False: HideLUT
    If Keycode = RightFlipperKey then Controller.Switch(90) = 0
    If vpmKeyUp(Keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
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

'*******************
'Solenoid Callbacks
'*******************

SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "solDocMagnet"
SolCallback(4) = "solDocOckVUK"
SolCallback(5) = "solDocOck"

SolCallback(7) = "Gate3.open ="
SolCallback(8) = "Gate6.open ="

SolCallback(12) = "solSandmanVUK"
SolCallback(13) = "solSandman"

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"

SolCallback(19) = "SolGoblin" 'shake
SolCallback(20) = "sol3Bank"

SolCallback(22) = "solDivert"

'Flashers

SolModCallback(21) = "Flasher21" 'doc ock
SolModCallback(23) = "Flasher23" 'sandman x2
SolModCallback(25) = "Flasher25" 'venom x2
SolModCallback(26) = "Flasher26" 'sandman arrow
SolModCallback(27) = "Flasher27" 'sandman dome
SolModCallback(28) = "Flasher28" 'green goblin x2
SolModCallback(29) = "Flasher29" 'back panel left
SolModCallback(30) = "Flasher30" 'back panel right
SolModCallback(31) = "Flasher31" 'pop bumper x3

Sub Flasher21(m): m = m /255: f21.State = m: End Sub
Sub Flasher23(m): m = m /255: f23.State = m: f23a.State = m: End Sub
Sub Flasher25(m): m = m /255: f25.State = m: f25a.State = m: End Sub
Sub Flasher26(m): m = m /255: f26.State = m: End Sub
Sub Flasher27(m): m = m /255: f27.State = m: f27c.State = m: End Sub
Sub Flasher28(m): m = m /255: f28.State = m: f28a.State = m: End Sub
Sub Flasher29(m): m = m /255: f29.State = m: End Sub
Sub Flasher30(m): m = m /255: f30.State = m: End Sub
Sub Flasher31(m): m = m /255: f31.State = m: f31a.State = m: f31b.State = m: End Sub

'*************
' ShakeGoblin
'*************

Dim GoblinPos

Sub SolGoblin(enabled)
    If enabled Then ShakeGoblin
End Sub

Sub ShakeGoblin
    GoblinPos = 8
    GoblinShakeTimer.Enabled = 1
End Sub

Sub GoblinShakeTimer_Timer
    Goblin.TransY = GoblinPos
    Glider.TransY = GoblinPos
    If GoblinPos = 0 Then GoblinShakeTimer.Enabled = 0:Exit Sub
    If GoblinPos < 0 Then
        GoblinPos = ABS(GoblinPos) - 1
    Else
        GoblinPos = - GoblinPos + 1
    End If
End Sub

'***************
'  Doc Magnet
'***************
' Magnet power is pulsed so wait before turning power off
Sub solDocMagnet(enabled)
    MagnetOffTimer.Enabled = Not enabled
    If enabled Then DocMagnet.MagnetOn = True
End Sub

' Magnet is turned off
' Sends ball/balls to hit Doc
Sub MagnetOffTimer_Timer
    Dim ball
    For Each ball In DocMagnet.Balls 'in case there are more than one ball in the magnet
        With ball
            .VelX = 10: .VelY = -20
        End With
    Next
    Me.Enabled = False:DocMagnet.MagnetOn = False
End Sub

'Solenoid Functions
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

Dim DocDown
DocDown = True
Sub solDocOck(Enabled)
    If Enabled Then
        Controller.Switch(57) = 0
        Controller.Switch(58) = 0
        sw63.TimerInterval = 1000
        sw63.TimerEnabled = 1
    End If
End Sub

Sub sw63_Timer
    If DocDown Then
        DocDown = False
        Controller.Switch(58) = 1
        sw63.IsDropped = 1
    Else
        DocDown = True
        Controller.Switch(57) = 1
        sw63.IsDropped = 0
    End If
    sw63.TimerEnabled = 0
End Sub

Dim SandmanDown
SandmanDown = True
Sub solSandman(Enabled)
    If Enabled Then
        Controller.Switch(53) = 0
        Controller.Switch(54) = 0
        sw42.TimerInterval = 1000
        sw42.TimerEnabled = 1
    End If
End Sub

Sub sw42_Timer
    If SandmanDown Then
        SandmanDown = False
        Controller.Switch(54) = 1
        sw42.IsDropped = 1
    Else
        SandmanDown = True
        Controller.Switch(53) = 1
        sw42.IsDropped = 0
    End If
    sw42.TimerEnabled = 0
End Sub

Sub solSandmanVUK(Enabled)
    If Enabled Then
        bsSandman.ExitSol_On
    End If
End Sub

Sub solDocOckVUK(Enabled)
    If Enabled Then
        bsDocOck.ExitSol_On
    End If
End Sub

Sub solDivert(Enabled)
    If Enabled Then
        Diverter.IsDropped = 0
    Else
        Diverter.IsDropped = 1
    End If
End Sub

'Drains and Kickers
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
    ShakeLeftSpider
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
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
    ShakeRightSpider
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Shake Spiders
Dim SpiderLPos, SpiderRPos

Sub ShakeLeftSpider
    SpiderLPos = 8
    SpiderLTimer.Enabled = 1
End Sub

Sub SpiderLTimer_Timer
    SpiderL.TransY = SpiderLPos
    If SpiderLPos = 0 Then Me.Enabled = 0:Exit Sub
    If SpiderLPos < 0 Then
        SpiderLPos = ABS(SpiderLPos) - 1
    Else
        SpiderLPos = - SpiderLPos + 1
    End If
End Sub

Sub ShakeRightSpider
    SpiderRPos = 8
    SpiderRTimer.Enabled = 1
End Sub

Sub SpiderRTimer_Timer
    SpiderR.TransY = SpiderRPos
    If SpiderRPos = 0 Then Me.Enabled = 0:Exit Sub
    If SpiderRPos < 0 Then
        SpiderRPos = ABS(SpiderRPos) - 1
    Else
        SpiderRPos = - SpiderRPos + 1
    End If
End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 30:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

'Rollovers

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
Sub sw8_Hit:Controller.Switch(8) = 1:PlaySoundAt "fx_sensor", sw8:End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Right
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

'Right Under Flipper
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "fx_sensor", sw46:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

'Spinner
Sub sw7_Spin:vpmTimer.PulseSw 7:PlaySoundAt "fx_spinner", sw7:End Sub

'Right Ramp
Sub sw45_Hit:Controller.Switch(45) = 1:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

'Left Ramp
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'Venom
Sub sw43b_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43b_UnHit:Controller.Switch(43) = 0:End Sub

'Doc Ock
Sub sw63_Hit:vpmTimer.PulseSw 63:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Sandman
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Lock
Sub sw6_Hit:vpmTimer.PulseSw 6:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Sandman
Sub sw9_Hit:vpmTimer.PulseSw 9:a3BankShake:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw10_Hit:vpmTimer.PulseSw 10:a3BankShake:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11:a3BankShake:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Green Goblin
Sub sw1_Hit:vpmTimer.PulseSw 1:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw2_Hit:vpmTimer.PulseSw 2:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Right 3Bank
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Switch 14
Sub sw14_Hit():vpmTimer.PulseSw 14:End Sub

'Sandman VUK

Sub sw59_Hit():bsSandman.AddBall Me:End Sub

'DocOck VUK

Sub sw36_Hit():bsDocOck.AddBall Me:End Sub

'*************
' 3Bank Shake
'*************

Dim ccBall
Const cMod = .65 'percentage of hit power transfered to the 3 Bank of targets

a3BankInit

Sub a3BankShake
    ccball.velx = activeball.velx * cMod
    ccball.vely = activeball.vely * cMod
    a3BankTimer.enabled = True
    b3BankTimer.enabled = True
End Sub

Sub a3BankShake2 'when nudging
    a3BankTimer.enabled = True
    b3BankTimer.enabled = True
End Sub

Sub a3BankInit
    Set ccBall = hball.CreateSizedBallWithMass(25, 1.6)
    hball.Kick 0, 0
End Sub

Sub a3BankTimer_Timer            'start animation
    Dim x, y
    x = (hball.x - ccball.x) / 4 'reduce the X axis movement
    y = (hball.y - ccball.y) / 2
    backbank.transy = x
    backbank.transx = - y
    swp9.transy = x
    swp9.transx = - y
    swp10.transy = x
    swp10.transx = - y
    swp11.transy = x
    swp11.transx = - y
End Sub

Sub b3BankTimer_Timer 'stop animation
    backbank.transx = 0
    backbank.transy = 0
    swp9.transz = 0
    swp9.transx = 0
    swp10.transz = 0
    swp10.transx = 0
    swp11.transz = 0
    swp11.transx = 0
    a3BankTimer.enabled = False
    b3BankTimer.enabled = False
End Sub

'******************
'Motor Bank Up Down
'******************
Dim BankDir, BankPos
RiseBank

Sub Sol3Bank(Enabled)
    If Enabled Then
        If BankDir = 1 Then
            RiseBank
        Else
            DropBank
        End If
    End If
End Sub

Sub RiseBank()
    PlaySound "fx_motor"
    'BankPos = 52
    BankDir = -1
    Controller.Switch(49) = 0
    BankTimer.Enabled = 1
End Sub

Sub DropBank()
    PlaySound "fx_motor"
    'BankPos = 0
    BankDir = 1
    Controller.Switch(50) = 0
    BankTimer.Enabled = 1
End Sub

Sub BankTimer_Timer
    BankPos = BankPos + BankDir
    If BankPos > 52 Then
        BankPos = 52
        Me.Enabled = 0
        Controller.Switch(49) = 1
    Else
        If BankPos < 0 Then
            BankPos = 0
            Me.Enabled = 0
            Controller.Switch(50) = 1
        Else
            Update3Bank
        End If
    End If
End Sub

Sub Update3Bank
    backbank.TransZ = - BankPos
    swp9.TransZ = - BankPos
    swp10.TransZ = - BankPos
    swp11.TransZ = - BankPos
    If BankPos > 40 Then
        sw9.Isdropped = 1
        sw10.Isdropped = 1
        sw11.Isdropped = 1
    End If
    If BankPos < 10 Then
        sw9.Isdropped = 0
        sw10.Isdropped = 0
        sw11.Isdropped = 0
    End If
End Sub

'**************
' Flipper Subs
'**************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup",DOFContactors), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown",DOFContactors),LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
if UseSolenoids = 2 then Controller.Switch(swURFlip)=Enabled
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup",DOFContactors), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown",DOFContactors),RightFlipper
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
RightFlipperOn = 0
    End If
End Sub

Dim lastLFAnimAngle : lastLFAnimAngle = -9999
Dim lastRFAnimAngle : lastRFAnimAngle = -9999
Dim lastRF1AnimAngle : lastRF1AnimAngle = -9999

Sub LeftFlipper_Animate()
	Dim a : a = LeftFlipper.CurrentAngle
	If a <> lastLFAnimAngle Then lastLFAnimAngle = a : LeftFlipperTop.RotZ = a
End Sub

Sub RightFlipper_Animate()
	Dim a : a = RightFlipper.CurrentAngle
	If a <> lastRFAnimAngle Then lastRFAnimAngle = a : RightFlipperTop.RotZ = a
End Sub

Sub RightFlipper1_Animate()
	Dim a : a = RightFlipper1.CurrentAngle
	If a <> lastRF1AnimAngle Then lastRF1AnimAngle = a : RightFlipperTop1.RotZ = a
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
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

' Pre-computed inverse table dimensions
Dim InvTWHalf, InvTHHalf
InvTWHalf = 2 / TableWidth
InvTHHalf = 2 / TableHeight

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Dim bv : bv = BallVel(ball)
    Vol = Csng(bv * bv / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table
    Dim tmp, t2, t4, t8
    tmp = ball.x * InvTWHalf - 1
    If tmp > 0 Then
        t2 = tmp * tmp : t4 = t2 * t2 : t8 = t4 * t4
        Pan = Csng(t8 * t2)
    ElseIf tmp < 0 Then
        Dim nt : nt = -tmp
        t2 = nt * nt : t4 = t2 * t2 : t8 = t4 * t4
        Pan = Csng(-(t8 * t2))
    Else
        Pan = 0
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
    Dim tmp, t2, t4, t8
    tmp = ball.y * InvTHHalf - 1
    If tmp > 0 Then
        t2 = tmp * tmp : t4 = t2 * t2 : t8 = t4 * t4
        AudioFade = Csng(t8 * t2)
    ElseIf tmp < 0 Then
        Dim nt : nt = -tmp
        t2 = nt * nt : t4 = t2 * t2 : t8 = t4 * t4
        AudioFade = Csng(-(t8 * t2))
    Else
        AudioFade = 0
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
'   Ball Rolling Timer:
'   -ball rolling sounds
'   -ball speed control
'   -Rothbauer's dropping sounds
'*****************************

Const tnob = 19   'total number of balls
Const lob = 1     'number of locked balls
Const maxvel = 42 'max ball velocity
ReDim rolling(tnob)

' Pre-built rolling sound strings
ReDim BallRollStr(19)
Dim brsI : For brsI = 0 To 19 : BallRollStr(brsI) = "fx_ballrolling" & brsI : Next

InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ubBot, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls
    ubBot = UBound(BOT)

    ' stop the sound of deleted balls
    For b = ubBot + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

    ' exit the sub if no balls on the table
    If ubBot = lob - 1 Then Exit Sub

    ' play the rolling sound for each ball and draw the shadow
    Dim ball, bvx, bvy, bvz, bz, bv

    For b = lob to ubBot
        Set ball = BOT(b)
        bvx = ball.VelX : bvy = ball.VelY : bvz = ball.VelZ : bz = ball.z
        bv = Sqr(bvx * bvx + bvy * bvy)
        If bv > 1 Then
            If bz < 30 Then
                ballpitch = bv * 20
                ballvol = Csng(bv * bv / 2000)
            Else
                ballpitch = bv * 20 + 50000
                ballvol = Csng(bv * bv / 2000) * 5
            End If
            rolling(b) = True
            PlaySound BallRollStr(b), -1, ballvol, Pan(ball), 0, ballpitch, 1, 0, AudioFade(ball)
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If bvz < -1 Then
          If bz < 55 Then
            If bz > 27 Then
              PlaySound "fx_balldrop", 0, ABS(bvz) / 17, Pan(ball), 0, bv*20, 1, 0, AudioFade(ball)
            End If
          End If
        End If

        ' jps ball speed & spin control
        ball.AngMomZ = ball.AngMomZ * 0.95
        If bvx <> 0 Then
          If bvy <> 0 Then
            speedfactorx = ABS(maxvel / bvx)
            speedfactory = ABS(maxvel / bvy)
            If speedfactorx < 1 Then
                ball.VelX = bvx * speedfactorx
                ball.VelY = bvy * speedfactorx
            End If
            If speedfactory < 1 Then
                ball.VelX = ball.VelX * speedfactory
                ball.VelY = ball.VelY * speedfactory
            End If
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

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
End Sub