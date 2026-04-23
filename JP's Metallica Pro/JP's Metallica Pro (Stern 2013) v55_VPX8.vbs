'Metallica (Pro) / IPD No. 6028 / 2013 / 4 Players
'Stern Pinball, Incorporated, of Chicago, Illinois,
'VPX8 table, jpsalas 2024, version 5.5.1

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

const UseVPMModSol = True

LoadVPM "03060000", "SAM.VBS", 3.26

'********************
'Standard definitions
'********************

Const cGameName = "mtl_180"
Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Dim bsTrough, bsLHole, bsRHole, bsSnake
Dim cbRight, PlungerIM, LMag, RMag, x

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Metallica (Stern 2013)" & vbNewLine & "VPX8 table by JPSalas 5.5.1"
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
    vpmNudge.TiltSwitch = -7
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003, LeftSlingshot, RightSlingshot)

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 62 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 23
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 21, 20, 19, 18, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
        .IsTrough = 1
    End With

    'Hole bsRHole - right eject
    Set bsRHole = New cvpmBallStack
    With bsRHole
        .InitSw 0, 51, 0, 0, 0, 0, 0, 0
        .InitKick sw51a, 220, 25
        '.KickZ = 0.4
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
    End With

    ' Saucer
    Set bsSnake = New cvpmBallStack
    With bsSnake
        .InitSaucer sw54, 54, 195, 26
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 0
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbRight"
        .Start
    End With
    CapKicker1.CreateSizedBallWithMass BallSize / 2, BallMass

    Set LMag = New cvpmMagnet
    With LMag
        .InitMagnet LMagnet, 70
        .Solenoid = 3
        .GrabCenter = True
        .CreateEvents "LMag"
    End With

    Set RMag = New cvpmMagnet
    With RMag
        .InitMagnet RMagnet, 45
        .Solenoid = 4
        .GrabCenter = True
        .CreateEvents "RMag"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"

    'Fast Flips
    On Error Resume Next
    InitVpmFFlipsSAM
    If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5"
    On Error Goto 0

    UpPost.Isdropped = true

    LoadLUT
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    BigHeadUpdate
    RollingUpdate
End Sub

Sub Gate001_Animate:GateRight.RotY = Gate001.CurrentAngle:End Sub
Sub Gate002_Animate:GateLeft.RotY = Gate002.CurrentAngle:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = LeftMagnaSave Then bLutActive = True:SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    if keycode = "3" then bigheadshake
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False:HideLUT
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
' 3 Left Magnet
' 4 Right Magnet
SolCallback(5) = "bsSnake.SolOut"
SolCallback(6) = "bsRHole.SolOut"
SolCallback(7) = "SolPost"
' 8 shaker motor
' 9 bumper left
'10 bumper right
'11 bumper bottom
SolCallback(12) = "ResetDropTargets"
' 13 left slingshot
' 14 right slingshot
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
SolCallback(18) = "SolSparkyHead"
SolCallback(24) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
'SolCallback(x) = "vpmNudge.SolGameOn"

' Modulated flashers
SolModCallback(19) = "Flasher19"
SolModCallback(20) = "Flasher20"
SolModCallback(21) = "Flasher21"
SolModCallback(22) = "Flasher22"
SolModCallback(23) = "Flasher23"
SolModCallback(25) = "Flasher25"
SolModCallback(26) = "Flasher26"
SolModCallback(27) = "Flasher27"
SolModCallback(28) = "Flasher28"
SolModCallback(29) = "Flasher29"
SolModCallback(30) = "Flasher30"
SolModCallback(31) = "Flasher31"
SolModCallback(32) = "Flasher32"

Sub Flasher19(m):m = m /255:f19.State = m:End Sub
Sub Flasher20(m):m = m /255:f20.State = m:f20a.State = m:End Sub
Sub Flasher21(m):m = m /255:f21.State = m:End Sub
Sub Flasher22(m):m = m /255:f22.State = m:End Sub
Sub Flasher23(m):m = m /255:f23.State = m:End Sub
Sub Flasher25(m):m = m /255:f25.State = m:End Sub
Sub Flasher26(m):m = m /255:f26.State = m:f26a.State = m:End Sub
Sub Flasher27(m):m = m /255:f27.State = m:f27a.State = m:End Sub
Sub Flasher28(m):m = m /255:f28.State = m:f28a.State = m:End Sub
Sub Flasher29(m):m = m /255:f29.State = m:End Sub
Sub Flasher30(m):m = m /255:f30.State = m:f30a.State = m:f30b.State = m:End Sub
Sub Flasher31(m):m = m /255:f31.State = m:f31a.State = m:End Sub
Sub Flasher32(m):m = m /255:f32.State = m:End Sub


Sub SolPost(Enabled)
    If Enabled Then
        UpPost.Isdropped = false
    Else
        UpPost.Isdropped = true
    End If
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

' OPT-9: Removed debug.print from GIUpdate. debug.print allocates COM strings
'         and writes to debug output on every GI callback.
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
    cemeterywall.SideImage = "stone3r"
End Sub

Sub GiOff
    Dim bulb
    PlaySound "fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
    cemeterywall.SideImage = "stone3l"
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    'DOF 101, DOFPulse
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 26
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
    'DOF 102, DOFPulse
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 27
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

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 30:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 31:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw51_Hit:PlaysoundAt "fx_hole_enter", sw51:bsRHole.AddBall Me:End Sub
Sub sw54_Hit:PlaysoundAt "fx_kicker_enter", sw54:bsSnake.AddBall 0:End Sub

' Rollovers
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw1_Hit:Controller.Switch(1) = 1:PlaySoundAt "fx_sensor", sw1:End Sub
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw3_Hit:Controller.Switch(3) = 1:PlaySoundAt "fx_sensor", sw3:End Sub
Sub sw3_UnHit:Controller.Switch(3) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "fx_sensor", sw46:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw50_Hit:Controller.Switch(50) = 1:PlaySoundAt "fx_sensor", sw50:End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAt "fx_sensor", sw52:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAt "fx_sensor", sw53:End Sub
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

'Targets
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw7_Hit:vpmTimer.PulseSw 7:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw9_Hit:vpmTimer.PulseSw 9:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFTargets):LightSeqF30.Play SeqRandom, 50, , 1000:End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Drop Targets

Sub sw60_hit:Controller.switch(60) = 0:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw60:End sub
Sub sw61_hit:Controller.switch(61) = 0:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw61:End sub
Sub sw62_hit:Controller.switch(62) = 0:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw62:End sub

Sub ResetDropTargets(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_resetdrop", DOFContactors), sw61
        sw60.Isdropped = false
        Controller.switch(60) = 1
        sw61.Isdropped = False
        Controller.switch(61) = 1
        sw62.Isdropped = False
        Controller.switch(62) = 1
    End If
End Sub

'Solenoid subs

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

'***********************************************
'***********************************************
'Sparky
'***********************************************
'***********************************************
Dim cBall
BigHeadInit

Sub BigHeadInit
    Set cBall = ckicker.CreateBall
    cBall.Mass = 1.2
    ckicker.Kick 0, 0
End Sub

Sub SolSparkyHead(Enabled)
    If Enabled Then
        BigHeadShake
    End If
End Sub

Sub BigHeadShake
    cball.vely = 5 + 2 * RND(1)
    cball.velx = 2 * (RND(1) - RND(1) )
End Sub

Sub BigHeadUpdate
    ' OPT-8: Cache COM properties into locals (called every RealTime_Timer tick)
    Dim cky, cby, cbx, dy
    cky = ckicker.y
    cby = cball.y
    cbx = cball.x
    dy = cky - cby
    Head1.rotx = dy
    Head1.transx = dy / 2
    Head1.roty = cbx - ckicker.x
End Sub

'*******************
' Flipper Subs Rev3
'*******************

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

' OPT-1: Flipper timer 1ms -> 10ms. 100Hz is visually identical to 1000Hz.
LeftFlipper.TimerInterval = 10
LeftFlipper.TimerEnabled = 1

' OPT-1: Pre-compute constant flipper strength value.
Dim FlipperSOSStrength : FlipperSOSStrength = FlipperPower * SOSTorque

Sub LeftFlipper_Timer 'flipper's tricks timer
    ' OPT-2: Cache CurrentAngle/StartAngle into locals.
    Dim lca, lsa, rca, rsa

    lca = LeftFlipper.CurrentAngle
    lsa = LeftFlipper.StartAngle

    'Start Of Stroke Flipper Stroke Routine
    If lca >= lsa - SOSAngle Then LeftFlipper.Strength = FlipperSOSStrength else LeftFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine
    If LeftFlipperOn = 1 Then
        If lca = LeftFlipper.EndAngle then
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

    rca = RightFlipper.CurrentAngle
    rsa = RightFlipper.StartAngle

    'Start Of Stroke Flipper Stroke Routine
    If rca <= rsa + SOSAngle Then RightFlipper.Strength = FlipperSOSStrength else RightFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine
    If RightFlipperOn = 1 Then
        If rca = RightFlipper.EndAngle Then
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

' OPT-3: Pre-compute inverse table dimensions for Pan/AudioFade.
Dim TableWidth, TableHeight
TableWidth = Table1.width
TableHeight = Table1.height
Dim InvTWHalf : InvTWHalf = 2 / TableWidth
Dim InvTHHalf : InvTHHalf = 2 / TableHeight

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    ' OPT-4: Inline BallVel, replace ^2 with multiply.
    Dim vx, vy : vx = ball.VelX : vy = ball.VelY
    Dim bv : bv = SQR(vx*vx + vy*vy)
    Vol = Csng(bv * bv / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table
    ' OPT-3: Use InvTWHalf. OPT-4: Replace ^10 with multiply chain.
    Dim tmp, t2, t4, t8
    tmp = ball.x * InvTWHalf - 1
    If tmp > 0 Then
        t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4
        Pan = Csng(t8 * t2)
    ElseIf tmp < 0 Then
        tmp = -tmp
        t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4
        Pan = Csng(-(t8 * t2))
    Else
        Pan = 0
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    ' OPT-4: Replace ^2 with multiply. Cache VelX/VelY.
    Dim vx, vy : vx = ball.VelX : vy = ball.VelY
    BallVel = SQR(vx*vx + vy*vy)
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    ' OPT-3: Use InvTHHalf. OPT-4: Replace ^10 with multiply chain.
    Dim tmp, t2, t4, t8
    tmp = ball.y * InvTHHalf - 1
    If tmp > 0 Then
        t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4
        AudioFade = Csng(t8 * t2)
    ElseIf tmp < 0 Then
        tmp = -tmp
        t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4
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

'***********************************************
'   JP's VP10 Rolling Sounds
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 1     'number of locked balls
Const maxvel = 40 'max ball velocity
ReDim rolling(tnob)
' OPT-6: Pre-built rolling sound string array to avoid per-tick string concatenation
ReDim BallRollStr(19)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
        BallRollStr(i) = "fx_ballrolling" & i
    Next
End Sub

Sub RollingUpdate()
    ' OPT-7: Full rewrite — COM caching, inlined helpers, pre-built strings, VBS And/Or fix
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    Dim ball, bvx, bvy, bvz, bz, bvel
    BOT = GetBalls

    ' OPT-7: Cache UBound once
    Dim ub: ub = UBound(BOT)

    ' stop the sound of deleted balls (OPT-6: use pre-built strings)
    For b = ub + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

    ' exit the sub if no balls on the table
    If ub = lob - 1 Then Exit Sub

    ' play the rolling sound for each ball
    For b = lob to ub
        ' OPT-7: Cache ball object and COM properties once per iteration
        Set ball = BOT(b)
        bvx = ball.VelX
        bvy = ball.VelY
        bvz = ball.VelZ
        bz = ball.z

        ' OPT-5: Inline BallVel — vx*vx + vy*vy instead of ^2
        bvel = Sqr(bvx*bvx + bvy*bvy)

        If bvel > 1 Then
            ' OPT-5: Inline Vol — avoid double COM reads, ^2 eliminated
            ballvol = Csng(bvel) * Csng(bvel) / 2000
            ' OPT-5: Inline Pitch
            ballpitch = ballvol * 10
            If bz < 30 Then
                ' normal
            Else
                ballpitch = ballpitch + 50000
                ballvol = ballvol * 10
            End If
            rolling(b) = True
            ' OPT-6: Use pre-built string. OPT-4/3: Pan/AudioFade already optimized.
            PlaySound BallRollStr(b), -1, ballvol, Pan(ball), 0, ballpitch, 1, 0, AudioFade(ball)
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        ' OPT-7: Nested If instead of And chain (VBS non-short-circuit fix)
        If bvz < -1 Then
            If bz < 55 Then
                If bz > 27 Then
                    PlaySound "fx_balldrop", 0, ABS(bvz) / 17, Pan(ball), 0, Pitch(ball), 1, 0, AudioFade(ball)
                End If
            End If
        End If

        ' jps ball speed & spin control
        ball.AngMomZ = ball.AngMomZ * 0.95
        ' OPT-7: Fix VBS And/Or non-short-circuit bug — nested If instead
        If bvx <> 0 Then
            If bvy <> 0 Then
                speedfactorx = ABS(maxvel / bvx)
                speedfactory = ABS(maxvel / bvy)
                If speedfactorx < 1 Then
                    ball.VelX = bvx * speedfactorx
                    ball.VelY = bvy * speedfactorx
                    ' Re-cache after modification
                    bvx = ball.VelX
                    bvy = ball.VelY
                End If
                If speedfactory < 1 Then
                    ball.VelX = bvx * speedfactory
                    ball.VelY = bvy * speedfactory
                End If
            End If
        End If
    Next
End Sub

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    ' OPT-9: Replace ^2 with multiply
    PlaySound("fx_collide"), 0, Csng(velocity) * Csng(velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    Dim x
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1) MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

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
    LUBack.imagea = "PostItNote"
    String = CL2(String)
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
    SetLUTLine ""
    LUBack.imagea = "PostitBL"
End Sub

Function CL2(NumString) 'center line
    Dim Temp, TempStr
    If Len(NumString)> 40 Then NumString = Left(NumString, 40)
    Temp = (40 - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL2 = TempStr
End Function