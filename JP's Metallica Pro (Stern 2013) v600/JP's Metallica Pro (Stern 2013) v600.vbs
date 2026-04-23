' *** VPX PERFORMANCE OPTIMIZATIONS APPLIED ***
' Optimized for 60fps real-time performance
' - 1. Flipper timer 1ms -> 10ms
' - 2. Cache flipper COM props (CurrentAngle) in FlipperTricks
' - 7a. Eliminate ^10 in Pan() via chain multiply
' - 7b. Eliminate ^10 in AudioFade() via chain multiply
' - 12a. Eliminate ^2 in Vol() via multiply
' - 12b. Eliminate ^2 in BallVel() and cache COM reads
' - 12c. Eliminate ^2 in OnBallBallCollision
' - 8. Pre-built BallRollStr array for rolling sounds
' - 11. Pre-computed BS_d2 = BallSize/2
' ***

'Metallica (Pro) / IPD No. 6028 / 2013 / 4 Players
'Stern Pinball, Incorporated, of Chicago, Illinois,
'VPX8 table, jpsalas 2024, v6.0.0

Option Explicit
Randomize

Const BallSize = 50
Dim BS_d2 : BS_d2 = BallSize / 2
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
        .SplashInfoLine = "Metallica (Stern 2013)" & vbNewLine & "VPX8 table by JPSalas v6.0.0"
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
    CapKicker1.CreateSizedBallWithMass BS_d2, BallMass

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
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
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

Sub GIUpdate(no, Enabled)
debug.print no
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
    Head1.rotx = ckicker.y - cball.y
    Head1.transx = (ckicker.y - cball.y) / 2
    Head1.roty = cball.x - ckicker.x
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
    Dim LFCurAngle: LFCurAngle = LeftFlipper.CurrentAngle
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
    Dim RFCurAngle: RFCurAngle = RightFlipper.CurrentAngle
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
    Dim bv: bv = BallVel(ball): Vol = Csng(bv * bv / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp> 0 Then
        Dim t2: t2 = tmp * tmp: Dim t4: t4 = t2 * t2: Dim t8: t8 = t4 * t4: Pan = Csng(t8 * t2)
    Else
        Dim nt: nt = -tmp: Dim nt2: nt2 = nt * nt: Dim nt4: nt4 = nt2 * nt2: Dim nt8: nt8 = nt4 * nt4: Pan = Csng(-(nt8 * nt2))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    Dim vx: vx = ball.VelX: Dim vy: vy = ball.VelY: BallVel = SQR(vx * vx + vy * vy)
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        Dim at2: at2 = tmp * tmp: Dim at4: at4 = at2 * at2: Dim at8: at8 = at4 * at4: AudioFade = Csng(at8 * at2)
    Else
        Dim ant: ant = -tmp: Dim ant2: ant2 = ant * ant: Dim ant4: ant4 = ant2 * ant2: Dim ant8: ant8 = ant4 * ant4: AudioFade = Csng(-(ant8 * ant2))
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
ReDim BallRollStr(tnob)
Dim iBS: For iBS = 0 to tnob: BallRollStr(iBS) = "fx_ballrolling" & iBS: Next
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

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 10
            End If
            rolling(b) = True
            PlaySound BallRollStr(b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
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
    PlaySound("fx_collide"), 0, Csng(velocity * velocity / 2000), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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