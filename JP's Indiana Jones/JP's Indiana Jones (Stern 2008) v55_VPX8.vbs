'JP's Indiana Jones based on Stern's machine from 2008
'VPX8 table, version 5.5.0
'Indiana Jones / IPD No. 5306 / Stern April, 2008 / 4 Players

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

Const UseVPMModSol = True
LoadVPM "01210000", "SAM.VBS", 3.1

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Set GICallback = GetRef("GIUpdate")

Dim bsTrough, bsArk, bsSaucer, bsHole, aMagnet, cbLeft, cbright, plungerIM, x
Dim cpBall1, cpBall2

Const cGameName = "ij4_210"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's Indiana Jones - Stern 2008" & vbNewLine & "VPX8 table by JPSalas 5.5.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
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
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot, LeftSlingshot1, RightSlingshot1)
    vpmNudge.SolGameOn 0

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 21, 20, 19, 18, 17, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 8
    End With

    'Ark
    Set bsArk = New cvpmBallStack
    With bsArk
        .InitSw 0, 0, 0, 0, 0, 0, 0, 0
        .InitKick sw39, 190, 16
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 0
        .KickForceVar = 2
        .KickAngleVar = 10
    End With

    ' Saucer
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw45, 45, 225, 14
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    'Hole
    Set bsHole = New cvpmBallStack
    With bsHole
        .InitSw 0, 11, 0, 0, 0, 0, 0, 0
        .InitKick sw11a, 165, 16
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' Magnet
    Set aMagnet = New cvpmMagnet
    With aMagnet
        .InitMagnet Magnet, 20
        .GrabCenter = False
        .solenoid = 3
        .CreateEvents "aMagnet"
    End With

    ' Captive ball left
    Set cbLeft = New cvpmCaptiveBall
    With cbLeft
        .InitCaptive CapTrigger2, CapWall2, Array(CapKicker2a), 320
        .NailedBalls = 0
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbLeft"
        .Start
    End With
    Set cpBall2 = CapKicker2.CreateSizedBallWithMass(BallSize / 2, BallMass)

    ' Captive ball top
    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 320
        .NailedBalls = 1
        .ForceTrans = .98
        .MinForce = 3.5
        .CreateEvents "cbRight"
        .Start
    End With
    CapKicker1.CreateSizedBallWithMass BallSize / 2, BallMass
    CapKicker1b.CreateSizedBallWithMass BallSize / 2, BallMass:CapKicker1b.Kick 180, 0

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 64 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw23, IMPowerSetting, IMTime
        .Random 0.3
        .switch 23
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' SAM Fast Flips
    On Error Resume Next
    InitVpmFFlipsSAM
    If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5"
    On Error Goto 0

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    LoadLUT
    StartRainbow
    ' other switches
    Controller.Switch(50) = 0 'ark up
    Controller.Switch(51) = 1 'ark down
    Controller.Switch(52) = 0 'temple up
    Controller.Switch(53) = 1 'temple down
    Controller.Switch(63) = 1 'swordsman back
    Controller.Switch(64) = 0 'swordsman front
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = LeftMagnaSave Then bLutActive = True:SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False:HideLUT
    If vpmKeyUp(keycode) Then Exit Sub
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

Dim GiIntensity
GiIntensity = 1               'can be used by the LUT changing to increase the GI lights when the table is darker

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
    LUBack.imagea = "PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
    SetLUTLine ""
    LUBack.imagea = "PostitBL"
End Sub

'*************************
' Rainbow Changing Lights
'*************************

Dim RGBStep, RGBFactor, rRed, rGreen, rBlue

Sub StartRainbow
    RGBStep = 0
    RGBFactor = 5
    rRed = 255
    rGreen = 0
    rBlue = 0
    RainbowTimer.Enabled = 1
End Sub

Sub StopRainbow()
    Dim obj
    RainbowTimer.Enabled = 0
    RainbowTimer.Enabled = 0
End Sub

Sub RainbowTimer_Timer 'rainbow led light color changing
    ' OPT-8: Pre-compute RGB values once before the loop instead of per-object.
    '         Eliminates N-1 redundant RGB() calls (N = RainbowLights count).
    Dim obj, c, cf
    Select Case RGBStep
        Case 0 'Green
            rGreen = rGreen + RGBFactor
            If rGreen> 255 then
                rGreen = 255
                RGBStep = 1
            End If
        Case 1 'Red
            rRed = rRed - RGBFactor
            If rRed <0 then
                rRed = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            rBlue = rBlue + RGBFactor
            If rBlue> 255 then
                rBlue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            rGreen = rGreen - RGBFactor
            If rGreen <0 then
                rGreen = 0
                RGBStep = 4
            End If
        Case 4 'Red
            rRed = rRed + RGBFactor
            If rRed> 255 then
                rRed = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            rBlue = rBlue - RGBFactor
            If rBlue <0 then
                rBlue = 0
                RGBStep = 0
            End If
    End Select
    c = RGB(rRed \ 10, rGreen \ 10, rBlue \ 10)
    cf = RGB(rRed, rGreen, rBlue)
    For each obj in RainbowLights
        obj.color = c
        obj.colorfull = cf
    Next
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    DOF 101, DOFPulse
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
    DOF 102, DOFPulse
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

Dim LStep2, RStep2

Sub LeftSlingShot1_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk1
    DOF 101, DOFPulse
    LeftSling008.Visible = 1
    LeftSling005.Visible = 0
    Lemk1.RotX = 26
    LStep2 = 0
    vpmTimer.PulseSw 46
    LeftSlingShot1.TimerEnabled = 1
End Sub

Sub LeftSlingShot1_Timer
    Select Case LStep2
        Case 1:LeftSLing008.Visible = 0:LeftSLing007.Visible = 1:Lemk1.RotX = 14
        Case 2:LeftSLing007.Visible = 0:LeftSLing006.Visible = 1:Lemk1.RotX = 2
        Case 3:LeftSLing006.Visible = 0:LeftSling005.Visible = 1:Lemk1.RotX = -20:LeftSlingShot1.TimerEnabled = 0
    End Select
    LStep2 = LStep2 + 1
End Sub

Sub RightSlingShot1_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk1
    DOF 102, DOFPulse
    RightSling008.Visible = 1
    RightSling005.Visible = 0
    Remk1.RotX = 26
    RStep2 = 0
    vpmTimer.PulseSw 47
    RightSlingShot1.TimerEnabled = 1
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep2
        Case 1:RightSLing008.Visible = 0:RightSLing007.Visible = 1:Remk1.RotX = 14
        Case 2:RightSLing007.Visible = 0:RightSLing006.Visible = 1:Remk1.RotX = 2
        Case 3:RightSLing006.Visible = 0:RightSling005.Visible = 1:Remk1.RotX = -20:RightSlingShot1.TimerEnabled = 0
    End Select
    RStep2 = RStep2 + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 31:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 30:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 33:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper4:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw45_Hit:PlaysoundAt "fx_kicker_enter", sw45:bsSaucer.AddBall 0:End Sub
Sub sw39a_Hit:PlaysoundAt "fx_hole_enter", sw39a:bsArk.AddBall Me:End Sub

' Holes

Sub sw12_Hit
    PlaySoundAt "fx_hole_enter", sw12
    vpmTimer.PulseSwitch(12), 150, "bsHole.addball 0 '"
    Me.DestroyBall
End Sub

Sub sw11_Hit
    PlaySoundAt "fx_hole_enter", sw11
    bsHole.AddBall Me
End Sub

' Rollovers
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAt "fx_sensor", sw54:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw49_Hit:Controller.Switch(49) = 1:PlaySoundAt "fx_sensor", sw49:End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

' opto
Sub sw39o_Hit:Controller.Switch(39) = 1:End Sub
Sub sw39o_UnHit:Controller.Switch(39) = 0:End Sub

Sub sw40o_Hit:Controller.Switch(40) = 1:End Sub
Sub sw40o_UnHit:Controller.Switch(40) = 0:End Sub

'Spinners

Sub sw6_Spin():vpmTimer.PulseSw 6:PlaySoundAt "fx_spinner", sw6:End Sub
Sub sw14_Spin():vpmTimer.PulseSw 14:PlaySoundAt "fx_spinner", sw14:End Sub

'Targets
Sub sw1_Hit:vpmTimer.PulseSw 1:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw2_Hit:vpmTimer.PulseSw 2:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw7_Hit:vpmTimer.PulseSw 7:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw9_Hit:vpmTimer.PulseSw 9:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw59_Hit:vpmTimer.PulseSw 59:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw55_Hit:vpmTimer.PulseSw 55:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw57_Hit:vpmTimer.PulseSw 57:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "SolTrough"
SolCallback(2) = "SolAutofire"
'3 ark magnet
SolCallback(4) = "bsHole.SolOut"
SolCallback(5) = "SolBallStop"
SolCallback(6) = "SolTempleMotor"
SolCallback(7) = "solArkDiverter"
'8 shaker
' 9, 10, 11, 12 pop bumpers
'13, 14 lower slingshots
'15,16 flippers
'17, 18 top slingshots
'24 optional
SolCallback(26) = "bsSaucer.SolOut"                          'map eject
SolCallback(27) = "SolArkLid"
SolCallback(30) = "SolSwordsman"
Solcallback(33) = "SolRun"

' Flashers
SolModCallback(19) = "Flasher19"
SolModCallback(20) = "Flasher20"
SolModCallback(21) = "Flasher21"
SolModCallback(22) = "Flasher22"
SolModCallback(23) = "Flasher23"
SolModCallback(25) = "Flasher25"
SolModCallback(28) = "Flasher28"
SolModCallback(29) = "Flasher29"
SolModCallback(31) = "Flasher31"
SolModCallback(32) = "Flasher32"
SolModCallback(54) = "Flasher54"
SolModCallback(55) = "Flasher55"
SolModCallback(56) = "Flasher56"

Sub Flasher19(m): m = m /255: f19.State = m: End Sub
Sub Flasher20(m): m = m /255: f20.State = m: End Sub
Sub Flasher21(m): m = m /255: f21.State = m: f21a.State = m: f21b.State = m: f21c.State = m: End Sub
Sub Flasher22(m): m = m /255: f22.State = m: End Sub
Sub Flasher23(m): m = m /255: f23.State = m: End Sub
Sub Flasher25(m): m = m /255: f25.State = m: f25a.State = m: f25b.State = m: f25c.State = m: End Sub
Sub Flasher28(m): m = m /255: f28.State = m: End Sub
Sub Flasher29(m): m = m /255: f29.State = m: End Sub
Sub Flasher31(m): m = m /255: f31.State = m: End Sub
Sub Flasher32(m): m = m /255: f32.State = m: End Sub
Sub Flasher54(m): m = m /255: f54.State = m: f54a.State = m: End Sub
Sub Flasher55(m): m = m /255: f55.State = m: End Sub
Sub Flasher56(m): m = m /255: f56.State = m: End Sub
 
Sub SolRun(Enabled)
	vpmNudge.SolGameOn Enabled
End Sub

' Solenoid Subs
Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 22
    End If
End Sub

Sub SolAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolBallStop(Enabled)
    If Enabled Then
        BallStop.Enabled = 1
    Else
        BallStop.Enabled = 0
        BallStop.Kick 180, 0
    End If
End Sub

'Temple animation
Dim TDir, TPos
TDir = 1
TPos = 0
Sub SolTempleMotor(Enabled)
    If Enabled Then
        TempleWall.TimerEnabled = 1
        PlaySoundat("fx_motor"), temple1
    Else
        TempleWall.TimerEnabled = 0
    End If
End Sub

Sub TempleWall_Timer
    Tpos = Tpos + TDir
    If Tpos = 17 Then
        Controller.Switch(52) = 1
        CapWall2.Isdropped = 1
    Else
        Controller.Switch(52) = 0
    End If
    If Tpos = 18 Then TDir = -1
    If Tpos = 0 Then
        Controller.Switch(53) = 1
        CapWall2.IsDropped = 0
    Else
        Controller.Switch(53) = 0
    End If
    If Tpos = -1 Then TDir = 1
    cpBall2.z = 47 + Tpos * 3
    temple1.rotx = - Tpos
    temple2.rotx = - Tpos
    temple3.rotx = - Tpos
End Sub

' Ark lid animation

Dim ADir, APos
ADir = 1
APos = 0

Sub SolArkDiverter(Enabled)
    If Enabled Then
        ArkDiverter.RotateToEnd
        PlaySoundAt("fx_solenoidon"), ArkDiverter
    Else
        ArkDiverter.RotateToStart
        PlaySoundAt("fx_solenoidoff"), ArkDiverter
    End If
End Sub

Sub solArkLid(Enabled)
    If Enabled Then
        ArkWall.TimerEnabled = 1
    Else
        ArkWall.TimerEnabled = 0
    End If
End Sub

Sub ArkWall_Timer
    Apos = Apos + ADir
    PlaySoundAt("fx_motor2"), Arklid
    If Apos = 20 Then
        Controller.Switch(50) = 1
        Sw39.TimerEnabled = 1
    Else
        Controller.Switch(50) = 0
    End If
    If Apos = 21 Then ADir = -1
    If Apos = 0 Then
        Controller.Switch(51) = 1
    Else
        Controller.Switch(51) = 0
    End If
    If Apos = -1 Then ADir = 1
    Arklid.roty = - Apos
End Sub

Sub sw39_Timer
    If bsArk.Balls> 0 Then
        bsArk.ExitSol_On
    Else
        sw39.TimerEnabled = 0
    End If
End Sub

' Swordsman animation

Dim SDir, SPos
SDir = 2
SPos = -8

Sub SolSwordsman(Enabled)
    If Enabled Then
        SwordmanPost.TimerEnabled = 1
    Else
        SwordmanPost.TimerEnabled = 0
    End If
End Sub

Sub SwordmanPost_Timer
    SPos = SPos + SDir
    If SPos = 66 Then
        Controller.Switch(64) = 1
    Else
        Controller.Switch(64) = 0
    End If
    If SPos = 68 Then SDir = -2
    If SPos = -6 Then
        Controller.Switch(63) = 1
    Else
        Controller.Switch(63) = 0
    End If
    If SPos = -8 Then SDir = 2
    SwordWall.rotz = Spos
End Sub

'*******************
' Flipper Subs v5.0
'*******************

SolCallback(16) = "SolRFlipper"
SolCallback(15) = "SolLFlipper"

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

Sub LeftFlipper_Animate()
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipper_Animate()
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
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
'         Eliminates ~900 function calls/sec.
LeftFlipper.TimerInterval = 10
LeftFlipper.TimerEnabled = 1

' OPT-1: Pre-compute constant flipper strength values at init.
Dim FlipperSOSStrength : FlipperSOSStrength = FlipperPower * SOSTorque

Sub LeftFlipper_Timer 'flipper's tricks timer
    ' OPT-2: Cache CurrentAngle and StartAngle into locals.
    '         Eliminates ~2000 COM reads/sec (was 10+ reads x 1000Hz, now 6 reads x 100Hz).
    Dim lca, lsa, rca, rsa

    lca = LeftFlipper.CurrentAngle
    lsa = LeftFlipper.StartAngle

    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If lca >= lsa - SOSAngle Then LeftFlipper.Strength = FlipperSOSStrength else LeftFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If LeftFlipperOn = 1 Then
        If lca = LeftFlipper.EndAngle then
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

    rca = RightFlipper.CurrentAngle
    rsa = RightFlipper.StartAngle

    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If rca <= rsa + SOSAngle Then RightFlipper.Strength = FlipperSOSStrength else RightFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If RightFlipperOn = 1 Then
        If rca = RightFlipper.EndAngle Then
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
Sub aTargets_Hit(idx):ActiveBall.VelZ = BallVel(Activeball) * (RND / 3):End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

' OPT-3: Pre-compute inverse table dimensions. Eliminates / TableWidth and / TableHeight
'         division per call in Pan/AudioFade. Table dimensions never change at runtime.
Dim TableWidth, TableHeight
TableWidth = Table1.width
TableHeight = Table1.height
Dim InvTWHalf : InvTWHalf = 2 / TableWidth
Dim InvTHHalf : InvTHHalf = 2 / TableHeight

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    ' OPT-4: Inline BallVel, replace ^2 with multiply. Eliminates 2 extra COM reads + Exp/Log.
    Dim vx, vy : vx = ball.VelX : vy = ball.VelY
    Dim bv : bv = SQR(vx*vx + vy*vy)
    Vol = Csng(bv * bv / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table
    ' OPT-3: Use pre-computed InvTWHalf. OPT-4: Replace ^10 with multiply chain.
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
    ' OPT-4: Replace ^2 with multiply. Cache VelX/VelY into locals.
    Dim vx, vy : vx = ball.VelX : vy = ball.VelY
    BallVel = SQR(vx*vx + vy*vy)
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    ' OPT-3: Use pre-computed InvTHHalf. OPT-4: Replace ^10 with multiply chain.
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

'***********************************
'   JP's VP10.8 Rolling Sounds
'***********************************

Const tnob = 19   'total number of balls
Const lob = 5     'number of locked balls
Const maxvel = 42 'max ball velocity
ReDim rolling(tnob)

' OPT-5: Pre-built rolling sound name array. Eliminates "fx_ballrolling" & b string
'         concatenation per ball per tick (~1400 string allocs/sec at 100Hz with 14 balls).
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

Sub RollingTimer_Timer
    ' OPT-5/6/7: Cache BOT(b) via Set, cache VelX/VelY/VelZ/z into locals,
    '            inline BallVel/Vol/Pitch, use pre-built BallRollStr, cache UBound.
    Dim BOT, b, ball, ballpitch, ballvol, speedfactorx, speedfactory
    Dim bvx, bvy, bvz, bz, bv, ubBot
    BOT = GetBalls
    ubBot = UBound(BOT)

    ' stop the sound of deleted balls
    For b = ubBot + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

    ' exit the sub if no balls on the table
    If ubBot = lob - 1 Then Exit Sub

    ' play the rolling sound for each ball
    For b = lob to ubBot
        Set ball = BOT(b)
        bvx = ball.VelX : bvy = ball.VelY : bz = ball.z
        bv = SQR(bvx*bvx + bvy*bvy)

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

        ' rothbauerw's Dropping Sounds — use cached bz, read VelZ once
        bvz = ball.VelZ
        If bvz < -1 Then
            If bz < 55 Then
                If bz > 27 Then
                    PlaySound "fx_balldrop", 0, ABS(bvz) / 17, Pan(ball), 0, bv * 20, 1, 0, AudioFade(ball)
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
                    bvx = ball.VelX : bvy = ball.VelY
                End If
                If speedfactory < 1 Then
                    ball.VelX = bvx * speedfactory
                    ball.VelY = bvy * speedfactory
                End If
            End If
        End If
    Next
End Sub

'***********************
' Ball Collision Sound
'***********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    ' OPT-4: Replace ^2 with multiply.
    PlaySound("fx_collide"), 0, Csng(velocity * velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
    End Select
End Sub