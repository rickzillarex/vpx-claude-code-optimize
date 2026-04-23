'World Poker Tour™ / IPD No. 5134 / February, 2006 / 4 Players
'Stern Pinball, Incorporated, of Chicago, Illinois,
'VPX8 table, jpsalas 2024, version 5.5.0

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
    ScoreText.Visible = 1
    VarHidden = 1 'hide the vpinmame dmd
Else
    UseVPMColoredDMD = False
    Scoretext.Visible = 0
    VarHidden = 0
End If

const UseVPMModSol = True 'this is not an option

LoadVPM "03060000", "SAM.VBS", 3.26

'********************
'Standard definitions
'********************

Const cGameName = "wpt_140a"
Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Dim bsTrough, bsVUK, visibleLock, bsTEject, bsSVUK, bsRScoop
Dim dtUDrop, dtLDropLower, dtLDropUpper, dtRDrop
Dim PlungerIM

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "World Joker Tour" & vbNewLine & "VPX8 table by JPSalas 5.5.0"
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
    Const IMPowerSetting = 50 'Plunger Power
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

    Set bsSVUK = New cvpmBallStack
    bsSVUK.InitSw 0, 3, 0, 0, 0, 0, 0, 0
    bsSVUK.InitKick sw3, 0, 50
    bsSVUK.Kickz = 1.5
    bsSVUK.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_solenoid", DOFContactors)

    Set bsVUK = New cvpmBallStack
    bsVUK.InitSw 0, 55, 0, 0, 0, 0, 0, 0
    bsVUK.InitKick sw55a, 180, 12
    bsVUK.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_solenoid", DOFContactors)

    Set bsRScoop = New cvpmBallStack
    bsRScoop.InitSw 0, 49, 0, 0, 0, 0, 0, 0
    bsRScoop.InitKick sw49, 0, 40
    bsRScoop.KickZ = 1.5
    bsRScoop.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_solenoid", DOFContactors)

    Set dtLDropLower = new cvpmDropTarget
    With dtLDropLower
        .Initdrop Array(sw33, sw34, sw35, sw36), Array(33, 34, 35, 36)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFDropTargets)
        .CreateEvents "dtLDropLower"
    End With

    Set dtLDropUpper = new cvpmDropTarget
    With dtLDropUpper
        .Initdrop Array(sw37, sw38, sw39, sw40), Array(37, 38, 39, 40)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFDropTargets)
        .CreateEvents "dtLDropUpper"
    End With

    Set dtUDrop = new cvpmDropTarget
    With dtUDrop
        .Initdrop Array(sw10, sw11, sw12, sw13), Array(10, 11, 12, 13)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFDropTargets)
        .CreateEvents "dtUDrop"
    End With

    Set dtRDrop = new cvpmDropTarget
    With dtRDrop
        .Initdrop Array(sw4, sw5, sw6, sw7), Array(4, 5, 6, 7)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFDropTargets)
        .CreateEvents "dtRDrop"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"

    'Fast Flips
    On Error Resume Next
    InitVpmFFlipsSAM
    If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5"
    On Error Goto 0
End Sub

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

'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "bsSVUK.SolOut"
SolCallback(4) = "bSVUK.SolOut"
SolCallback(5) = "dtLDropLower.SolDropUp"
SolCallback(6) = "dtLDropUpper.SolDropUp"
SolCallback(7) = "dtUDrop.SolDropUp"
SolCallback(8) = "dtRDrop.SolDropUp"
'(9)  'left pop bumper
'(10)  'right pop bumper
'(11)  'top pop bumper
SolCallback(12) = "SolJailUp"
SolCallback(13) = "SolULFlipper"
SolCallback(14) = "SolURFlipper"
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
'(17)'left slingshot
'(18)'right slingshot
SolCallback(19) = "SolJailLatch" 'jail latch
SolCallback(20) = "SolLPost"     'left ramp post
SolCallback(32) = "SolRpost"     'right ramp post
SolCallback(21) = "bsRScoop.SolOut"
SolCallback(24) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"

' Modulated Flashers
SolModCallback(22) = "Flasher22" 'left slingshot flasher
SolModCallback(23) = "Flasher23" 'right slingshot flasher
SolModCallback(25) = "Flasher25" 'flash left spinner
SolModCallback(26) = "Flasher26" 'back panel 1 left
SolModCallback(27) = "Flasher27" 'back panel 2
SolModCallback(28) = "Flasher28" 'back panel 3
SolModCallback(29) = "Flasher29" 'back panel 4
SolModCallback(30) = "Flasher30" 'back panel 5 right
SolModCallback(31) = "Flasher31" 'right vuk flash

Sub Flasher22(m):m = m /255:f22.State = m:f22a.State = m:End Sub
Sub Flasher23(m):m = m /255:f23.State = m:f23a.State = m:End Sub
Sub Flasher25(m):m = m /255:f25.State = m:f25a.State = m:End Sub
Sub Flasher26(m):m = m /255:f26.State = m:End Sub
Sub Flasher27(m):m = m /255:f27.State = m:End Sub
Sub Flasher28(m):m = m /255:f28.State = m:End Sub
Sub Flasher29(m):m = m /255:f29.State = m:End Sub
Sub Flasher30(m):m = m /255:f30.State = m:End Sub
Sub Flasher31(m):m = m /255:f31.State = m:f31a.State = m:End Sub

'*************************
' GI - needs new vpinmame
'*************************

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

Sub ChangeGiIntensity(scale) 'changes the intensity scale, 1 = normal
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = scale
    Next
End Sub

' leds
Set LampCallback = GetRef("LedsUpdate")

ReDim LED(69) 'from freneticamnesic table.

LED(0) = Array(D1, D2, D3, D4, D5, D6, D7)
LED(1) = Array(D8, D9, D10, D11, D12, D13, D14)
LED(2) = Array(D15, D16, D17, D18, D19, D20, D21)
LED(3) = Array(D22, D23, D24, D25, D26, D27, D28)
LED(4) = Array(D29, D30, D31, D32, D33, D34, D35)
LED(5) = Array(D36, D37, D38, D39, D40, D41, D42)
LED(6) = Array(D43, D44, D45, D46, D47, D48, D49)
LED(7) = Array(D50, D51, D52, D53, D54, D55, D56)
LED(8) = Array(D57, D58, D59, D60, D61, D62, D63)
LED(9) = Array(D64, D65, D66, D67, D68, D69, D70)
LED(10) = Array(D71, D72, D73, D74, D75, D76, D77)
LED(11) = Array(D78, D79, D80, D81, D82, D83, D84)
LED(12) = Array(D85, D86, D87, D88, D89, D90, D91)
LED(13) = Array(D92, D93, D94, D95, D96, D97, D98)
LED(14) = Array(D99, D100, D101, D102, D103, D104, D105)
LED(15) = Array(D106, D107, D108, D109, D110, D111, D112)
LED(16) = Array(D113, D114, D115, D116, D117, D118, D119)
LED(17) = Array(D120, D121, D122, D123, D124, D125, D126)
LED(18) = Array(D127, D128, D129, D130, D131, D132, D133)
LED(19) = Array(D134, D135, D136, D137, D138, D139, D140)
LED(20) = Array(D141, D142, D143, D144, D145, D146, D147)
LED(21) = Array(D148, D149, D150, D151, D152, D153, D154)
LED(22) = Array(D155, D156, D157, D158, D159, D160, D161)
LED(23) = Array(D162, D163, D164, D165, D166, D167, D168)
LED(24) = Array(D169, D170, D171, D172, D173, D174, D175)
LED(25) = Array(D176, D177, D178, D179, D180, D181, D182)
LED(26) = Array(D183, D184, D185, D186, D187, D188, D189)
LED(27) = Array(D190, D191, D192, D193, D194, D195, D196)
LED(28) = Array(D197, D198, D199, D200, D201, D202, D203)
LED(29) = Array(D204, D205, D206, D207, D208, D209, D210)
LED(30) = Array(D211, D212, D213, D214, D215, D216, D217)
LED(31) = Array(D218, D219, D220, D221, D222, D223, D224)
LED(32) = Array(D225, D226, D227, D228, D229, D230, D231)
LED(33) = Array(D232, D233, D234, D235, D236, D237, D238)
LED(34) = Array(D239, D240, D241, D242, D243, D244, D245)
LED(35) = Array(D246, D247, D248, D249, D250, D251, D252)
LED(36) = Array(D253, D254, D255, D256, D257, D258, D259)
LED(37) = Array(D260, D261, D262, D263, D264, D265, D266)
LED(38) = Array(D267, D268, D269, D270, D271, D272, D273)
LED(39) = Array(D274, D275, D276, D277, D278, D279, D280)
LED(40) = Array(D281, D282, D283, D284, D285, D286, D287)
LED(41) = Array(D288, D289, D290, D291, D292, D293, D294)
LED(42) = Array(D295, D296, D297, D298, D299, D300, D301)
LED(43) = Array(D302, D303, D304, D305, D306, D307, D308)
LED(44) = Array(D309, D310, D311, D312, D313, D314, D315)
LED(45) = Array(D316, D317, D318, D319, D320, D321, D322)
LED(46) = Array(D323, D324, D325, D326, D327, D328, D329)
LED(47) = Array(D330, D331, D332, D333, D334, D335, D336)
LED(48) = Array(D337, D338, D339, D340, D341, D342, D343)
LED(49) = Array(D344, D345, D346, D347, D348, D349, D350)
LED(50) = Array(D351, D352, D353, D354, D355, D356, D357)
LED(51) = Array(D358, D359, D360, D361, D362, D363, D364)
LED(52) = Array(D365, D366, D367, D368, D369, D370, D371)
LED(53) = Array(D372, D373, D374, D375, D376, D377, D378)
LED(54) = Array(D379, D380, D381, D382, D383, D384, D385)
LED(55) = Array(D386, D387, D388, D389, D390, D391, D392)
LED(56) = Array(D393, D394, D395, D396, D397, D398, D399)
LED(57) = Array(D400, D401, D402, D403, D404, D405, D406)
LED(58) = Array(D407, D408, D409, D410, D411, D412, D413)
LED(59) = Array(D414, D415, D416, D417, D418, D419, D420)
LED(60) = Array(D421, D422, D423, D424, D425, D426, D427)
LED(61) = Array(D428, D429, D430, D431, D432, D433, D434)
LED(62) = Array(D435, D436, D437, D438, D439, D440, D441)
LED(63) = Array(D442, D443, D444, D445, D446, D447, D448)
LED(64) = Array(D449, D450, D451, D452, D453, D454, D455)
LED(65) = Array(D456, D457, D458, D459, D460, D461, D462)
LED(66) = Array(D463, D464, D465, D466, D467, D468, D469)
LED(67) = Array(D470, D471, D472, D473, D474, D475, D476)
LED(68) = Array(D477, D478, D479, D480, D481, D482, D483)
LED(69) = Array(D484, D485, D486, D487, D488, D489, D490)

Sub LedsUpdate
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLed = Controller.ChangedLEDs(&HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF, &HFFFFFFFF)
    If Not IsEmpty(ChgLED) Then
        Dim ubLed : ubLed = UBound(chgLED)
        For ii = 0 To ubLed
            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
            For Each obj In LED(num)
                If chg And 1 Then obj.State = stat And 1
                chg = chg \ 2 : stat = stat \ 2
            Next
        Next
    End If
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
Sub sw14_Hit:vpmTimer.PulseSw 14:End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 30:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 31:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw3_Hit:PlaysoundAt "fx_kicker_enter", sw3:bsSVUK.AddBall Me:End Sub
Sub sw49_Hit:PlaysoundAt "fx_kicker_enter", sw49:bsRScoop.AddBall Me:End Sub
Sub sw55_Hit:PlaysoundAt "fx_kicker_enter", sw55:bsVUK.AddBall Me:End Sub

' Rollovers
Sub sw50_Hit:Controller.Switch(50) = 1:PlaySoundAt "fx_sensor", sw50:End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

Sub sw51_Hit:Controller.Switch(51) = 1:PlaySoundAt "fx_sensor", sw51:End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAt "fx_sensor", sw52:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAt "fx_sensor", sw53:End Sub
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAt "fx_sensor", sw54:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

Sub sw56_Hit:Controller.Switch(56) = 1:PlaySoundAt "fx_sensor", sw56:End Sub
Sub sw56_UnHit:Controller.Switch(56) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw59_Hit:Controller.Switch(59) = 1:PlaySoundAt "fx_sensor", sw59:End Sub
Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub

'Targets
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw57_Hit:vpmTimer.PulseSw 57:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw60_Hit:vpmTimer.PulseSw 60:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw61_Hit:vpmTimer.PulseSw 61:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw62_Hit:vpmTimer.PulseSw 62:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' Spinners

Sub sw8_Spin:vpmTimer.PulseSw 8:PlaySoundAt "fx_spinner", sw8:End Sub
Sub sw43_Spin:vpmTimer.PulseSw 43:PlaySoundAt "fx_spinner", sw43:End Sub

'Solenoid subs

Sub SolRpost(Enabled)
    If Enabled Then
        RPost.IsDropped = False
    Else
        RPost.IsDropped = True
    End If
End Sub

Sub SolLPost(Enabled)
    If Enabled Then
        LPost.IsDropped = False
    Else
        LPost.IsDropped = True
    End If
End Sub

Sub SolJailLatch(Enabled)
    If Enabled Then 'close jail
        sw57.IsDropped = False
        sw57a.IsDropped = False
        sw57b.IsDropped = False
        Controller.Switch(63) = 0
    End If
End Sub

Sub SolJailUp(Enabled)
    If Enabled Then
        sw57.IsDropped = true
        sw57a.IsDropped = true
        sw57b.IsDropped = true
        Controller.Switch(63) = 1
    End If
End Sub

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

'*******************
' Flipper Subs Rev3
'*******************

Sub SolULFlipper(Enabled)
    If Enabled Then
        PlayFastSoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper001
        LeftFlipper001.RotateToEnd
    Else
        PlayFastSoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper001
        LeftFlipper001.RotateToStart
    End If
End Sub

Sub SolURFlipper(Enabled)
    If Enabled Then
        PlayFastSoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper001
        RightFlipper001.RotateToEnd
    Else
        PlayFastSoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper001
        RightFlipper001.RotateToStart
    End If
End Sub

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlayFastSoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlayFastSoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlayFastSoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlayFastSoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
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
    ' Cache flipper COM properties into locals
    Dim lcaL : lcaL = LeftFlipper.CurrentAngle
    Dim lsaL : lsaL = LeftFlipper.StartAngle
    Dim leaL : leaL = LeftFlipper.EndAngle

    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If lcaL >= lsaL - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If LeftFlipperOn = 1 Then
        If lcaL = leaL then
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

    ' Cache right flipper COM properties
    Dim lcaR : lcaR = RightFlipper.CurrentAngle
    Dim lsaR : lsaR = RightFlipper.StartAngle
    Dim leaR : leaR = RightFlipper.EndAngle

    'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If lcaR <= lsaR + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If RightFlipperOn = 1 Then
        If lcaR = leaR Then
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

Sub Trigger001_Hit:PlaySoundAt "fx_PlasticHit", Trigger001:End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'***************************************************************

Dim TableWidth, TableHeight
Dim InvTWHalf, InvTHHalf

TableWidth = Table1.width
TableHeight = Table1.height
InvTWHalf = 2 / TableWidth
InvTHHalf = 2 / TableHeight

' Pre-built ball roll strings
ReDim BallRollStr(19)
Dim brsI : For brsI = 0 To 19 : BallRollStr(brsI) = "fx_ballrolling" & brsI : Next

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Dim bv : bv = BallVel(ball)
    Vol = Csng(bv * bv / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table
    Dim tmp
    tmp = ball.x * InvTWHalf - 1
    If tmp > 0 Then
        Dim t2p, t4p, t8p : t2p = tmp*tmp : t4p = t2p*t2p : t8p = t4p*t4p
        Pan = Csng(t8p * t2p)
    Else
        Dim ntp : ntp = -tmp
        Dim n2p, n4p, n8p : n2p = ntp*ntp : n4p = n2p*n2p : n8p = n4p*n4p
        Pan = Csng(-(n8p * n2p))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    Dim vx, vy : vx = ball.VelX : vy = ball.VelY
    BallVel = SQR(vx*vx + vy*vy)
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * InvTHHalf - 1
    If tmp > 0 Then
        Dim t2f, t4f, t8f : t2f = tmp*tmp : t4f = t2f*t2f : t8f = t4f*t4f
        AudioFade = Csng(t8f * t2f)
    Else
        Dim ntf : ntf = -tmp
        Dim n2f, n4f, n8f : n2f = ntf*ntf : n4f = n2f*n2f : n8f = n4f*n4f
        AudioFade = Csng(-(n8f * n2f))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound with a small random pitch at X and Y position of an object, mostly bumpers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlayFastSoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly flippers, without a random pitch
    PlaySound soundname, 0, 1, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'*************************************************************
'   JP's VP 10.8 Rolling Sounds & Ball speed and spin control
'*************************************************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 46 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls
    Dim ubBot : ubBot = UBound(BOT)

    ' stop the sound of deleted balls
    For b = ubBot + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

    ' exit the sub if no balls on the table
    If ubBot = lob - 1 Then Exit Sub

    ' play the rolling sound for each ball (inlined BallVel, cached COM reads)
    Dim ball, bvx, bvy, bvz, bz, bv
    For b = lob to ubBot
        Set ball = BOT(b)
        bvx = ball.VelX : bvy = ball.VelY : bvz = ball.VelZ : bz = ball.z
        bv = SQR(bvx*bvx + bvy*bvy)
        If bv > 1 Then
            If bz < 30 Then
                ballpitch = bv * 20
                ballvol = Csng(bv * bv / 2000)
            Else
                ballpitch = bv * 20 + 50000
                ballvol = Csng(bv * bv / 2000) * 3
            End If
            rolling(b) = True
            PlaySound BallRollStr(b), -1, ballvol, Pan(ball), 0, ballpitch, 1, 0, AudioFade(ball)
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds (use cached bvz/bz)
        If bvz < -1 Then
            If bz < 55 and bz > 27 Then PlaySound "fx_balldrop", 0, ABS(bvz) / 17, Pan(ball), 0, bv*20, 1, 0, AudioFade(ball)
            If bz < 10 and bz > -10 Then PlaySound "fx_hole_enter", 0, ABS(bvz) / 17, Pan(ball), 0, bv*20, 1, 0, AudioFade(ball)
        End If

        ' jps ball speed & spin control
            ball.AngMomZ = ball.AngMomZ * 0.95
        If bvx AND bvy <> 0 Then
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
    Next
End Sub

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity * velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
    ScoreText.visible = x

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