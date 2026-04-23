' Elvis - Stern 2004
' VPX8 v6.0.0 by JPSalas 2025

Option Explicit
Randomize

Const Ballsize = 50
Const Ballmass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Language Roms
Const cGameName = "elvis" 'English

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

Const UseVPMModSol = True ' it needs vpinmame 3.7

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
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SCoin = "fx_Coin"

Dim bsTrough, mMag, mCenter, dtBankL, bsUpper, bsJail, plungerIM, mElvis, x, i

'************
' Table init.
'************

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1 'enable the rom sound
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .SplashInfoLine = "Elvis - Stern 2004" & vbNewLine & "VPX8 table by JPSalas v6.0.0"
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .HandleKeyboard = 0
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
    End With

    ' Magnet
    Set mMag = New cvpmMagnet
    With mMag
        .InitMagnet Magnet, 60
        .Solenoid = 5
        .GrabCenter = 1
        .CreateEvents "mMag"
    End With

    ' Droptargets
    set dtBankL = new cvpmdroptarget
    With dtBankL
        .initdrop array(sw17, sw18, sw19, sw20, sw21), array(17, 18, 19, 20, 21)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBankL"
    End With

    ' Jail Lock
    Set bsJail = new cvpmBallStack
    With bsJail
        .InitSaucer sw34, 34, 186, 18
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickAngleVar = 1
        .KickForceVar = 1
    End With

    ' Upper Lock
    Set bsUpper = new cvpmBallStack
    With bsUpper
        .InitSaucer sw32, 32, 90, 10
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickAngleVar = 3
        .KickForceVar = 3
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 60 ' Plunger Power
    Const IMTime = 0.8        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("fx_autoplunger", DOFContactors), SoundFX("fx_autoplunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Initialize Elvis
    ElvisArms 0
    ElvisLegs 0

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1
End Sub

'****
'Keys
'****

Sub Table1_KeyDown(ByVal keycode)
    If keycode = RightFlipperKey Then Controller.Switch(88) = 1
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = keyFront Then Controller.Switch(55) = 1
    If KeyDownHandler(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = RightFlipperKey Then Controller.Switch(88) = 0
    If keycode = keyFront Then Controller.Switch(55) = 0
    If KeyUpHandler(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'**********************
'Elvis movement up/down
'**********************

Sub UpdateElvis(aNewPos)
    pStand.x = 518 + aNewPos / 3
    pLarm.x = 497 + aNewPos / 3
    pLegs.x = 516 + aNewPos / 3
    pRarm.x = 538 + aNewPos / 3
    pBody.x = 518 + aNewPos / 3
    pHead.x = 517 + aNewPos / 3
    pStand.y = 482 + aNewPos
    pLarm.y = 486 + aNewPos
    pLegs.y = 476 + aNewPos
    pRarm.y = 474 + aNewPos
    pBody.y = 482 + aNewPos
    pHead.y = 482 + aNewPos
End Sub

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(13) = "SolULFlipper"
SolCallback(14) = "SolURFlipper"

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

Sub SolULFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), LeftFlipper1
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), LeftFlipper1
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolURFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), RightFlipper1
        RightFlipper1.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), RightFlipper1
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
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

Sub LeftFlipper_Timer
    Dim LFca : LFca = LeftFlipper.CurrentAngle
    Dim LFsa : LFsa = LeftFlipper.StartAngle

    If LFca >= LFsa - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque Else LeftFlipper.Strength = FlipperPower : End If

    If LeftFlipperOn = 1 Then
        If LFca = LeftFlipper.EndAngle Then
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

    Dim RFca : RFca = RightFlipper.CurrentAngle
    Dim RFsa : RFsa = RightFlipper.StartAngle

    If RFca <= RFsa + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque Else RightFlipper.Strength = FlipperPower : End If

    If RightFlipperOn = 1 Then
        If RFca = RightFlipper.EndAngle Then
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


'*********
'Solenoids
'*********

SolCallBack(1) = "SolTrough"
SolCallBack(2) = "Auto_Plunger"
SolCallBack(3) = "dtBankL.SolDropUp"
SolCallBack(6) = "bsJail.SolOut"
SolCallBack(7) = "SolHotelLock"
SolCallBack(8) = "CGate.Open ="
SolCallBack(12) = "bsUpper.SolOut"
SolCallBack(19) = "SolHotelDoor"
SolCallBack(24) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(25) = "SolElvis" 'Solenoids 25, 26, 27 and 28

Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolHotelLock(Enabled)
    If Enabled Then
        hotelLock.IsDropped = 1
        PlaySound SoundFX("fx_flipperup", DOFContactors)
    Else
        hotelLock.IsDropped = 0
        PlaySound SoundFX("fx_flipperdown", DOFContactors)
    End If
End Sub

Sub SolHotelDoor(Enabled)
    If Enabled Then
        fDoor.RotatetoEnd
        doorwall.IsDropped = 1
        PlaySoundAt "fx_SolenoidOn", door
    Else
        fDoor.RotatetoStart
        doorwall.IsDropped = 0
        PlaySoundAt "fx_SolenoidOff", door
    End If
End Sub

Sub SolElvis(Enabled)
    ElvisTimer.Enabled = Enabled
End Sub

'Initialize Elvis position
Dim ElvisPosition
ElvisPosition = 0

Sub ElvisTimer_Timer
    ElvisPosition = ElvisPosition + Controller.getMech(0)

    'Keep Elvis from overflowing (the actual machine has overflow margin)
    If ElvisPosition <0 Then
        ElvisPosition = 0
    ElseIf ElvisPosition> 336 Then
        ElvisPosition = 336
    End If

    UpdateElvis(ElvisPosition / 2.125)

    If ElvisPosition = 0 Then
        Controller.Switch(33) = 1
    Else
        Controller.Switch(33) = 0
    End If
End Sub

'*********
' Flashers
'*********
If UseVPMModSol Then
    SolModCallback(20) = "Flasher20"
    SolModCallback(21) = "Flasher21"
    SolModCallback(22) = "Flasher22"
    SolModCallback(23) = "Flasher23"
    SolModCallback(31) = "Flasher31"
    SolModCallback(32) = "Flasher32"
    l100.Fader = 0
    l101.Fader = 0
    f22.Fader = 0
    f23.Fader = 0
    f23a.Fader = 0
    l105.Fader = 0
    f32.Fader = 0
    f32a.Fader = 0
Else
    SolCallback(20) = "vpmFlasher l100,"
    SolCallback(21) = "vpmFlasher l101,"
    SolCallback(22) = "vpmFlasher f22,"
    SolCallback(23) = "vpmFlasher Array(f23,f23a),"
    SolCallback(31) = "vpmFlasher l105,"
    SolCallback(32) = "vpmFlasher Array(f32,f32a),"
    l100.Fader = 2
    l101.Fader = 2
    f22.Fader = 2
    f23.Fader = 2
    f23a.Fader = 2
    l105.Fader = 2
    f32.Fader = 2
    f32a.Fader = 2
End If

Sub Flasher20(m):m = m / 255:l100.State = m:End Sub
Sub Flasher21(m):m = m / 255:l101.State = m:End Sub
Sub Flasher22(m):m = m / 255:f22.State = m:End Sub
Sub Flasher23(m):m = m / 255:f23.State = m:f23a.State = m:End Sub
Sub Flasher31(m):m = m / 255:l105.State = m:End Sub
Sub Flasher32(m):m = m / 255:f32.State = m:f32a.State = m:End Sub

'****************
' Elvis animation
'****************

SolCallBack(29) = "ElvisLegs"
SolCallBack(30) = "ElvisArms"

Sub ElvisLegs(Enabled)
    If Enabled Then
        fLegs.RotatetoStart
    Else
        fLegs.RotatetoEnd
    End If
End Sub

Sub ElvisArms(Enabled)
    If Enabled Then
        fArms.RotatetoStart
    Else
        fArms.RotatetoEnd
    End If
End Sub

' ************************************
' Switches, bumpers, lanes and targets
' ************************************

Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
'Sub Drain_Hit:Me.destroyball:End Sub 'debug

Sub sw9_Hit:vpmTimer.PulseSw 9:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:End Sub
Sub sw10_unHit:Controller.Switch(10) = 0:End Sub
Sub sw17_Hit:dtBankL.Hit 1:End Sub
Sub sw18_Hit:dtBankL.Hit 2:End Sub
Sub sw19_Hit:dtBankL.Hit 3:End Sub
Sub sw20_Hit:dtBankL.Hit 4:End Sub
Sub sw21_Hit:dtBankL.Hit 5:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySoundAt SoundFX("fx_target", DOFTargets), sw22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySoundAt SoundFX("fx_target", DOFTargets), sw23:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:PlaySoundAt SoundFX("fx_target", DOFTargets), sw24:End Sub
Sub sw25_Spin:PlaySoundAt "fx_spinner", sw25:vpmTimer.PulseSw 25:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", sw26:End Sub
Sub sw26_unHit:Controller.Switch(26) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_sensor", sw27:End Sub
Sub sw27_unHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_unHit:Controller.Switch(28) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAt "fx_sensor", sw30:End Sub
Sub sw30_unHit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub
Sub sw31_unHit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:playsoundAt "fx_kicker_enter", sw32:bsUpper.AddBall 0:End Sub
Sub sw34_Hit:playsoundAt "fx_kicker_enter", sw34:bsJail.AddBall 0:End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw41:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw42:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub
Sub sw45_unHit:Controller.Switch(45) = 0:End Sub
Sub sw46_Hit
    Controller.Switch(46) = 1
    x = ActiveBall.velY
    If x> 8 Then
        ActiveBall.VelY = 8
    End If
End Sub
Sub sw46_unHit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_unHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_unHit:Controller.Switch(48) = 0:End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

Sub sw52_Hit:vpmTimer.PulseSw 52:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw57_Hit::Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1::PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub
Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_Slingshot", DOFContactors), Lemk
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
    PlaySoundAt SoundFX("fx_Slingshot", DOFContactors), Remk
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

Sub doorwall_Hit:vpmTimer.PulseSw 63:PlaySoundAtBall SoundFX("fx_target", DOFTargets) End Sub

Sub sw64_Hit
    vpmTimer.PulseSw 64
    Dim abvy : abvy = ABS(ActiveBall.Vely) : str = INT(abvy * abvy * abvy)
    'debug.print str
    DogDir = 5 'upwards
    HoundDogTimer.Enabled = 1
    PlaySoundAtBall SoundFX("fx_target", DOFTargets)
End Sub

' Animate Hound Dog
Dim str 'strength of the hit
Dim DogStep, DogDir
DogStep = 0
DogDir = 0

Sub HoundDogTimer_Timer()
    DogStep = DogStep + DogDir
    HoundDog.TransZ = DogStep
    If DogStep> 100 Then DogDir = -5
    If DogStep> str Then DogDir = -5
    If DogStep <5 Then HoundDogTimer.Enabled = 0
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
    Dim chgLamp, ii, cIdx, cVal
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            cIdx = chgLamp(ii, 0) : cVal = chgLamp(ii, 1)
            LampState(cIdx) = cVal
            FadingStep(cIdx) = cVal
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
    Lamp 60, l60a
    Lamp 61, l61a
    Lamp 62, l62a
    Lamp 63, l63
    Lamp 64, l64
    Lamp 65, l65a
    Lamp 66, l66a
    Lamp 67, l67a
    Lamp 68, l68a
    Lamp 69, l69a
    Lamp 70, l70l
    Lamp 71, l71l
    Lamp 72, l72l
    Lamp 73, l73l
    Lamp 74, l74l
    Lamp 75, l75l
    Lamp 76, l76l
    Lamp 77, l77l
    Lamp 78, l78l
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
            If tmp> 0 Then
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
            If tmp> 0 Then
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

Dim InvTWHalf : InvTWHalf = 2 / TableWidth
Dim InvTHHalf : InvTHHalf = 2 / TableHeight
Dim BS_d2 : BS_d2 = Ballsize / 2

Function Vol(ball)
    Dim bv : bv = BallVel(ball) : Vol = Csng(bv * bv / 2000)
End Function

Function Pan(ball)
    Dim tmp : tmp = ball.x * InvTWHalf - 1
    If tmp > 0 Then
        Dim t2a,t4a,t8a : t2a=tmp*tmp : t4a=t2a*t2a : t8a=t4a*t4a
        Pan = Csng(t8a * t2a)
    Else
        Dim nt : nt = -tmp
        Dim t2b,t4b,t8b : t2b=nt*nt : t4b=t2b*t2b : t8b=t4b*t4b
        Pan = Csng(-(t8b * t2b))
    End If
End Function

Function Pitch(ball)
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball)
    Dim vx, vy : vx = ball.VelX : vy = ball.VelY
    BallVel = SQR(vx * vx + vy * vy)
End Function

Function AudioFade(ball)
    Dim tmp : tmp = ball.y * InvTHHalf - 1
    If tmp > 0 Then
        Dim t2a,t4a,t8a : t2a=tmp*tmp : t4a=t2a*t2a : t8a=t4a*t4a
        AudioFade = Csng(t8a * t2a)
    Else
        Dim nt : nt = -tmp
        Dim t2b,t4b,t8b : t2b=nt*nt : t4b=t2b*t2b : t8b=t4b*t4b
        AudioFade = Csng(-(t8b * t2b))
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
Const maxvel = 42 'max ball velocity
ReDim rolling(tnob)
InitRolling

Dim BallRollStr
BallRollStr = Array("fx_ballrolling0", "fx_ballrolling1", "fx_ballrolling2", "fx_ballrolling3", "fx_ballrolling4", "fx_ballrolling5", "fx_ballrolling6", "fx_ballrolling7", "fx_ballrolling8", "fx_ballrolling9", "fx_ballrolling10", "fx_ballrolling11", "fx_ballrolling12", "fx_ballrolling13", "fx_ballrolling14", "fx_ballrolling15", "fx_ballrolling16", "fx_ballrolling17", "fx_ballrolling18", "fx_ballrolling19")

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

    For b = ubBot + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
        aBallShadow(b).Y = 3000
    Next

    If ubBot < lob Then Exit Sub

    Dim bx, by, bz, bvx, bvy, bvz, bv
    For b = lob to ubBot
        bx = BOT(b).X : by = BOT(b).Y : bz = BOT(b).Z
        aBallShadow(b).X = bx
        aBallShadow(b).Y = by
        aBallShadow(b).Height = bz - BS_d2

        bvx = BOT(b).VelX : bvy = BOT(b).VelY
        bv = SQR(bvx * bvx + bvy * bvy)

        If bv > 1 Then
            If bz < 30 Then
                ballpitch = bv * 20
                ballvol = Csng(bv * bv / 2000)
            Else
                ballpitch = bv * 20 + 25000
                ballvol = Csng(bv * bv / 2000) * 2
            End If
            rolling(b) = True
            PlaySound BallRollStr(b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If

        bvz = BOT(b).VelZ
        If bvz < -1 And bz < 55 And bz > 27 Then
            PlaySound "fx_balldrop", 0, ABS(bvz) / 17, Pan(BOT(b)), 0, bv * 20, 1, 0, AudioFade(BOT(b))
        End If

        BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If bvx And bvy <> 0 Then
            speedfactorx = ABS(maxvel / bvx)
            speedfactory = ABS(maxvel / bvy)
            If speedfactorx < 1 Then
                BOT(b).VelX = bvx * speedfactorx
                BOT(b).VelY = bvy * speedfactorx
            End If
            If speedfactory < 1 Then
                BOT(b).VelX = bvx * speedfactory
                BOT(b).VelY = bvy * speedfactory
            End If
        End If
    Next
End Sub

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound "fx_collide", 0, Csng(velocity * velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
    door.RotX = fdoor.CurrentAngle
    Dim armsCA : armsCA = fArms.CurrentAngle
    pRarm.objRoty = armsCA
    pLarm.objRoty = armsCA
    pLegs.objRotY = fLegs.CurrentAngle
    LeftflipperTop.Rotz = LeftFlipper.CurrentAngle
    LeftflipperTop1.Rotz = LeftFlipper1.CurrentAngle
    RightflipperTop.Rotz = RightFlipper.CurrentAngle
    RightflipperTop1.Rotz = RightFlipper1.CurrentAngle
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")
GiUpdate 0, 0

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
End Sub

Sub Break_Hit
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
End Sub

'******
' Rules
'******
Dim Msg(20)
Sub Rules()
    Msg(0) = "Elvis - Stern 2004" &Chr(10) &Chr(10)
    Msg(1) = ""
    Msg(2) = "OBJECTIVE:Get to Graceland by lighting the following:"
    Msg(3) = "*FEATURED HITS COMPLETED (start all 5 song modes)"
    Msg(4) = "   Hound Dog (Shoot HOUND DOG Target)"
    Msg(5) = "   Blue Suede Shoes (Shoot CENTER LOOP with Upper Right Flipper)"
    Msg(6) = "   Heartbreak Hotel (Shoot balls into HEARTBREAK HOTEL on Upper Playfield)"
    Msg(7) = "   Jailhouse Rock (Shoot balls into the JAILHOUSE EJECT HOLE)"
    Msg(8) = "   All Shook Up (Shoot ALL-SHOOK shots)"
    Msg(9) = "*GIFTS FROM ELVIS COMPLETED ~ Shoot E-L-V-I-S Drop Targets to light GIFT"
    Msg(10) = " FROM ELVIS on the TOP EJECT HOLE"
    Msg(11) = "*TOP TEN COUNTDOWN COMPLETED ~ Shoot lit'music lights's to advance TOP"
    Msg(12) = " 10 COUNTDOWN."
    Msg(13) = "SKILL SHOT: Plunge ball in the WLVS Top Lanes or E-L-V-I-S Drop Targets"
    Msg(14) = "MYSTERY:Ball in the Pop Bumpers will change channels until all 3 TVs match."
    Msg(15) = "EXTRA BALL: Shoot Right Ramp ro light Extra Ball."
    Msg(16) = "TCB: Complete T-C-B to double all scoring."
    Msg(17) = "ENCORE: Spell E-N-C-O-R-E (letters lit in Back Panel) to earn"
    Msg(18) = "an Extra Multiball after the game."
    Msg(19) = ""
    Msg(20) = ""
    For X = 1 To 20
        Msg(0) = Msg(0) + Msg(X) &Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
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

    ' Elvis suit color
    x = Table1.Option("Elvis Suit Color", 0, 1, 1, 0, 0, Array("Black", "White") )
    If x Then 'white
        pbody.Image = "ElvisBodyW"
        pHead.Image = "ElvisHeadW"
        pRarm.Image = "ElvisRightArmW"
        pLegs.Image = "ElvisLegsW"
        pLarm.Image = "ElvisRightArmW"
    Else
        pbody.Image = "ElvisBody"
        pHead.Image = "ElvisHead"
        pRarm.Image = "ElvisRightArm"
        pLegs.Image = "ElvisLegs"
        pLarm.Image = "ElvisRightArm"
    End If

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