' The X Files
' Based on The X Files / IPD No. 4137 / September 09, 1997 / 6 Players
' http://www.ipdb.org/machine.cgi?id=4137
' VP91x version 1.01 by JPSalas March 2011
' VP10.2 version 1.0 by Sliderpoint September 2016
' VP10.2 4k version by Hanibal


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim UseVPMDMD
UseVPMDMD = 1

LoadVPM "01560000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************
Const cGameName = "xfiles"
Const UseSolenoids = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin"

Dim bsTrough, mFCMag, mechFC, UpdateGI
Dim DesktopMode: DesktopMode = Table1.ShowDT

'************
' Table init.
'************
  If DesktopMode = True Then 
		RRail.Visible = 1
		LRail.Visible = 1
		Primitive149.visible = 1
		Primitive148.visible = 1
  Else
		RRail.Visible = 0
		LRail.Visible = 0
		Primitive149.visible = 0
		Primitive148.visible = 0
  End If

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "The X Files" & vbNewLine & "VP10 4k Edition by Hanibal"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 15, 14, 13, 12, 0, 0, 0
        .InitKick BallRelease, 90, 8
        .InitExitSnd "ballrel", "Solenoid"
        .Balls = 4
    End With

    'File Cabinet Magnet
    Set mFCMag = New cvpmMagnet
    With mFCMag
        .InitMagnet fcMagnet, 40
        .Solenoid = 14
        .GrabCenter = 0
    End With

    ' File Cabinet Motor
    Set mechFC = New cvpmMech
    With mechFC
        .MType = vpmMechLinear + vpmMechReverse + vpmMechOneSol + vpmMechLengthSw
        .Sol1 = 21
        .Length = 135
        .Steps = 135
        .AddSw 45, 0, 4
        .AddSw 44, 67, 69
        .AddSw 43, 131, 135
        .CallBack = GetRef("UpdateFC")
        .Start
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMameInterval
    PinMAMETimer.Enabled = 1

	'other Init
    plunger.pullback
    Trapdoor1.Collidable = 0:Trapdoor2.IsDropped = 1:Trapdoor3.IsDropped = 1
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********
Sub table1_KeyDown(ByVal Keycode)
    If KeyDownHandler(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then controller.switch(53) = true
    If keycode = LeftTiltKey Then LeftNudge 110, 1.3, 20:PlaySound "nudge_left"
    If keycode = RightTiltKey Then RightNudge 250, 1.3, 20:PlaySound "nudge_right"
    If keycode = CenterTiltKey Then CenterNudge 180, 1.8, 25:PlaySound "nudge_forward"
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyUpHandler(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then controller.switch(53) = false
End Sub

'*********
' Switches
'*********
' Slings & div switches
Sub LeftSlingShot_Slingshot:PlaySound SoundFX("slingshot",DOFContactors):vpmTimer.PulseSw 59:End Sub
Sub RightSlingShot_Slingshot:PlaySound SoundFX("slingshot",DOFContactors):vpmTimer.PulseSw 62:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySound SoundFX("bumper1",DOFContactors):End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySound SoundFX("bumper2",DOFContactors):End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySound SoundFX("bumper1",DOFContactors):End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaySound "drain":vpmTimer.PulseSw 10:bsTrough.AddBall Me:UpdateGI = 0:updateGITimer:End Sub
Sub sw39_Hit:PlaySound "hole_enter":vpmTimer.PulseSw 39:UpdateGI = 2:updateGITimer:End Sub
Sub sw42a_Hit:PlaySound "hole_enter":End Sub:
Sub sw42_hit:vpmTimer.PulseSw 42:End Sub:
Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub:
sub sw41a_Hit:PlaySound "hole_enter":UpdateGI = 2:updateGITimer:end sub
Sub sw48_Hit:controller.Switch(48) = 1:end Sub
Sub sw48_unHit:controller.Switch(48) = 0:End sub

' Rollovers & Ramp Switches
Sub sw16_Hit:Controller.Switch(16) = 1:PlaySound "sensor":Plungerlight.state = 1:UpdateGI = 1:updateGITimer:End Sub
Sub sw16_Unhit:Controller.Switch(16) = 0:Plungerlight.state = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySound "sensor":End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "sensor":UpdateGI = 0:UpdateGITimer:GIFlash.enabled = 1:End Sub
Sub sw26_Unhit:Controller.Switch(26) = 0:PlaySound "metalrolling":End Sub
Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "sensor":End Sub
Sub sw27_Unhit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "sensor":UpdateGI = 0:UpdateGITimer:GIFlash.enabled = 1:End Sub
Sub sw28_Unhit:Controller.Switch(28) = 0:PlaySound "metalrolling":End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:End Sub
Sub sw46_Unhit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "sensor":End Sub
Sub sw47_Unhit:Controller.Switch(47) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySound "sensor":End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1:PlaySound "sensor":End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "sensor":End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySound "sensor":End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "sensor":UpdateGI = 0:UpdateGITimer:GIFlash.enabled = 1:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "sensor":End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:PlaySound "sensor":UpdateGI = 0:UpdateGITimer:GIFlash.enabled = 1:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySound "sensor":End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "sensor":UpdateGI = 0:UpdateGITimer:GIFlash.enabled = 1:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw57_Hit:Controller.Switch(57) = 1:PlaySound "sensor":UpdateGI = 0:UpdateGITimer:GIFlash.enabled = 1:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub sw58_Hit:Controller.Switch(58) = 1:PlaySound "sensor":End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub sw61_Hit:Controller.Switch(61) = 1:PlaySound "sensor":End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1:PlaySound "sensor":UpdateGI = 0:UpdateGITimer:GIFlash.enabled = 1:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

' Targets
Sub sw17_Hit:vpmTimer.PulseSw 17:sw17.TransY = 5:TargetTimer.Enabled = 1:PlaySound "target":End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:sw18.TransY = 5:TargetTimer.Enabled = 1:PlaySound "target":End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:sw19.TransY = 5::TargetTimer.Enabled = 1:PlaySound "target":End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:sw20.TransY = -5:TargetTimer.Enabled = 1:PlaySound "target":End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:sw21.TransY = -5:TargetTimer.Enabled = 1:PlaySound "target":End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:sw22.TransY = -5:TargetTimer.Enabled = 1:PlaySound "target":End Sub

Sub TargetTimer_Timer
	sw17.TransY = 0
	sw18.TransY = 0
	sw19.transy = 0
	sw20.TransY = 0
	sw21.transY = 0
	sw22.TransY = 0
	TargetTimer.Enabled = 0
end Sub

'*********
'Solenoids
'*********
SolCallBack(1) = "bsTrough.SolOut"
SolcallBack(2) = "vpmsolautoplunger plunger,32,"
Solcallback(3) = "VUK"
Solcallback(17) = "SolTrapdoor"
Solcallback(18) = "SolRampdiv"
Solcallback(19) = "SolLoopdiv"
' (20) ' Not used 
' (21) 'FileCabinet Motor 
SolCallback(22) = "Flasher22" 'Fash top Mid x2
'SolCallback(23) = "Flasher23" 'Flash Trapdoor x1 under trapdoor (not visible?) 
SolCallback(25) = "Flasher25" 'Flash RT Ramp Top x 1
SolCallback(26) = "Flasher26"   'Flash Pops x2
SolCallback(27) = "Flasher27"   'Flash LT Ramp x2
SolCallback(28) = "Flasher28"   'Flash Bot LT x2
SolCallback(29) = "Flasher29"	'Flash Bot RT x1
SolCallback(30) = "Flasher30"	'Flash RT Ramp Top x1
SolCallback(31) = "Flasher31" 'Flash RT Ramp BOT x1
SolCallback(32) = "Flasher32" 'Flash Miniloop x1

Sub SolTrapdoor(Enabled)
    If Enabled Then
        trapdoor1.Collidable = 1
		TDCase = 1
		sw42a.timerenabled = 1
        trapdoor2.isdropped = false
        trapdoor3.isdropped = false
        sw42a.Enabled = true
        controller.switch(40) = true
		Trappdoorlight.state = 1

    Else
        trapdoor1.Collidable = 0
		TDCase = 3
		sw42a.timerenabled = 1
        trapdoor2.isdropped = true
        trapdoor3.isdropped = true
        sw42a.Enabled = false
        controller.switch(40) = false
		Trappdoorlight.state = 0
    End If
End Sub

Dim TDCase
Sub sw42a_timer
	Select Case TDCase
			Case 1: TrapDoor.objRotX = 10:TrapDoor.objRotY = 2: TDCase = 2
			Case 2: TrapDoor.ObjRotX = 25:TrapDoor.objRotY = 5: PlaySound SoundFX("Trapdoor",DOFContactors): Sw42a.TimerEnabled = 0
			Case 3:	TrapDoor.ObjRotX = 10:TrapDoor.objRotY = 2: TDCase = 4
			Case 4: TrapDoor.ObjRotX = 0:TrapDoor.objRotY = 0: PlaySound SoundFX("Trapdoor",DOFContactors): sw42a.timerEnabled = 0
	End Select
End Sub

Sub VUK(enabled)
	sw48.kickz 0, 86, 0, 175
	Playsound SoundFX("solenoid",DOFContactors)
	controller.Switch(48)=0
End Sub

Sub SolRampdiv(Enabled)
    If Enabled Then
        Diverter1.isdropped = false
		diverter.ObjRoty = 0
    Else
        Diverter1.isdropped = true
		diverter.ObjRotY = 15
    End If
End Sub

Sub SolLoopdiv(Enabled)
    If Enabled Then
        loopdiv.isdropped = false
    Else
        loopdiv.isdropped = true
    end if
End Sub

Sub fcMagnet_Hit():mFCMag.addball activeball:End Sub
Sub fcMagnet_Unhit():mFCMag.removeball activeball:End Sub

'**************
' Flipper Subs
'**************
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX(SFlipperOn,DOFFlippers):LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX(SFlipperOff,DOFFlippers):LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX(SFlipperOn,DOFFlippers):RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX(SFlipperOff,DOFFlippers):RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "rubber_flipper"
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "rubber_flipper"
End Sub

' *********************
' File Cabinet subs
' *********************
Dim FCFlashOn, FCCurrPos, FCLastPos
FCFlashOn = 0:FCCurrPos = 0

Sub UpdateFC(aCurrPos, aSpeed, aLastPos)
	PlaySound SoundFX("Motor1",DOFGear)
    FCCurrPos = aCurrPos
	FileCabinet.transY = (aCurrPos -135) 
	FileCabinetDoor.Z = (aCurrPos +75)
	CabLightG.TransY = (aCurrPos -135)
	CabLightR.TransY = (aCurrPos -135)
	CabLightSrew1.TransY = (aCurrPos -135)
	CabLightSrew2.TransY = (aCurrPos -135)
	l73.bulbHaloHeight = aCurrPos +135
	l74.bulbhaloHeight = aCurrPos +135
	'l001.bulbhaloHeight = aCurrPos +150
	If aCurrPos < 4 Then CabFront.isDropped = 1
	If aCurrPos > 4 Then CabFront.isDropped = 0
end Sub

Sub cabtimer_Timer
		If controller.Switch(45) = True Then
		CabFront.isdropped = 1
	Else
		CabFront.isdropped = 0
	End If
	me.enabled = 0
End Sub

Sub CabFront_Hit
	UpdateGI = 0
	UpdateGITimer
	PlaySound "PlastikHit", 0, Vol(ActiveBall)*2, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	GIFlash.enabled = 1
End Sub

'*****************
'Flasher Caps subs
'*****************
Dim PrimCase1, PrimCase2, PrimCase3, PrimCase4





Sub Flasher22(Enabled)
	If Enabled Then
	F22.State = 1
	F22a.State = 1
	F22b.State = 1
	F22c.State = 1
	Flasher22x.visible = 1
	Else
	F22.State = 0
	F22a.State = 0
	F22b.State = 0
	F22c.State = 0
	Flasher22x.visible = 0
	End If
End Sub

Sub Flasher25(Enabled)
	If Enabled Then
	F1.State = 1
	F1a.State = 1
	Else
	F1.State = 0
	F1a.State = 0
	End If
End Sub

Sub Flasher26(Enabled)
	If Enabled Then
		F2.State = 1
		F2a.State = 1
		F2b.State = 1
		F2c.State = 1
	Else
		F2.State = 0
		F2a.State = 0
		F2b.State = 0
		F2c.State = 0
	End If
End Sub

Sub Flasher27(Enabled)
	If Enabled Then
		Light27.State = 1
		Light27a.State = 1
		Light27a1.State = 1
		Light27b.State = 1
		Flasher2.visible = 1
		Flasher2x.state = 1
		PrimCase1 = 1
		FTimer27.Enabled = 1
		Light28.State = 1
		Light28a.State = 1
		Light28a1.State = 1
		Light28b.State = 1
		Flasher1.visible = 1
		Flasher1x.state = 1
		PrimCase2 = 1
		FTimer28.Enabled = 1
	Else
		Light27.State = 0
		Light27a.State = 0
		Light27a1.State = 0
		Light27b.State = 0
		Flasher2.visible = 0
		Flasher2x.state = 0
		PrimCase1 = 4
		FTimer27.Enabled = 1
		Light28.State = 0
		Light28a.State = 0
		Light28a1.State = 0
		Light28b.State = 0
		Flasher1.visible = 0
		Flasher1x.state = 0
		PrimCase2 = 4
		FTimer28.Enabled = 1
	End If
End Sub

Sub Flasher28(Enabled)
	If Enabled Then
	F4.State = 1
	F4a.State = 1
	Flasher4z.visible = 1
	Else
	F4.State = 0
	F4a.State = 0
	Flasher4z.visible = 0
	End If
End Sub

Sub Flasher29(Enabled)
	If Enabled Then
	F5.State = 1
	F5a.State = 1
	Flasher5z.visible = 1
	Else
	F5.State = 0
	F5a.State = 0
	Flasher5z.visible = 0
	End If
End Sub


Sub Flasher30(Enabled)
	If Enabled Then
		Light30.State = 1
		Light30a.State = 1
		Light30a1.State = 1
		Light30b.State = 1
		Flasher3.visible = 1
		Flasher3x.state = 1
		Flukeman.image = "FlukemanB_ON1"
		PrimCase4 = 1
		FTimer30.Enabled = 1
	Else
		Light30.State = 0
		Light30a.State = 0
		Light30a1.State = 0
		Light30b.State = 0
		Flasher3.visible = 0
		Flasher3x.state = 0
		Flukeman.image = "Flukeman_CompleteMap"
		PrimCase4 = 4
		FTimer30.Enabled = 1
	End If
End Sub

Sub Flasher31(Enabled)
	If Enabled Then
		Light31.State = 1
		Light31a.State = 1
		Light31a1.State = 1
		Light31b.State = 1
		Flasher4.visible = 1
		Flasher4x.state = 1
		Flukeman.image = "Flukeman_ON1"
		Tunnel.image = "FTunnel_ON1"
		PrimCase3 = 1
		FTimer31.Enabled = 1
	Else
		Light31.State = 0
		Light31a.State = 0
		Light31a1.State = 0
		Light31b.State = 0
		Flasher4.visible = 0
		Flasher4x.state = 0
		Flukeman.image = "Flukeman_CompleteMap"
		Tunnel.image = "FTunnel_CompleteMap"
		PrimCase3 = 4
		FTimer31.Enabled = 1
	End If
End Sub

Sub Flasher32(Enabled)
	If Enabled Then
	F8.State = 1
	F8a.State = 1
	Else
	F8.State = 0
	F8a.State = 0
	End If
End Sub

Sub FTimer27_Timer()
    Select Case PrimCase1
		Case 1:Prim27.image = "dome3_clear_B":PrimCase1 = 2
        Case 2:Prim27.image = "dome3_clear_A":PrimCase1 = 3
        Case 3:Prim27.image = "dome3_clear_ON":Me.Enabled = 0
        Case 4:Prim27.image = "dome3_clear_A":PrimCase1 = 5
        Case 5:Prim27.image = "dome3_clear_B":PrimCase1 = 6
        Case 6:Prim27.image = "dome3_clear_OFF":Me.Enabled = 0
    End Select
End Sub

Sub FTimer28_Timer()
    Select Case PrimCase2
        Case 1:Prim28.image = "dome3_clear_B":PrimCase2 = 2
        Case 2:Prim28.image = "dome3_clear_A":PrimCase2 = 3
        Case 3:Prim28.image = "dome3_clear_ON":Me.Enabled = 0
        Case 4:Prim28.image = "dome3_clear_A":PrimCase2 = 5
        Case 5:Prim28.image = "dome3_clear_B":PrimCase2 = 6
        Case 6:Prim28.image = "dome3_clear_OFF":Me.Enabled = 0
    End Select
End Sub

Sub FTimer30_Timer()
    Select Case PrimCase4
        Case 1:Prim30.image = "dome3_clear_B":PrimCase4 = 2
        Case 2:Prim30.image = "dome3_clear_A":PrimCase4 = 3
        Case 3:Prim30.image = "dome3_clear_ON":Me.Enabled = 0
        Case 4:Prim30.image = "dome3_clear_A":PrimCase4 = 5
        Case 5:Prim30.image = "dome3_clear_B":PrimCase4 = 6
        Case 6:Prim30.image = "dome3_clear_OFF":Me.Enabled = 0
    End Select
End Sub

Sub FTimer31_Timer()
     Select Case PrimCase3
         Case 1:Prim31.image = "dome3_clear_B":PrimCase3 = 2
         Case 2:Prim31.image = "dome3_clear_A":PrimCase3 = 3
         Case 3:Prim31.image = "dome3_clear_ON":Me.Enabled = 0
         Case 4:Prim31.image = "dome3_clear_A":PrimCase3 = 5
         Case 5:Prim31.image = "dome3_clear_B":PrimCase3 = 6
         Case 6:Prim31.image = "dome3_clear_OFF":Me.Enabled = 0
    End Select
End Sub

Set LampCallback = GetRef("Lamps") ' individual lampcallbacks - optimized: read each Controller.Lamp once, assign to multiples
	Sub Lamps
		Dim s
		L1.State = Controller.Lamp(1)
		L2.State = Controller.Lamp(2)
		L3.State = Controller.Lamp(3)
		L4.State = Controller.Lamp(4)
		L5.State = Controller.Lamp(5)
		L6.State = Controller.Lamp(6)
		L7.State = Controller.Lamp(7)
		L8.State = Controller.Lamp(8)
		s = Controller.Lamp(9) : L9.State = s : l9a.State = s
		s = Controller.Lamp(10) : L10.State = s : l10a.State = s
		s = Controller.Lamp(11) : L11.State = s : l11a.State = s
		s = Controller.Lamp(12) : L12.State = s : l12a.State = s
		s = Controller.Lamp(13) : L13.State = s : l13a.State = s
		s = Controller.Lamp(14) : L14.State = s : l14a.State = s
		s = Controller.Lamp(15) : L15.State = s : l15a.State = s
		s = Controller.Lamp(16) : L16.State = s : l16a.State = s
		s = Controller.Lamp(17) : L17.State = s : l17a.State = s
		s = Controller.Lamp(18) : L18.State = s : l18a.State = s
		s = Controller.Lamp(19) : L19.State = s : l19a.State = s
		s = Controller.Lamp(20) : L20.State = s : l20a.State = s
		s = Controller.Lamp(21) : L21.State = s : l21a.State = s
		s = Controller.Lamp(22) : L22.State = s : l22a.State = s
		s = Controller.Lamp(23) : L23.State = s : l23a.State = s
		s = Controller.Lamp(24) : L24.State = s : l24a.State = s
		s = Controller.Lamp(25) : L25.State = s : l25a.State = s
		s = Controller.Lamp(26) : L26.State = s : l26a.State = s
		s = Controller.Lamp(27) : L27.State = s : l27a.State = s
		s = Controller.Lamp(28) : L28.State = s : l28a.State = s
		s = Controller.Lamp(29) : L29.State = s : l29a.State = s
		s = Controller.Lamp(30) : L30.State = s : l30a.State = s
		s = Controller.Lamp(31) : L31.State = s : l31a.State = s
		s = Controller.Lamp(32) : L32.State = s : l32a.State = s
		s = Controller.Lamp(33) : L33.State = s : l33a.State = s
		s = Controller.Lamp(34) : L34.State = s : l34a.State = s
		s = Controller.Lamp(35) : L35.State = s : l35a.State = s
		s = Controller.Lamp(36) : L36.State = s : l36a.State = s
		s = Controller.Lamp(37) : L37.State = s : l37a.State = s
		s = Controller.Lamp(38) : L38.State = s : l38a.State = s
		s = Controller.Lamp(39) : L39.State = s : l39a.State = s
		s = Controller.Lamp(40) : L40.State = s : l40a.State = s
		s = Controller.Lamp(41) : L41.State = s : l41a.State = s
		s = Controller.Lamp(42) : L42.State = s : l42a.State = s
		s = Controller.Lamp(43) : L43.State = s : l43a.State = s
		s = Controller.Lamp(44) : L44.State = s : l44a.State = s
		s = Controller.Lamp(45) : L45.State = s : l45a.State = s
		s = Controller.Lamp(46) : L46.State = s : l46a.State = s
		s = Controller.Lamp(47) : L47.State = s : l47a.State = s
		L48.State = Controller.Lamp(48)
		s = Controller.Lamp(49) : L49.State = s : l49a.State = s
		s = Controller.Lamp(50) : L50.State = s : l50a.State = s
		s = Controller.Lamp(51) : L51.State = s : l51a.State = s
		s = Controller.Lamp(52) : L52.State = s : L52a.State = s : L52b.State = s
		s = Controller.Lamp(53) : L53.State = s : L53a.State = s : L53b.State = s
		s = Controller.Lamp(54) : L54.State = s : L54a.State = s : L54b.State = s
		s = Controller.Lamp(55) : L55.State = s : L55b.state = s : L55b1.state = s : L55b2.state = s
		s = Controller.Lamp(56) : L56.State = s : L56b.state = s : L56b1.state = s
		s = Controller.Lamp(57) : L57.State = s : L57a.State = s
		s = Controller.Lamp(58) : L58.State = s : L58a.State = s
		s = Controller.Lamp(59) : L59.State = s : L59a.State = s
		s = Controller.Lamp(60) : L60.State = s : L60a.State = s
		s = Controller.Lamp(61) : L61.State = s : L61a.State = s
		s = Controller.Lamp(62) : L62.State = s : L62a.State = s
		s = Controller.Lamp(63) : L63.State = s : L63a.State = s
		s = Controller.Lamp(64) : L64.State = s : l64a.State = s
		L65.State = Controller.Lamp(65)
		s = Controller.Lamp(66) : L66.State = s : L66a.State = s
		s = Controller.Lamp(68) : L68.State = s : L68a.State = s
		s = Controller.Lamp(71) : L71.State = s : l71a.State = s
		L73.State = Controller.Lamp(73)
		L74.State = Controller.Lamp(74)
End Sub

'*******
'GI subs
'*******
Dim ig, ig2, ig3
Sub UpdateGITimer
	For each ig in GI
		If updateGI = 1 then
		ig.state = 1
		ElseIf UpdateGI = 2 Then
		ig.state = 2
		Else
		ig.state = 0
		End If
	Next
	For each ig2 in GI2
		If updateGI = 1 then
		ig2.state = 1
		ElseIf UpdateGI = 2 Then
		ig2.state = 2
		Else
		ig2.state = 0
		End If
	Next
	For each ig3 in GI3
		If updateGI = 1 then
		ig3.state = 1
		ElseIf UpdateGI = 2 Then
		ig3.state = 2
		Else
		ig3.state = 0
		End If
	Next
	GILite.enabled = 1
End Sub

Sub GILite_Timer
	UpdateGI = 1
	UpdateGITimer
	me.enabled = 0
End Sub

Sub GIFlash_Timer
	UpdateGI = 1
	UpdateGITimer
	me.enabled = 0
End Sub

'extra subs
' Last-state tracking for texture swap guards (eliminates redundant COM image writes)
Dim lastAlienState : lastAlienState = -99
Dim lastCabState : lastCabState = -99
Dim lastGate3Angle : lastGate3Angle = -9999
Dim lastGate4Angle : lastGate4Angle = -9999
Dim lastSw25Angle : lastSw25Angle = -9999
Dim lastSw27Angle : lastSw27Angle = -9999
Dim lastGate6Angle : lastGate6Angle = -9999

Sub GateTimer_Timer
	Dim ga
	ga = Gate3.currentangle
	If ga <> lastGate3Angle Then lastGate3Angle = ga : Gate3Prim.RotX = ga * -.85
	ga = Sw25.currentangle
	If ga <> lastSw25Angle Then lastSw25Angle = ga : Sw25prim.RotX = ga * -.85
	ga = Gate4.currentangle
	If ga <> lastGate4Angle Then lastGate4Angle = ga : Gate4Prim.RotX = ga * -.85
	ga = sw27.currentangle
	If ga <> lastSw27Angle Then lastSw27Angle = ga : Sw27prim.RotX = ga * -.85
	ga = Gate6.currentAngle
	If ga <> lastGate6Angle Then lastGate6Angle = ga : FileCabinetDoor.RotX = ga

	'Alien baby Texture swapping (guarded: only write when state changes)
	Dim s66, s68, alienCode
	s66 = L66.State : s68 = L68.State
	alienCode = s66 * 2 + s68  ' unique code per combination
	If alienCode <> lastAlienState Then
		lastAlienState = alienCode
		If s68 = -1 And s66 = 0 Then
			Primitive7.image = "GlassmapR"
			Primitive_AlienHead.image = "AlienHead_RRED_2"
			Primitive_AlienBody.image = "AlienBody_RRED2"
		ElseIf s66 = -1 And s68 = 0 Then
			Primitive7.image = "GlassmapL"
			Primitive_AlienHead.image = "AlienHead_LRED_2"
			Primitive_AlienBody.image = "AlienBody_LRED2"
		ElseIf s66 = -1 And s68 = -1 Then
			Primitive7.image = "GlassmapON"
			Primitive_AlienHead.image = "AlienHead_dualRED_2"
			Primitive_AlienBody.image = "AlienBody_DualRED2"
		Else
			Primitive7.image = "GlassMapOff"
			Primitive_AlienHead.image = "AlienHead_CompleteMapPale"
			Primitive_AlienBody.image = "AlienBody_CompleteMapPale"
		End If
	End If

	'File Cabinet Texture Swapping (guarded)
	Dim sf22, s55, s56, cabCode
	sf22 = F22.State : s55 = L55.State : s56 = L56.State
	cabCode = sf22 * 4 + s55 * 2 + s56
	If cabCode <> lastCabState Then
		lastCabState = cabCode
		If sf22 = 1 Then
			FileCabinet.image = "CabinetMap_FlasherON"
			FileCabinetDoor.image = "CabinetMap_FlasherON"
		ElseIf sf22 = 0 Then
			If s55 = -1 And s56 = 0 Then
				FileCabinet.Image = "CabinetMap_LON"
				FileCabinetDoor.Image = "CabinetMap_LON"
			ElseIf s55 = 0 And s56 = -1 Then
				FileCabinet.Image = "CabinetMap_RON"
				FileCabinetDoor.Image = "CabinetMap_RON"
			ElseIf s55 = -1 And s56 = -1 Then
				FileCabinet.Image = "CabinetMap_L+RON"
				FileCabinetDoor.Image = "CabinetMap_L+RON"
			Else
				FileCabinet.Image = "CabinetMap"
				FileCabinetDoor.Image = "CabinetMap"
			End If
		End If
	End If
End Sub

' Extra Sounds
Sub Metals_Hit(idx):PlaySound "metalhit":End Sub
Sub Gate1_Hit:PlaySound "gate":End Sub
Sub Gate2_Hit:PlaySound "gate":End Sub
Sub Gate3_Hit:PlaySound "gate":End Sub
Sub Gate4_Hit:PlaySound "gate":End Sub
Sub Gate5_Hit:PlaySound "gate":End Sub
Sub RHelp1_Hit:StopSound "metalrolling":PlaySound "BallHit":End Sub
Sub RHelp2_Hit:StopSound "metalrolling":PlaySound "BallHit":End Sub
Sub RHelp3_Hit:PlaySound "BallHit":End Sub
Sub RHelp4_Hit:PlaySound "BallHit":End Sub

Sub Rubber_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
'**********************************************************************
' Pre-computed inverse table dimensions
Dim InvTWHalf, InvTHHalf
InvTWHalf = 2 / table1.width
InvTHHalf = 2 / table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Dim bv : bv = BallVel(ball)
    Vol = Csng(bv * bv / 1000)
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
    BallVel = INT(SQR(vx * vx + vy * vy))
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************
Const tnob = 5 ' total number of balls
ReDim rolling(tnob)

' Pre-built rolling sound strings
ReDim BallRollStr(5)
Dim brsI : For brsI = 0 To 5 : BallRollStr(brsI) = "fx_ballrolling" & brsI : Next

InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ubBot
    BOT = GetBalls
    ubBot = UBound(BOT)

	' stop the sound of deleted balls
    For b = ubBot + 1 to tnob
        rolling(b) = False
        StopSound BallRollStr(b)
    Next

	' exit the sub if no balls on the table
    If ubBot = -1 Then Exit Sub

	' play the rolling sound for each ball
    Dim ball, bvx, bvy, bz, bv

    For b = 0 to ubBot
        Set ball = BOT(b)
        bvx = ball.VelX : bvy = ball.VelY : bz = ball.z
        bv = Int(Sqr(bvx * bvx + bvy * bvy))
        If bv > 1 Then
          If bz > 0 Then
            rolling(b) = True
            PlaySound BallRollStr(b), -1, Csng(bv * bv / 1000)*2, Pan(ball), 0, bv*20, 1, 0
          End If
        Else
            If rolling(b) = True Then
                StopSound BallRollStr(b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************
Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity * velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub


 Sub Augen_Timer

' ******Hanibals Random Lights Script

F22.Intensity = (35+(5*Rnd))
F22a.Intensity = F22.Intensity
F22b.Intensity = F22.Intensity
F22c.Intensity = F22.Intensity
F4.Intensity = (80+(5*Rnd))
F4a.Intensity = F4.Intensity
F5.Intensity = (80+(5*Rnd))
F5a.Intensity = F5.Intensity
Trappdoorlight.Intensity = (5+(2*Rnd))
L52a.Intensity = (30+(5*Rnd))
L53a.Intensity = L52a.Intensity
L54a.Intensity = L52a.Intensity
L52b.Intensity = L52a.Intensity/10
L53b.Intensity = L52a.Intensity/10
L54b.Intensity = L52a.Intensity/10

 End Sub


