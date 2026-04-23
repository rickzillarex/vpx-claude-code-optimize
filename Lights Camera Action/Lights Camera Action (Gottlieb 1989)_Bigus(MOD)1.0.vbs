'       ===================================================
'        Lights, Camera, Action by Gottlieb / Premier 1989
'       ===================================================
rem  Factory reset
'go to pinmametest press 0
'press ctrl and go to adjustments press 0
'press 0 again to load the factory reset
'reset rom with F3
Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'Const cGameName="lca" 'Original
Const cGameName="lca2" 'Rev 2

'----- FlexDMD Options -----
Dim UseFlexDMD:UseFlexDMD = 1		' 1 = on, 0 = off .Intended for Real DMD users but will work on LCD
Const FlexColour = 2			' 0 = default, 1 = cyan, 2 = red, 3 = yellow, 4 = Green, 5 = blue, 6 = white, 7 = magenta


Dim VarRol, VarHidden, HideLoop
VarHidden=1
If Table1.ShowDT = true then
	VarRol=0
	For Each HideLoop in fullscreenhide
		HideLoop.Visible = True
	Next
	Primitive20.Visible=True:Primitive21.Visible=True:Primitive31.Visible=True:Primitive32.Visible=True
Else
	VarRol=1
	For Each HideLoop in fullscreenhide
		HideLoop.Visible = False
	Next
	Primitive20.Visible=False:Primitive21.Visible=False:Primitive31.Visible=False:Primitive32.Visible=False:Primitive25.Visible=False
End If

If B2SOn = true Then VarHidden=1

LoadVPM "01120100","GTS3.VBS",3.02

Const UseSolenoids=2,UseLamps=1,UseSync=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"

'*************************************
' PRE-COMPUTED CONSTANTS (optimization)
'*************************************
Const tnob = 5 ' total number of balls

' Pre-built string arrays for rolling sounds (eliminates per-tick string concatenation)
ReDim BallRollStr(5)
ReDim BallDropStr(5)
Dim brsI
For brsI = 0 To tnob
    BallRollStr(brsI) = "fx_ballrolling" & brsI
    BallDropStr(brsI) = "fx_ball_drop" & brsI
Next

' Pre-built FlexDMD segment name strings (eliminates per-LED string concat)
ReDim SegStr(39)
Dim segI
For segI = 0 To 39
    SegStr(segI) = "Seg" & segI
Next

' Table dimension inverses (assigned in Table1_Init since COM props not available yet)
Dim InvTWHalf, InvTHHalf

' Const for volume calc
Dim RollVol
RollVol = 1 ' default, may be overridden

SolCallback(sLLFlipper)="SolFlipper LeftFlipper,Flipper1,"
SolCallback(sLRFlipper)="SolFlipper RightFlipper,Nothing,"
SolCallback(1)="vpmSolSound ""Jet3"","
SolCallback(2)="vpmSolSound ""Jet3"","
SolCallback(3)="vpmSolSound ""Jet3"","
SolCallback(4)="vpmSolSound ""Sling"","
SolCallback(5)="vpmSolSound ""Sling"","
SolCallback(6)="vpmSolAutoPlunger Plunger1,0,"
SolCallback(7)="dtTDrop.SolDropUp"
'SolCallback(8)="dtTDrop.SolDropDown"
SolCallback(9)="goodguy"'Left Gunfighter
SolCallback(10)="badguy"'Right Gunfighter
SolCallback(11)="bsTLHole.SolOut"
SolCallback(12)="bsTRHole.SolOut"
SolCallback(13)="bsBLHole.SolOut" 'changed destruk
SolCallback(14)="bsBRHole.SolOut"
SolCallback(15)="dtRDrop.SolDropDown"
SolCallback(16)="dtRDrop.SolDropUp"
'SolCallback(17)="vpmflasher light29,"'Playboard Release - Add MNPG
SolCallback(17)="PlayState"'Playboard Release - Add MNPG and Playfield middle left flasher
SolCallback(18)="flasher18" 'Playfield lower left flashers
SolCallback(19)="flasher19" 'Playfield lower right flashers
SolCallback(20)="flasher20" 'Playfield upper right flashers
'SolCallback(21)=Loop
SolCallback(22)="vpmflasher light8,"'Left Gunfighter (Police) - Add MNPG
SolCallback(23)="vpmFlasher light18,"'Right Gunfighter (Outlaw) - Add MNPG
SolCallback(24)="vpmflasher light28,"'Flood Light Blue (Backglass Right) - Add MNPG
SolCallback(25)="vpmflasher light19,"'Flood Light Red (Backglass Left) - Add MNPG
SolCallback(26)="vpmflasher light9,"'Motor - Add MNPG
'SolCallback(27)=Lightbox Illumination
SolCallback(28)="bsTrough.SolOut"
SolCallback(29)="bsTrough.SolIn"
SolCallback(30)="vpmSolSound ""Knocker"","
SolCallback(31)="pfgilights"'Playfield GI
'SolCallback(32)=Switched +48v DC

Sub flasher18(Enabled)
	If Enabled then Flasher10.state=1:Flasher11.state=1 Else Flasher10.state=0:Flasher11.state=0
End Sub

Sub flasher19(Enabled)
	If Enabled then Flasher12.state=1:Flasher13.state=1 Else Flasher12.state=0:Flasher13.state=0
End Sub

Sub flasher20(Enabled)
	If Enabled then Flasher14.state=1:Flasher15.state=1:Flasher16.state=1 Else Flasher14.state=0:Flasher15.state=0:Flasher16.state=0
End Sub

Dim gil
Sub pfgilights(Enabled)
	For Each gil in gilights
		gil.State = not Enabled
	Next
End Sub

Sub badguy(Enabled)
	If Enabled then gunupdownr=1 Else gunupdownr=-1
	badguyt.enabled=true
End Sub

Sub goodguy(Enabled)
	If Enabled then gunupdownl=-1 Else gunupdownl=1
	goodguyt.enabled=true
End Sub

'Sub SolBLHole(Enabled) 'changed Destruk - prevents more than one ball at a time in bottom left popper
'	If Enabled Then
'		'Kicker4.DestroyBall
'		bsBLHole.ExitSol_On
'	End If
'End Sub

Dim bsTrough,dtTDrop,dtRDrop
Dim bsTLHole,bsTRHole,bsBLHole,bsBRHole
Sub Table1_Init
	FlexDMD_Init
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Lights, Camera, Action; Premier 1989." & vbNewLine & "by Kristian, Mnpg & Destruk"
		.HandleKeyboard=0
		.ShowTitle=0
		.ShowDMDOnly=1
		.HandleMechanics=0
		.ShowFrame=0
		.Hidden = VarHidden
		.Games(cGameName).Settings.Value("rol")=VarRol
		.DIP(0)=&H00
		If UseFlexDMD Then ExternalEnabled = .Games(cGameName).Settings.Value("showpindmd")
		If UseFlexDMD Then .Games(cGameName).Settings.Value("showpindmd") = 0
		On Error Resume Next
		.Run
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=151
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper10,LSling,RSling)

	vpmMapLights clights

	' Pre-compute table dimension inverses (COM props only available after init)
	InvTWHalf = 2 / Table1.Width
	InvTHHalf = 2 / Table1.Height

'Add MNPG
	'Wall65.ISDropped=0
	Wall62.IsDropped=1:rampflip.enabled = true
	Light29.state=lightstateon
	Ramp.IsDropped=0

	Set bsTrough = New cvpmBallStack
	bsTrough.InitSw 33,34,0,0,0,0,0,0
	bsTrough.InitKick BallRelease, 90, 5
	bsTrough.InitEntrySnd "Outhole", "Solenoid"
	bsTrough.InitExitSnd "BallRel", "Solenoid"
    bsTrough.Balls=2

    set bsTLHole = new cvpmBallStack
 	bsTLHole.InitSaucer Kicker2,21,180,1 'changed to saucer ballstack to correct instant multiball - destruk
	bsTLHole.InitEntrySnd "Solenoid", "Solenoid"
	bsTLHole.InitExitSnd "popper", "Solenoid"
    bsTLHole.balls = 0

	set bsTRHole = new cvpmBallStack
 	'bsTRHole.InitSw 0,31,0,0,0,0,0,0
	bsTRHole.KickZ = 1.56
 	bsTRHole.InitSaucer Kicker3,31,0,23
 	'bsTRHole.InitKick Kicker8,220,10 ' Add MNPG
	bsTRHole.InitEntrySnd "Solenoid", "Solenoid"
	bsTRHole.InitExitSnd "popper", "Solenoid"

    set bsBLHole = new cvpmBallStack
	'bsBLHole.InitSw 0,41,0,0,0,0,0,0
	bsBLHole.KickZ = 1.56
 	bsBLHole.InitSaucer Kicker4,41,0,25 ' Add MNPG
	bsBLHole.InitEntrySnd "Solenoid", "Solenoid"
	bsBLHole.InitExitSnd "popper", "Solenoid"

    set bsBRHole = new cvpmBallStack
 	'bsBRHole.InitSw 0,51,0,0,0,0,0,0
	bsBRHole.KickZ = 1.56
 	bsBRHole.InitSaucer Kicker5,51,0,23
 	'bsBRHole.InitKick Kicker7,210,10 ' Add MNPG
	bsBRHole.InitEntrySnd "Solenoid", "Solenoid"
	bsBRHole.InitExitSnd "popper", "Solenoid"

   Set dtTDrop = New cvpmDropTarget
   dtTDrop.InitDrop Array(Wall1,Wall2,Wall3),Array(6,16,26)
   dtTDrop.InitSnd "flapclos","flapopen"
   dtTDrop.CreateEvents "dtTDrop"

   Set dtRDrop = New cvpmDropTarget
   dtRDrop.InitDrop Array(Wall4,Wall6,Wall7,Wall8,Wall10,Wall11),Array(7,17,27,37,47,57)
   dtRDrop.InitSnd "flapclos","flapopen"
   dtRDrop.CreateEvents "dtRDrop"

	Kicker1.CreateBall
	Kicker1.Kick 135,4
	Plunger1.Pullback

End Sub

Const swStartButton=3

'Destruk forgot the high score intials code so I added it back -Kristian
Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode=LeftFlipperKey Then Controller.Switch(60)=1
	If KeyCode=RightFlipperKey Then Controller.Switch(61)=1
	If vpmKeyDown(KeyCode) Then Exit Sub
	If KeyCode=PlungerKey Then Plunger.Pullback
    If keycode=LeftFlipperKey then Controller.Switch(4)=1
    If keycode=Rightflipperkey then controller.switch(5)=1
'    if Keycode = LeftMagnaSave then msgbox Join(LEDMAPPING, vbCrlf)
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
'if keycode=30 then dttdrop.hit 1
'if keycode=31 then dttdrop.hit 2
'if keycode=32 then dttdrop.hit 3
	If KeyCode=LeftFlipperKey Then Controller.Switch(60)=0
	If KeyCode=RightFlipperKey Then Controller.Switch(61)=0
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode = PlungerKey Then PlaySound "Plunger" : Plunger.Fire
	If keycode=LeftFlipperKey then Controller.Switch(4)=0
    If keycode=Rightflipperkey then controller.switch(5)=0
End Sub

'4 Left Advance
'5 Right Advance
'Sub Wall1_Hit:dtTDrop.Hit 1:End Sub						'6
'Sub Wall4_Hit:dtRDrop.Hit 1:End Sub						'7
Sub Bumper1_Hit:vpmTimer.PulseSw 10:End Sub				'10
Sub Bumper2_Hit:vpmTimer.PulseSw 11:End Sub				'11
Sub Bumper10_Hit:vpmTimer.PulseSw 12:End Sub			'12
Sub LSling_Slingshot:vpmTimer.PulseSw 13:End Sub		'13
Sub RSling_Slingshot:vpmTimer.PulseSw 14:End Sub		'14
Sub LeftOutlane_Hit:Controller.Switch(15)=1:End Sub		'15
Sub LeftOutlane_unHit:Controller.Switch(15)=0:End Sub
'Sub Wall2_Hit:dtTDrop.Hit 2:End Sub						'16
'Sub Wall6_Hit:dtRDrop.Hit 2:End Sub						'17
Sub Wall16_Hit:vpmTimer.PulseSw 20:Primitive57.TransZ=8:sttimer.enabled = true:End Sub				'20
Sub Kicker2_Hit:bsTLHole.AddBall 0:End Sub			    '21
Sub Wall15_Hit:vpmTimer.PulseSw 22:End Sub				'22
Sub LeftInlane_Hit:Controller.Switch(23)=1:End Sub		'23
Sub LeftInlane_unHit:Controller.Switch(23)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(24)=1:End Sub		'24
Sub RightInlane_unHit:Controller.Switch(24)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(25)=1:End Sub	'25
Sub RightOutlane_unHit:Controller.Switch(25)=0:End Sub
'Sub Wall3_Hit:dtTDrop.Hit 3:End Sub						'26
'Sub Wall7_Hit:dtRDrop.Hit 3:End Sub						'27
Sub Wall17_Hit:vpmTimer.PulseSw 30:Primitive58.TransZ=8:sttimer.enabled = true:End Sub				'30
Sub Kicker3_Hit:bsTRHole.AddBall 0:End Sub				'31 - Add MNPG Kicker3.destroyball:
Sub Wall20_Hit:vpmTimer.PulseSw 32:End Sub				'32
Sub Drain_Hit:bsTrough.AddBall Me:PlaySound "fx_drain":End Sub				'33/34
Sub Trigger1_Hit:Controller.Switch(35)=1:End Sub		'35
Sub Trigger1_unHit:Controller.Switch(35)=0:End Sub
Sub Wall12_Hit:vpmTimer.PulseSw 36:End Sub				'36
'Sub Wall8_Hit:dtRDrop.Hit 4:End Sub						'37
Sub Wall18_Hit:vpmTimer.PulseSw 40:Primitive59.TransZ=8:sttimer.enabled = true:End Sub				'40
Sub Kicker4_Hit:bsBLHole.AddBall 0:End Sub				'41 - Add MNPG/changed Destruk
Sub Spinner1_Spin:vpmTimer.PulseSw 42: End Sub			'42
Sub Trigger3_Hit:Controller.Switch(45)=1:End Sub		'45
Sub Trigger3_unHit:Controller.Switch(45)=0:End Sub
Sub Wall13_Hit:vpmTimer.PulseSw 46:End Sub				'46
'Sub Wall10_Hit:dtRDrop.Hit 5:End Sub					'47
Sub Wall19_Hit:vpmTimer.PulseSw 50:Primitive60.TransZ=8:sttimer.enabled = true:End Sub				'50
Sub Kicker5_Hit:bsBRHole.AddBall 0:End Sub				'51 - Add MNPG Kicker5.destroyball:
Sub Trigger2_Hit:Controller.Switch(52)=1:End Sub		'52
Sub Trigger2_unHit:Controller.Switch(52)=0:End Sub
Sub Wall14_Hit:vpmTimer.PulseSw 56:End Sub				'56
'Sub Wall11_Hit:dtRDrop.Hit 6:End Sub					'57
'60 Left Flippers(2)
'61 Right Flipper

'MNPG Adds
Dim stat
stat=0
Sub PlayState(Enabled)
	If Enabled Then
		If Stat=1 Then
		'Wall65.ISDropped=0
		Wall62.IsDropped=1:rampflip.enabled = true
		Light29.state=lightstateon
		Ramp.IsDropped=0
		stat=0
		Else
		'Wall65.ISDropped=1
		Wall62.IsDropped=0:rampflip.enabled = true
		Light29.state=lightstateoff
		Ramp.IsDropped=1
		stat=1
		End if
	End If
End Sub



Sub TShowKicker_Hit:Kicker5.Enabled=True:End Sub		'52
Sub THideKicker_Hit:Kicker5.Enabled=False:End Sub


Sub Gate1_Hit()PlaySound "Gate"
End Sub
Sub Gate2_Hit()PlaySound "Gate"
End Sub
Sub Trigger4_Hit()PlaySound "PlungeBall"
End Sub

' Flipper timer: cache CurrentAngle + delta guards
Dim lastLFAngle, lastLF1Angle, lastRFAngle
lastLFAngle = -9999 : lastLF1Angle = -9999 : lastRFAngle = -9999

Sub flippers_Timer()
    Dim aLF, aLF1, aRF
    aLF = LeftFlipper.CurrentAngle
    aLF1 = Flipper1.CurrentAngle
    aRF = RightFlipper.CurrentAngle
    If aLF <> lastLFAngle Then
        lastLFAngle = aLF
        LeftFlipperP.objRotZ = aLF - 90
    End If
    If aLF1 <> lastLF1Angle Then
        lastLF1Angle = aLF1
        LeftFlipperP1.objRotZ = aLF1 - 90
    End If
    If aRF <> lastRFAngle Then
        lastRFAngle = aRF
        RightFlipperP.objRotZ = aRF - 90
    End If
End Sub

Dim rotposp:rotposp=0
Sub rampflip_Timer()
	rotposp=rotposp-1
	Dim rz : rz = rotposp   ' cache to avoid 5 identical writes when value unchanged
	Primitive87.RotZ=rz:Primitive88.RotZ=rz:Primitive89.RotZ=rz:Primitive90.RotZ=rz:Primitive91.RotZ=rz
	If rotposp=-360 then rotposp=0
	If Wall62.IsDropped = True and rotposp = -180 then rampflip.enabled = false
	If Wall62.IsDropped = False and rotposp = 0 then rampflip.enabled = false
End Sub

' targettimer: cache light states and guard flasher writes with delta checks
ReDim lastTgtVis(8)
Dim tgtI
For tgtI = 0 To 8 : lastTgtVis(tgtI) = -1 : Next

Sub targettimer_Timer()
    Dim s45, s46, s47, s27, s34, s33, s32, s31, s30
    s45 = Light45.State : s46 = Light46.State : s47 = Light47.State : s27 = Light27.State
    s34 = Light34.State : s33 = Light33.State : s32 = Light32.State : s31 = Light31.State : s30 = Light30.State

    Dim v
    v = (s45 = 1) : If v <> lastTgtVis(0) Then lastTgtVis(0) = v : Flasher1.Visible = v
    v = (s46 = 1) : If v <> lastTgtVis(1) Then lastTgtVis(1) = v : Flasher2.Visible = v
    v = (s47 = 1) : If v <> lastTgtVis(2) Then lastTgtVis(2) = v : Flasher4.Visible = v
    v = (s27 = 1) : If v <> lastTgtVis(3) Then lastTgtVis(3) = v : Flasher3.Visible = v

    If s34 = 1 Then v = 100 Else v = 0
    If v <> lastTgtVis(4) Then lastTgtVis(4) = v : Flasher5.Amount = v
    If s33 = 1 Then v = 100 Else v = 0
    If v <> lastTgtVis(5) Then lastTgtVis(5) = v : Flasher6.Amount = v
    If s32 = 1 Then v = 100 Else v = 0
    If v <> lastTgtVis(6) Then lastTgtVis(6) = v : Flasher7.Amount = v
    If s31 = 1 Then v = 100 Else v = 0
    If v <> lastTgtVis(7) Then lastTgtVis(7) = v : Flasher8.Amount = v
    If s30 = 1 Then v = 100 Else v = 0
    If v <> lastTgtVis(8) Then lastTgtVis(8) = v : Flasher9.Amount = v
End Sub

Sub sttimer_Timer()
	Dim stloop
	For Each stloop in standuptargets
		stloop.TransZ=0
	Next
	sttimer.enabled = false
End Sub

Dim gunrotr:gunrotr=-90
Dim gunupdownr
Sub badguyt_Timer()
	gunrotr=gunrotr+gunupdownr
	Primitive20.RotY=gunrotr
	If gunrotr = 0 or gunrotr = -90 then badguyt.enabled = false
End Sub

Dim gunrotl:gunrotl=90
Dim gunupdownl
Sub goodguyt_Timer()
	gunrotl=gunrotl+gunupdownl
	Primitive21.RotY=gunrotl
	If gunrotl = 0 or gunrotl = 90 then goodguyt.enabled = false
End Sub

Sub Trigger5_Hit()
	Kicker2.Kick 180,1
    Controller.Switch(21)=0
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * InvTWHalf - 1
  If tmp > 0 Then
    Dim t2,t4,t8
    t2 = tmp*tmp : t4 = t2*t2 : t8 = t4*t4
    Pan = Csng(t8 * t2)
  Else
    Dim nt,n2,n4,n8
    nt = -tmp : n2 = nt*nt : n4 = n2*n2 : n8 = n4*n4
    Pan = Csng(-(n8 * n2))
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Dim bv : bv = BallVel(ball)
  Vol = Csng(bv * bv / 500)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  Dim bv : bv = BallVel(ball)
  VolMulti = Csng(bv * bv / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  Dim bv : bv = BallVel(ball)
  DVolMulti = Csng(bv * bv / 150 ) * Multiplier
  ' debug.print removed (COM string alloc per call)
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  Dim bv : bv = BallVel(ball)
  BallRollVol = Csng(bv * bv / (80000 - (79900 * Log(RollVol) / Log(100))))
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

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
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

	' play the rolling sound for each ball
    Dim ball, bvx, bvy, bv, bz, bvz

    For b = 0 to UBound(BOT)
      Set ball = BOT(b)
      ' Inline BallVel: cache COM reads once
      bvx = ball.VelX : bvy = ball.VelY
      bv = Int(Sqr(bvx * bvx + bvy * bvy))
      bz = ball.z

      If bv > 1 Then
        rolling(b) = True
        if bz < 30 Then ' Ball on playfield
          PlaySound BallRollStr(b), -1, Csng(bv*bv/500)*.03, AudioPan(ball), 0, bv*20, 1, 0, AudioFade(ball)
        Else ' Ball on raised ramp
          PlaySound BallRollStr(b), -1, Csng(bv*bv/500)*.02, AudioPan(ball), 0, bv*20+50000, 1, 0, AudioFade(ball)
        End If
      Else
        If rolling(b) = True Then
          StopSound BallRollStr(b)
          rolling(b) = False
        End If
      End If
 ' play ball drop sounds
      bvz = ball.VelZ
      If bvz < -1 and bz < 55 and bz > 27 Then 'height adjust for ball drop sounds
          PlaySound BallDropStr(b), 0, ABS(bvz)/17, AudioPan(ball), 0, bv*20, 1, 0, AudioFade(ball)
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	Dim v2 : v2 = velocity * velocity
	PlaySound("fx_collide"), 0, Csng(v2) / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub BallShadowUpdate_Timer()
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6)
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
		BallShadow(b).X = bx
		ballShadow(b).Y = by + 10
        If bz > 20 and bz < 200 Then newVis = 1 Else newVis = 0
        If BallShadow(b).visible <> newVis Then BallShadow(b).visible = newVis
        If bz > 30 Then
            ballShadow(b).height = bz - 20
        Else
            ballShadow(b).height = bz - 24
        End If
        ballShadow(b).opacity = 90
    Next
    if light68.state=0 then table1.ColorGradeImage = "colorgrade_4" else	table1.ColorGradeImage = "colorgrade_8":end if
End Sub

 '**********************
'Flipper Shadows
'***********************
Dim lastLFS, lastRFS
lastLFS = -9999 : lastRFS = -9999

Sub RealTime_Timer
  Dim aLF, aRF
  aLF = LeftFlipper.CurrentAngle
  aRF = RightFlipper.CurrentAngle
  If aLF <> lastLFS Then lastLFS = aLF : lfs.RotZ = aLF
  If aRF <> lastRFS Then lastRFS = aRF : rfs.RotZ = aRF
End Sub

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


Sub Gates_Hit (idx)
	PlaySound "gate", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub


Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub




' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
  If UseFlexDMD then
    If Not FlexDMD is Nothing Then
      FlexDMD.Show = False
      FlexDMD.Run = False
      FlexDMD = NULL
    End if
    Controller.Games(cGameName).Settings.Value("showpindmd") = ExternalEnabled
  End if
End Sub

ReDim LEDMAPPING(40)
Sub FlexTimer_Timer

		Dim ChgLED, ii, num, chgstat
		ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)

		If Not IsEmpty(ChgLED)Then

			FlexDMD.LockRenderThread

			For ii=0 To UBound(chgLED)
				num=chgLED(ii, 0) : chgstat=chgLED(ii, 2)
				UpdateFlexChar num, chgstat
				LEDMAPPING(num) = num & vbtab & chgstat

			Next

			FlexDMD.UnlockRenderThread

		End if

End Sub


'**********************************************************************************************************
' FlexDMD code - scutters
'**********************************************************************************************************
Dim FlexDMD
DIm FlexDMDDict
Dim FlexDMDScene
Dim ExternalEnabled

' Pre-resolved FlexDMD image name cache (eliminates Dictionary lookup + NewImage alloc per LED update)
Dim FlexDMDImageCache
Dim FlexDMDSpaceImage

Sub FlexDMD_Init() 'default/startup values

	'setup flex dmd

	If UseFlexDMD = 0 then
		Set FlexDMD = nothing
		exit sub
	end if

	Dim i

	' populate the lookup dictionary for mapping display characters
	FlexDictionary_Init

	Set FlexDMD = CreateObject("FlexDMD.FlexDMD")

	If Not FlexDMD is Nothing Then

		FlexDMD.GameName = cGameName
 		FlexDMD.TableFile = Table1.Filename & ".vpx"
		FlexDMD.RenderMode = 2
		FlexDMD.Width = 128
		FlexDMD.Height = 32
		FlexDMD.Clear = True
		FlexDMD.Run = True

		Set FlexDMDScene = FlexDMD.NewGroup("Scene")

		With FlexDMDScene
			'populate blank display
			.AddActor FlexDMD.NewImage("BackG", "FlexDMD.Resources.dmds.black.png")

			.AddActor FlexDMD.NewFrame("Frame")
			.GetFrame("Frame").Visible = True
			Select Case FlexColour
			Case 1
				.GetFrame("Frame").FillColor = vbYellow
				.GetFrame("Frame").BorderColor = vbYellow
			Case 2
				.GetFrame("Frame").FillColor = vbBlue
				.GetFrame("Frame").BorderColor = vbBlue
			Case 3
				.GetFrame("Frame").FillColor = vbCyan
				.GetFrame("Frame").BorderColor = vbCyan
			Case 4
				.GetFrame("Frame").FillColor = vbGreen
				.GetFrame("Frame").BorderColor = vbGreen
			Case 5
				.GetFrame("Frame").FillColor = vbRed
				.GetFrame("Frame").BorderColor = vbRed
			Case 6
				.GetFrame("Frame").FillColor = vbWhite
				.GetFrame("Frame").BorderColor = vbWhite
			Case 7
				.GetFrame("Frame").FillColor = RGB(178,0,220)
				.GetFrame("Frame").BorderColor = RGB(178,0,220)
			Case Else
				.GetFrame("Frame").FillColor = RGB(255,128,0)
				.GetFrame("Frame").BorderColor = RGB(255,128,0)
			End Select
			.GetFrame("Frame").Height = 32
			.GetFrame("Frame").Width= 128
			.GetFrame("Frame").Fill= True
			.GetFrame("Frame").Thickness= 1

			.AddActor FlexDMD.NewImage("Back", "VPX.DMD_Background")

			'40 segment display holders
			for i = 0 to 19
				'first line
				.AddActor FlexDMD.NewImage(SegStr(i), "VPX.DMD_Space")
				.GetImage(SegStr(i)).SetAlignedPosition 4 + i * 6,0,0
				'second line
				.AddActor FlexDMD.NewImage(SegStr(i+20), "VPX.DMD_Space")
				.GetImage(SegStr(i+20)).SetAlignedPosition 4 + i * 6,16,0
			next

		End With

		FlexDMD.LockRenderThread
		FlexDMD.Stage.AddActor FlexDMDScene

		FlexDMD.Show = True
		FlexDMD.UnlockRenderThread

		FlexTimer.Enabled = 1

		' Pre-build image cache: resolve all dictionary values to actual FlexDMD image objects
		Set FlexDMDImageCache = CreateObject("Scripting.Dictionary")
		Set FlexDMDSpaceImage = FlexDMD.NewImage("_space_", "VPX.DMD_Space")
		Dim k, keys
		keys = FlexDMDDict.Keys()
		For Each k In keys
			Set FlexDMDImageCache(k) = FlexDMD.NewImage("_cache_" & k, FlexDMDDict.Item(k))
		Next

	Else

		UseFlexDMD = 0

	End If

End Sub

Sub FlexDictionary_Init

	Set FlexDMDDict = CreateObject("Scripting.Dictionary")

	FlexDMDDict.Add 0, "VPX.DMD_Space"
	FlexDMDDict.Add 63, "VPX.DMD_O"
	FlexDMDDict.Add 8704, "VPX.DMD_1"
	FlexDMDDict.Add 2139, "VPX.DMD_2"
	FlexDMDDict.Add 2127, "VPX.DMD_3"
	FlexDMDDict.Add 2150, "VPX.DMD_4"
	FlexDMDDict.Add 2157, "VPX.DMD_S"
	FlexDMDDict.Add 2173, "VPX.DMD_6"
	FlexDMDDict.Add 7, "VPX.DMD_7"
	FlexDMDDict.Add 2175,"VPX.DMD_8"
	FlexDMDDict.Add 2159,"VPX.DMD_9"

	FlexDMDDict.Add 32959,"VPX.DMD_Odot"
	FlexDMDDict.Add 41600, "VPX.DMD_1dot"
	FlexDMDDict.Add 35035, "VPX.DMD_2dot"
	FlexDMDDict.Add 35023, "VPX.DMD_3dot"
	FlexDMDDict.Add 35046, "VPX.DMD_4dot"
	FlexDMDDict.Add 35053, "VPX.DMD_Sdot"
	FlexDMDDict.Add 35069, "VPX.DMD_6dot"
	FlexDMDDict.Add 32903, "VPX.DMD_7dot"
	FlexDMDDict.Add 35071, "VPX.DMD_8dot"
	FlexDMDDict.Add 35055, "VPX.DMD_9dot"


	FlexDMDDict.Add 2167, "VPX.DMD_A"
	FlexDMDDict.Add 10767, "VPX.DMD_B"
	FlexDMDDict.Add 57, "VPX.DMD_C"
	FlexDMDDict.Add 8719, "VPX.DMD_D"
	FlexDMDDict.Add 121, "VPX.DMD_E"
	FlexDMDDict.Add 113, "VPX.DMD_F"
	FlexDMDDict.Add 2109, "VPX.DMD_G"
	FlexDMDDict.Add 2166, "VPX.DMD_H"
	FlexDMDDict.Add 8713, "VPX.DMD_I"
	FlexDMDDict.Add 30, "VPX.DMD_J"
	FlexDMDDict.Add 5232, "VPX.DMD_K"
	FlexDMDDict.Add 56, "VPX.DMD_L"
	FlexDMDDict.Add 1334, "VPX.DMD_M"
	FlexDMDDict.Add 4406, "VPX.DMD_N"
        ' "O" = 0
	FlexDMDDict.Add 2163, "VPX.DMD_P"
	FlexDMDDict.Add 4159, "VPX.DMD_Q"
	FlexDMDDict.Add 6259, "VPX.DMD_R"
        ' "S" = 5
	FlexDMDDict.Add 8705, "VPX.DMD_T"
	FlexDMDDict.Add 62, "VPX.DMD_U"
	FlexDMDDict.Add 17456, "VPX.DMD_V"
	FlexDMDDict.Add 20534, "VPX.DMD_W"
	FlexDMDDict.Add 21760, "VPX.DMD_X"
	FlexDMDDict.Add 9472, "VPX.DMD_Y"
	FlexDMDDict.Add 17417, "VPX.DMD_Z"

	FlexDMDDict.Add &h400,"VPX.DMD_SingleQuote"
	FlexDMDDict.Add 16640, "VPX.DMD_CloseBracket"
	FlexDMDDict.Add 5120, "VPX.DMD_OpenBracket"
	FlexDMDDict.Add 2120, "VPX.DMD_Equals"
	FlexDMDDict.Add 10275, "VPX.DMD_Question"
	FlexDMDDict.Add 2112, "VPX.DMD_Minus"
	FlexDMDDict.Add 10861, "VPX.DMD_Dollar"
	FlexDMDDict.Add 6144, "VPX.DMD_GreaterThan"
	FlexDMDDict.Add 65535, "VPX.DMD_Hash"
	FlexDMDDict.Add 32576, "VPX.DMD_Asterick"
	FlexDMDDict.Add 10816, "VPX.DMD_Plus"
	FlexDMDDict.Add 16688, "VPX.DMD_ArrowLeft"
	FlexDMDDict.Add 5126, "VPX.DMD_ArrowRight"
'	FlexDMDDict.Add 1, "VPX.DMD_UpperScore"
'	FlexDMDDict.Add 9, "VPX.DMD_2Stripes"
'	FlexDMDDict.Add 8, "VPX.DMD_LowerScore"

End sub

Sub UpdateFlexChar(id, value)

	If id < 40 Then
		Dim cachedImg
		if FlexDMDImageCache.Exists(value) then
			Set cachedImg = FlexDMDImageCache.Item(value)
		Else
			Set cachedImg = FlexDMDSpaceImage
		end if
		FlexDMDScene.GetImage(SegStr(id)).Bitmap = cachedImg.Bitmap
	End If

End Sub

'**********************************************************************************************************
