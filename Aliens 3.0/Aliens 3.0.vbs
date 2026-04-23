' ==========================================================================
' VPX PERFORMANCE OPTIMIZATION PATCH
' Applied optimizations:
'   #3 AudioFade: replaced ^10 with chain multiplication
'   #6 AudioFade: cached Table1.height -> TableHeight
'   #3 Pan: replaced ^10 with chain multiplication
'   #6 Pan: cached table1.width -> TableWidth
'   #4 BallVel: replaced ^2 with x*x
'   #4 Vol: eliminated SQR+^2, use VelX*VelX+VelY*VelY directly
'   #4 OnBallBallCollision: replaced velocity ^2 with velocity*velocity
'   #6 Added TableWidth/TableHeight cache from Table1.width/height
'   #7 Pre-built BallRollStr() string array
' ==========================================================================
' VPX Aliens by Delta23 2023, version 3.0 - A heavy mod of JP's Sorcerer
' VPX Sorcerer by JPSalas 2018, version 1.0.0

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1.3

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "S11.VBS", 3.26
Dim bsTrough, bsSaucer, dtBank, x, a, liarray

'Const cGameName = "sorcr_l1"
Const cGameName = "sorcr_l2"

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'************************
' Switch - Music Control
'************************

Dim MusicCount, MusicEnd
Sub Music_Hit()
    If MusicCount = 0 then PlayMusic "A-TRACK 1.mp3": End If
    If MusicCount = 1 then PlayMusic "A-TRACK 2.mp3": End If
	If MusicCount = 2 then PlayMusic "A-TRACK 3.mp3": End If
    If MusicCount = 3 then PlayMusic "A-TRACK 4.mp3": End If
	If MusicCount = 4 then PlayMusic "A-TRACK 5.mp3": End If
	If MusicCount = 5 then PlayMusic "A-TRACK 6.mp3": End If
	If MusicCount = 6 then PlayMusic "A-TRACK 7.mp3": End If
	If MusicCount = 7 then PlayMusic "A-TRACK 8.mp3": End If
	MusicCount = Int(rnd*8)
End Sub

Sub Table1_MusicDone
    Music_Hit
End Sub

'***************************
' Switches - Sound Efffects
'***************************

Dim FX4Count
Sub SFX4_Hit()
	FX4Count=Int(rnd*4)
	Controlli1.state = 0
	Controlli4.state = 0
	If FX4Count = 0 then PlaySound "Vents 3": End If
	If FX4Count = 1 then PlaySound "Alien 9": End If
	If FX4Count = 2 then PlaySound "Vents 1": End If
	If FX4Count = 3 then PlaySound "Alien 5": End If
End Sub

Sub SFX5_Hit()
	ScannerON
	Flasher6ON
End Sub

Sub SFX6_Hit()
	ScannerOFF
End Sub 

Dim FX7Count
Sub SFX7_Hit() StopSound "Combat Drop"
	Controlli2.state = 0
	DataLoadOFF
	StopSFXb
	If FX7Count = 0 then PlaySound "Sgt Apone 1": End If
	If FX7Count = 1 then PlaySound "Hudson 7": End If
	If FX7Count = 2 then PlaySound "Ripley 13": End If		
	If FX7Count = 3 then PlaySound "Bishop 1": End If	
	If FX7Count = 4 then PlaySound "Sgt Apone 2": End If
	If FX7Count = 5 then PlaySound "Ripley 24": End If
	If FX7Count = 6 then PlaySound "Sgt Apone 32": End If
	If FX7Count = 7 then PlaySound "Sgt Apone 6": End If
	If FX7Count = 8 then PlaySound "Burke 3": End If
	If FX7Count = 9 then PlaySound "Sgt Apone 30": End If
	If FX7Count = 10 then PlaySound "Gorman 4": End If
	If FX7Count = 11 then PlaySound "Ripley 14": End If
	If FX7Count = 12 then PlaySound "Sgt Apone 29": End If
	If FX7Count = 13 then PlaySound "Sgt Apone 24": End If
	If FX7Count = 14 then PlaySound "Random 34": End If
	If FX7Count = 15 then PlaySound "Sgt Apone 4": End If
	If FX7Count = 16 then PlaySound "Ripley 3": End If
	If FX7Count = 17 then PlaySound "Gorman 6": End If
	If FX7Count = 18 then PlaySound "Sgt Apone 33": End If
	FX7Count = Int(rnd*19)
End Sub

Dim FX8Count
Sub SFX8_Hit()
	If FX8Count = 0 then PlaySound "Swoosh 16": End If
	If FX8Count = 1 then PlaySound "Swoosh 14": End If
	FX8Count = (FX8Count+1) mod 2
End Sub

Dim FX9Counta, FX9Countb
Sub SFX9_hit()
	FlashSequence2
	If Controlli2.state = 1 Then
	controlli5.state = 0
	SFX23_Hit
	Else
		If FX9Counta = 0 then PlaySound "Hit 11": End If
		If FX9Counta = 1 then PlaySound "Hit 13": End If
		FX9Counta = (FX9Counta+1) mod 2

		If FX9Countb = 2 then
		controlli5.state = 1
		RandomTargetOnline
		End If
		FX9Countb = (FX9Countb+1) mod 30
	End If
End Sub

Dim FX10Count, FX10Countb
Sub SFX10_hit()
	FX10Countb = Int(rnd*13)
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
		If FX10Count = 0 then
		FX11Count=1
		FX18Count=1
		FX19Count=1
			If FX10Countb = 0 then PlaySound "Sgt Apone 11": End If		
			If FX10Countb = 1 then PlaySound "Sgt Apone 16b": End If
			If FX10Countb = 2 then Playsound "Random 21": End If
			If FX10Countb = 3 then PlaySound "Hicks 1": End If
			If FX10Countb = 4 then PlaySound "Comms 32": End If	
			If FX10Countb = 5 then PlaySound "Comms 23": End If
			If FX10Countb = 6 then PlaySound "Sgt Apone 28": End If	
			If FX10Countb = 7 then PlaySound "Bishop 4a": End If
			If FX10Countb = 8 then PlaySound "Sgt Apone 16a": End If		
			If FX10Countb = 9 then PlaySound "Hicks 11": End If
			If FX10Countb = 10 then PlaySound "Sgt Apone 5": End If
			If FX10Countb = 11 then PlaySound "Hudson 31": End If
			If FX10Countb = 12 then PlaySound "Sgt Apone 7": End If
		End If
		If FX10Count = 1 then
		FX11Count=0
		FX18Count=0
		FX19Count=0
		End If
		FX10Count = (FX10Count+1) mod 2	
	End If	
End Sub

Dim FX11Count, FX11Countb
Sub SFX11_hit()
	FX11Countb = Int(rnd*13)
	If controlli6.state = 1 then FX11Count = 1: End If
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
		If FX11Count = 0 then
		FX10Count=1
		FX18Count=1
		FX19Count=1
			If FX11Countb = 0 then PlaySound "Sgt Apone 8": End If
			If FX11Countb = 1 then PlaySound "Random 18": End If
			If FX11Countb = 2 then PlaySound "Comms 34": End If
			If FX11Countb = 3 then PlaySound "Random 17": End If
			If FX11Countb = 4 then PlaySound "Comms 35": End If
			If FX11Countb = 5 then PlaySound "Sgt Apone 27": End If
			If FX11Countb = 6 then PlaySound "Sgt Apone 9": End If
			If FX11Countb = 7 then PlaySound "Ripley 12": End If
			If FX11Countb = 8 then PlaySound "Comms 20": End If
			If FX11Countb = 9 then PlaySound "Random 16": End If
			If FX11Countb = 10 then PlaySound "Comms 17a": End If
			If FX11Countb = 11 then PlaySound "Comms 17b": End If
			If FX11Countb = 12 then PlaySound "Hudson 5": End If
		End If
		If FX11Count = 1 then
		FX10Count=0
		FX18Count=0
		FX19Count=0
		End If
		FX11Count = (FX11Count+1) mod 2
	End If
End Sub

Sub SFX12_hit() PlaySound "Hydraulic Door 2"
	If Controlli8.state = 0 then
	GiOFF
	vpmTimer.AddTimer 500, "GiON '"
	End If
End Sub

Dim FX13Count, FX13Countb
Sub SFX13_hit()
	ScannerOFF
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
		If FX13Count = 0 then PlaySound "Swoosh 14": End If
		If FX13Count = 1 then PlaySound "Swoosh 16": End If
		FX13Count = (FX13Count+1) mod 2
	End If
End Sub

Dim FX14Count
Sub SFX14_Hit() PlaySound "fx_metalrolling"
	FX14Count = Int(rnd*3)
	ScannerOFF
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
		If FX14Count = 0 then PlaySound "Random 10": End If
		If FX14Count = 1 then PlaySound "Random 11": End If
		If FX14Count = 2 then PlaySound "Random 12": End If
	End If
End Sub

Sub SFX14end_Hit() StopSound "fx_metalrolling": End Sub

Dim FX14bCount
Sub SFX14b_Hit()
	ScannerOFF
	If FX14bCount = 0 then PlaySound "Hit 11": End If
	If FX14bCount = 1 then PlaySound "Hit 13": End If
	FX14bCount = (FX14Count+1) mod 2
End Sub

Sub SFX14c_Hit() PlaySound "fx_metalrolling4": End Sub

Dim FX15Count
Sub SFX15_Hit() PlaySound "fx_metalrolling"
	ScannerOFF
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
		If FX15Count = 0 then PlaySound "Random 13": End If
		If FX15Count = 1 then PlaySound "Random 14": End If
		If FX15Count = 2 then PlaySound "Random 15": End If
		FX15Count = (FX15Count+1) mod 3
	End If
End Sub

Sub SFX15end_Hit() StopSound "fx_metalrolling": End Sub

Dim FX16Count
Sub SFX16_Hit()
	Controlli9.state = 0
	StopSFXf
	If Controlli2.state = 1 then 
	PlaySound "Alien 11"
	Else
	EndMusic
	 	If FX16Count = 0 then PlaySound "Vasquez 1": End If
		If FX16Count = 1 then PlaySound "Alien 2": End If
		If FX16Count = 2 then PlaySound "Hudson 16": End If
		If FX16Count = 3 then PlaySound "Alien 6": End If
		FX16Count = (FX16Count+1) mod 4
	End If
End Sub

Dim FX17Count
Sub SFX17_Hit()
	Controlli9.state = 0
	StopSFXf
	If Controlli2.state = 1 then
	PlaySound "Alien 12"
	Else
	EndMusic
		If FX17Count = 0 then PlaySound "Hudson 12": End If
		If FX17Count = 1 then PlaySound "Alien 8": End If
		If FX17Count = 2 then PlaySound "Hudson 17": End If
		If FX17Count = 3 then PlaySound "Alien 1": End If
		FX17Count = (FX17Count+1) mod 4
	End If
End Sub

Dim FX18Count, FX18Countb
Sub SFX18_Hit()
	FX18Countb = Int(rnd*13)
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
	If FX18Count = 0 then
	FX10Count=1
	FX11Count=1
	FX19Count=1
		If controlli5.state = 0 then
			If FX18Countb = 0 then PlaySound "Comms 1a": End If
			If FX18Countb = 1 then PlaySound "Comms 1b": End If
			If FX18Countb = 2 then PlaySound "Comms 3a": End If
			If FX18Countb = 3 then PlaySound "Comms 3b": End If
			If FX18Countb = 4 then PlaySound "Comms 18a": End If
			If FX18Countb = 5 then PlaySound "Comms 18b": End If
			If FX18Countb = 6 then PlaySound "Gorman 1": End If
			If FX18Countb = 7 then PlaySound "Comms 10": End If
			If FX18Countb = 8 then PlaySound "Comms 4": End If
			If FX18Countb = 9 then PlaySound "Comms 33": End If
			If FX18Countb = 10 then PlaySound "Random 20": End If
			If FX18Countb = 11 then PlaySound "Gorman 3": End If
			If FX18Countb = 12 then PlaySound "Comms 21": End If
		End If
	End If
	If FX18Count = 1 then
	FX10Count=0
	FX11Count=0
	FX19Count=0
	End If
	FX18Count = (FX18Count+1) mod 2
	End If
End Sub

Dim FX19Count, FX19Countb
Sub SFX19_Hit()
	FX19Countb = Int(rnd*13)
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
	If FX19Count = 0 then
	FX10Count=1
	FX11Count=1
	FX18Count=1
		If controlli5.state = 0 then
			If FX19Countb = 0 then PlaySound "Comms 5": End If
			If FX19Countb = 1 then PlaySound "Comms 2a": End If
			If FX19Countb = 2 then PlaySound "Comms 2b": End If
			If FX19Countb = 3 then PlaySound "Comms 6a": End If
			If FX19Countb = 4 then PlaySound "Comms 6b": End If
			If FX19Countb = 5 then PlaySound "Comms 15": End If
			If FX19Countb = 6 then PlaySound "Gorman 2": End If
			If FX19Countb = 7 then PlaySound "Comms 36": End If
			If FX19Countb = 8 then PlaySound "Comms 7": End If
			If FX19Countb = 9 then PlaySound "Comms 8a": End If
			If FX19Countb = 10 then PlaySound "Comms 8b": End If
			If FX19Countb = 11 then PlaySound "Gorman 9": End If
			If FX19Countb = 12 then PlaySound "Random 19": End If
		End If
	End If
	If FX19Count = 1 then
	FX10Count=0
	FX11Count=0
	FX18Count=0
	End If   
	FX19Count = (FX19Count+1) mod 2
	End If
End Sub

Dim FX20Count
Sub SFX20_Hit() StopSound "Loader"
	FX20Count = Int(rnd*4)
	ScannerOFF
	GiOFF
	AlertGiOFF
	FFXe
	vpmTimer.AddTimer 900, "FFXe '"
	If liLock.state = 1 then
		controlli5.state = 0
		RandomTargetOffline
		RaiseBlock2
		RaiseBlock4
		RaiseBlock5
		If FX20Count = 0 then EndMusic: PlayMusic "A-MBALL 1.mp3": End If
		If FX20Count = 1 then EndMusic: PlayMusic "A-MBALL 2.mp3": End If
		If FX20Count = 2 then EndMusic: PlayMusic "A-MBALL 3.mp3": End If
		If FX20Count = 3 then EndMusic: PlayMusic "A-MBALL 4.mp3": End If
		vpmTimer.AddTimer 8250, "BlueGiON '"
		vpmTimer.AddTimer 8250, "RaisePosts '"
		vpmTimer.AddTimer 8250, "RaiseAllTargets '"
		vpmTimer.AddTimer 8250, "GiON '"
	End If
	If Controlli2.state = 1 then RaisePostd: End If
	vpmTimer.AddTimer 10000, "DropPostd '"
End Sub

Dim FX21Count
Sub SFX21_Hit()
	liLock.state = 0
	If FX21Count = 0 then Controlli3.state = 1: End If
	If FX21Count = 1 then Controlli3.state = 0: End If
	FX21Count = (FX21Count+1) mod 2
End Sub

Dim FX22Count
Sub SFX22_Hit()
	Controlli1.state = 0
	If FX22Count = 0 then PulseRifle: End If
	If FX22Count = 1 then PlaySound "Critter 1": End If
	FX22Count = (FX22Count+1) mod 2
	vpmTimer.AddTimer 750,"LeftFlipper2.RotateToStart '"
	vpmTimer.AddTimer 750,"FlipperRest '"
End Sub

Sub FlipperRest
	PlaySound "fx_flipperdown"
End Sub

Dim FX23Count, FX23Countb, FX23Countc, FX23Countd, FX23Counte, FX23Countf
Sub SFX23_Hit()
	If Controlli2.state = 1 Then
		If FX23Count = 0 then
			FX23Countb = Int(rnd*21)
			If FX23Countb = 0 then PlaySound "Random 33": End If
			If FX23Countb = 1 then PlaySound "Random 36a": End If
			If FX23Countb = 2 then PlaySound "Random 36b": End If
			If FX23Countb = 3 then PlaySound "Random 8": End If
			If FX23Countb = 4 then PlaySound "Hudson 33": End If
			If FX23Countb = 5 then PlaySound "Random 60": End If
			If FX23Countb = 6 then PlaySound "Random 59": End If
			If FX23Countb = 7 then PlaySound "Random 58": End If
			If FX23Countb = 8 then PlaySound "Random 55": End If
			If FX23Countb = 9 then PlaySound "Hicks 5": End If
			If FX23Countb = 10 then PlaySound "Random 54b": End If
			If FX23Countb = 11 then PlaySound "Random 50a": End If
			If FX23Countb = 12 then PlaySound "Random 50b": End If
			If FX23Countb = 13 then PlaySound "Hicks 2": End If
			If FX23Countb = 14 then PlaySound "Random 6": End If
			If FX23Countb = 15 then PlaySound "Random 10a": End If
			If FX23Countb = 16 then PlaySound "Random 10b": End If
			If FX23Countb = 17 then PlaySound "Random 10c": End If
			If FX23Countb = 18 then PlaySound "Random 43a": End If
			If FX23Countb = 19 then PlaySound "Random 43b": End If
			If FX23Countb = 20 then PlaySound "Random 38": End If
		End If
		If FX23Count = 1 then 
			If FX23Countc = 0 then PlaySound "Alien 1": End If
			If FX23Countc = 1 then PlaySound "Alien 2": End If
			If FX23Countc = 2 then PlaySound "Alien 15": End If
			If FX23Countc = 3 then PlaySound "Alien 5": End If
			If FX23Countc = 4 then PlaySound "Alien 12": End If
			FX23Countc = (FX23Countc+1) mod 5
		End If	
		If FX23Count = 2 Then
			FX23Countd = Int(rnd*21)
			If FX23Countd = 0 then PlaySound "Random 51": End If
			If FX23Countd = 1 then PlaySound "Random 28": End If
			If FX23Countd = 2 then PlaySound "Random 41": End If
			If FX23Countd = 3 then PlaySound "Random 56a": End If
			If FX23Countd = 4 then PlaySound "Random 56b": End If
			If FX23Countd = 5 then PlaySound "Loader 2": End If
			If FX23Countd = 6 then PlaySound "Random 54a": End If
			If FX23Countd = 7 then PlaySound "Random 45": End If
			If FX23Countd = 8 then PlaySound "Random 46": End If
			If FX23Countd = 9 then PlaySound "Hudson 25": End If
			If FX23Countd = 10 then PlaySound "Hudson 23": End If
			If FX23Countd = 11 then PlaySound "Hudson 24a": End If
			If FX23Countd = 12 then PlaySound "Hudson 24b": End If
			If FX23Countb = 13 then PlaySound "Hudson 24c": End If
			If FX23Countd = 14 then PlaySound "Vasquez 4": End If
			If FX23Countd = 15 then PlaySound "Hudson 9": End If
			If FX23Countd = 16 then PlaySound "Random 53": End If
			If FX23Countd = 17 then PlaySound "Random 52": End If
			If FX23Countd = 18 then PlaySound "Random 48": End If
			If FX23Countd = 19 then PlaySound "Random 5": End If
			If FX23Countd = 20 then PlaySound "Random 47": End If
		End If
		If FX23Count = 3 then
			If FX23Counte = 0 then PulseRifle: FFXc: End If
			If FX23Counte = 1 then PulseRifle: FFXe: End If
			If FX23Counte = 2 then PulseRifle: FFXa: End If
			FX23Counte = (FX23Counte+1) mod 3
		End If
		If FX23Count = 4 Then
			If FX23Countf = 0 then PlaySound "Gorman 5": End If
			If FX23Countf = 1 then PlaySound "Comms 29": End If	
			If FX23Countf = 2 then PlaySound "Random 9": End If
			If FX23Countf = 3 then PlaySound "Random 49": End If
			If FX23Countf = 4 then PlaySound "Random 37": End If
			If FX23Countf = 5 then PlaySound "Random 29a": End If
			If FX23Countf = 6 then PlaySound "Random 29b": End If
			If FX23Countf = 7 then PlaySound "Comms 27": End If
			If FX23Countf = 8 then PlaySound "Comms 30": End If
			If FX23Countf = 9 then PlaySound "Comms 28": End If
			If FX23Countf = 10 then PlaySound "Hicks 18": End If
			If FX23Countf = 11 then PlaySound "Random 57": End If
			If FX23Countf = 12 then PlaySound "Random 40": End If
			If FX23Countf = 13 then PlaySound "Random 44": End If
			If FX23Countf = 14 then PlaySound "Hudson 30": End If
			If FX23Countf = 15 then PlaySound "Random 39": End If
			If FX23Countf = 16 then PlaySound "Hicks 17": End If
			FX23Countf = (FX23Countf+1) mod 17
		End If
		FX23Count = (FX23Count+1) mod 5
	End If
End Sub

Dim PRCount, PRCountb
Sub PulseRifle()
	PRCount = Int(rnd*8): PRCountb = Int(rnd*29)
	If PRCount = 0 then PFflasher: PlaySound "Pulse Rifle 1": End If
	If PRCount = 1 then PFflasher: PlaySound "Pulse Rifle 2": End If
	If PRCount = 2 then PFflasher: PlaySound "Pulse Rifle 3": End If
	If PRCount = 3 then PFflasher: PlaySound "Pulse Rifle 4": End If
	If PRCount = 4 then PFflasher: PlaySound "Pulse Rifle 5": End If
	If PRCount = 5 then PFflasher: PlaySound "Pulse Rifle 6": End If
	If PRCount = 6 then PFflasher: PlaySound "Pulse Rifle 7": End If
	If PRCount = 7 then PFflasher: PlaySound "Pulse Rifle 8": End If	

	If PRCountb = 14 Then PlaySound "Ricochet3": End If
	If PRCountb = 28 Then PlaySound "Ricochet4": End If
End Sub

Dim PRTCount
Sub PRTrigger_Hit()
	If Timer3.Enabled = True then PulseRifleMod: End If
	If Controlli2.state = 1 and Controlli9.state = 1 then
		If PRTCount = 15 then ReRaiseTargets: End If
		PRTCount = (PRTCount+1) mod 16
	End If
	Timer3.Enabled = False
End Sub

Sub PulseRifleMod()
	PlaySound "Pulse Rifle 1": PlaySound "Ricochet3": PFflasher
End Sub

'***************
' Kickers
'***************

Dim Kicker2Count
Sub Kicker2_hit() PlaySound "fx_kicker_enter": PlaySound "DataLoad": StopSound "Loader"
	Kicker2.TimerInterval = 6000
	Kicker2.TimerEnabled = 1
	Timer1.Enabled = False
	GiOFF
	AlertGiOFF
	GreenGiON
	Kicker2Count = Int(rnd*18)
	If Controlli2.state = 1 then
	SFX23_Hit
	Else
		If Kicker2Count = 0 then PlaySound "Ripley 10": End If
		If Kicker2Count = 1 then PlaySound "Burke 1": End If
		If Kicker2Count = 2 then PlaySound "Hicks 4": End If
		If Kicker2Count = 3 then PlaySound "Hudson 35": End If
		If Kicker2Count = 4 then PlaySound "Ripley 19": End If
		If Kicker2Count = 5 then PlaySound "Random 4": End If
		If Kicker2Count = 6 then PlaySound "Hudson 1": End If
		If Kicker2Count = 7 then PlaySound "Hudson 26": End If
		If Kicker2Count = 8 then PlaySound "Ripley 15": End If
		If Kicker2Count = 9 then PlaySound "Hudson 3": End If
		If Kicker2Count = 10 then PlaySound "Hudson 27": End If
		If Kicker2Count = 11 then PlaySound "Bishop 2": End If
		If Kicker2Count = 12 then PlaySound "Hudson 28": End If
		If Kicker2Count = 13 then PlaySound "Ripley 22": End If
		If Kicker2Count = 14 then PlaySound "Ripley 11": End If
		If Kicker2Count = 15 then PlaySound "Vasquez 3": End If
		If Kicker2Count = 16 then PlaySound "Random 1": End If
		If Kicker2Count = 17 then PlaySound "Ripley 18": End If
	End If
End Sub
	
Sub Kicker2_Timer
	Kicker2.Kick 130,6: PlaySound"fx_kicker": PlaySound"DataLoad out"
	Kicker2.TimerEnabled = 0
	GreenGiOFF
	FX11Count = 1
	vpmTimer.AddTimer 1000, "GiON '"
End Sub

Sub Kicker3_Hit() PlaySound "Sentry gun activated"
	Kicker3.TimerInterval = 1000
	Kicker3.TimerEnabled = 1
End sub

Sub Kicker3_Timer
	Kicker3.Kick 260,3
	Kicker3.TimerEnabled = 0
End Sub

Dim Kicker4Count
Sub Kicker4_Hit() PlaySound "fx_kicker_enter"
	Kicker4.TimerInterval = 2500
	Kicker4.TimerEnabled = 1
	Kicker4Count = Int(rnd*5)
	FX10Count = 1
	If Kicker4Count = 0 then PlaySound "Hicks 13": End If
	If Kicker4Count = 1 then PlaySound "Hudson 38": End If
	If Kicker4Count = 2 then PlaySound "Hicks 15": End If
	If Kicker4Count = 3 then PlaySound "Hudson 39": End If
	If Kicker4Count = 4 then PlaySound "Hicks 14": End If
End Sub

Sub Kicker4_Timer
	Kicker4.Kick 260,6: PlaySound "fx_kicker"
	Kicker4.TimerEnabled =0
End Sub

Dim Kicker5Count
Sub Kicker5_Hit() PlaySound "fx_kicker_enter"
	Kicker5.TimerInterval = 2500
	Kicker5.TimerEnabled = 1
	FFXb
	Kicker5Count = Int(rnd*6)
	If Controlli2.state = 1 then
	SFX23_Hit
	Else
		If Kicker5Count = 0 then PlaySound "Medic 2": End If
		If Kicker5Count = 1 then PlaySound "Ripley 6": End If
		If Kicker5Count = 2 then PlaySound "Random 35": End If
		If Kicker5Count = 3 then PlaySound "Ripley 2": End If
		If Kicker5Count = 4 then PlaySound "Hicks 3": End If
		If Kicker5Count = 5 then PlaySound "Bishop 4b": End If
	End If
End Sub

Sub Kicker5_Timer
	Kicker5.Kick 0,90,1: PlaySound "fx_kicker": PlaySound "Pneumatic"
	Kicker5.TimerEnabled = 0
	ScannerON
	Flasher5ON
End Sub

Dim Kicker6Count
Sub Kicker6_Hit() PlaySound "fx_kicker_enter"
	If Timer2.Enabled = True then
		Kicker6Count = 15
		Kicker6.TimerInterval = 4500
		Kicker6.TimerEnabled = 1
	Else
		Kicker6.TimerInterval = 3000
		Kicker6.TimerEnabled = 1
	End If
	If Controlli2.state = 1 Then
	SFX23_Hit
	Else
		If Kicker6Count = 0 then PlaySound "Comms 12": End If
		If Kicker6Count = 1 then PlaySound "Hicks 9": End If
		If Kicker6Count = 2 then PlaySound "Sgt Apone 10": End If
		If Kicker6Count = 3 then PlaySound "Hudson 22": End If
		If Kicker6Count = 4 then PlaySound "Sgt Apone 13": End If
		If Kicker6Count = 5 then PlaySound "Comms 9": End If
		If Kicker6Count = 6 then PlaySound "Comms 24a": End If
		If Kicker6Count = 7 then PlaySound "Comms 24b": End If
		If Kicker6Count = 8 then PlaySound "Comms 13": End If
		If Kicker6Count = 9 then PlaySound "Sgt Apone 26": End If
		If Kicker6Count = 10 then PlaySound "Sgt Apone 15": End If
		If Kicker6Count = 11 then PlaySound "Random 7": End If
		If Kicker6Count = 12 then PlaySound "Sgt Apone 14": End If
		If Kicker6Count = 13 then PlaySound "Medic 1": End If
		If Kicker6Count = 14 then PlaySound "Hudson 14": End If
		If Kicker6Count = 15 then PlaySound "No Sound": End If
		Kicker6Count = Int(rnd*16)
	End If
End Sub

Sub Kicker6_Timer
	Kicker6.Kick 180,8: PlaySound "fx_kicker"
	Kicker6.TimerEnabled = 0
End Sub

Dim K7Count, K7Countb
Sub Kicker7_Hit() PlaySound "fx_kicker_enter": PlaySound "Door": StopSound "Loader"
	If K7Count = 0 then
	K7Countb = Int(rnd*3)
		If K7Countb = 0 then Kicker7.TimerInterval = 1000: End if
		If K7Countb = 1 then Kicker7.TimerInterval = 2000: End if
		If K7Countb = 2 then Kicker7.TimerInterval = 3000: End if
	Kicker7.TimerEnabled = 1
	Flasher5ON
	ScannerOFF
	End If
	If K7Count = 1 then
	Kicker7.TimerInterval = 8000
	Kicker7.TimerEnabled = 1
	Controlli1.state = 1
	Flasher5ON
	ScannerOFF
	AlertGiOFF
		If Controlli2.state = 0 then
		vpmTimer.AddTimer 1000, "GiOFF '"
		vpmTimer.AddTimer 1000, "EndMusic '"
		vpmTimer.AddTimer 2000, "Kicker7SFX '"
		vpmTimer.AddTimer 2000, "GalleryGiON '"
		vpmTimer.AddTimer 3000, "RaiseTargetX '"
		vpmTimer.AddTimer 5000, "Kicker7SFXb '"
		End If
	End If
End Sub

Dim K7Countc
Sub Kicker7_Timer()
	K7Countc = Int(rnd*3)
	If K7Count = 0 then
		If K7Countc = 0 then Kicker7.Kick 205,10: End If
		If K7Countc = 1 then Kicker7.Kick 205,20: End If
		If K7Countc = 2 then Kicker7.Kick 205,30: End If
	Kicker7.TimerEnabled = 0
	PlaySound "fx_kicker": PlaySound "Pneumatic"
	Flasher5ON
	ScannerON
	End If
	If K7Count = 1 then
	Kicker7.Kick 130,2: PlaySound "fx_kicker": PlaySound "Pneumatic"
	Kicker7.TimerEnabled = 0
	End If
	K7Count = (K7Count+1) mod 2
End Sub

Sub Kicker7SFX
	PlaySound "Ambient 1"
End Sub

Sub StopKicker7SFX
	StopSound "Ambient 1"
End Sub

Dim K7Countd
Sub Kicker7SFXb()
	If K7Countd = 0 then PlaySound "Random 22": End If
	If K7Countd = 1 then PlaySound "Vents 3 ": End If
	If K7Countd = 2 then PlaySound "Sgt Apone 25": End If
	K7Countd = (K7Countd+1) mod 3
End Sub

Sub Kicker8_Hit() PlaySound "fx_kicker_enter"
	TargetLightOFF
	vpmTimer.AddTimer 1000, "GalleryGiOFF '"
	Gate16.Collidable = False
	If Controlli4.state = 0 then
		Kicker8.TimerInterval = 3000
		Kicker8.TimerEnabled = 1
		vpmTimer.AddTimer 1000, "Kicker8SFXb '"
	Else
		Kicker8.TimerInterval = 6000
		Kicker8.TimerEnabled = 1
		vpmTimer.AddTimer 1000, "PlaySFXf '"
		vpmTimer.AddTimer 2500, "Kicker8SFX '"
		vpmTimer.AddTimer 5000, "FlashersON_mod '"
		Controller.Switch(17) = 0
		Controller.Switch(18) = 0
		Controller.Switch(19) = 0
		Controller.Switch(20) = 0
	End If
End Sub

Sub Kicker8_Timer() PlaySound "fx_kicker"
	Kicker8.Kick 0,32
	Kicker8.TimerEnabled = 0	
End Sub

Dim K8Count
Sub Kicker8SFX
	If K8Count = 0 then Playsound "Sgt Apone 17a": End If
	If K8Count = 1 then Playsound "Sgt Apone 24": End If
	If K8Count = 2 then Playsound "Sgt Apone 18": End If
	K8Count = (K8Count+1) mod 3
End Sub

Dim K8Countb
Sub Kicker8SFXb
	K8Countb = Int(rnd*4)
	If K8Countb = 0 then PlaySound "Hardware1": End If
	If K8Countb = 1 then PlaySound "Hardware4": End If
	If K8Countb = 2 then PlaySound "Hardware6": End If
	If K8Countb = 3 then PlaySound "Hardware7": End If
End Sub

Sub Kicker9_Hit() PlaySound "fx_kicker_enter": StopSound "Loader"
	If SGli.state = 0 and li7.state = 0 then
		Kicker9.TimerInterval = 4500
		Kicker9.TimerEnabled = 1
		GiOFF
		DropBlock1
	End If
	If SGli.state = 1 then
		Kicker9.TimerInterval = 6500
		Kicker9.TimerEnabled = 1
		EndMusic
		GiOFF
		K9K10SFX
		vpmTimer.AddTimer 6500, "MusicMod1 '"
	End If
End Sub

Sub Kicker9_Timer() PlaySound "Hit 2": PlaySound "fx_kicker"
	Kicker9.Kick 30,130
	Kicker9.TimerEnabled = 0
	Gate8.Collidable = False
	FlashersON
	RaiseBlock3
	vpmTimer.AddTimer 500, "DropBlock3 '"
	vpmTimer.AddTimer 750, "GiON '"
	vpmTimer.AddTimer 1000, "Gate8Reset '"
End Sub

Sub Kicker10_Hit() PlaySound "fx_kicker_enter": StopSound "Loader"
	FX11Count = 1
	If SGli.state = 0 and li7.state = 0 then
		Kicker10.TimerInterval = 4500
		Kicker10.TimerEnabled = 1
		GiOFF
		DropBlock1
	End If
	If SGli.state = 1 then
		Kicker10.TimerInterval = 6500
		Kicker10.TimerEnabled = 1
		EndMusic
		GiOFF
		K9K10SFX
		vpmTimer.AddTimer 6500, "MusicMod1 '"
	End If
End Sub

Sub Kicker10_Timer() PlaySound "Hit 2": PlaySound "fx_kicker"
	Kicker10.Kick 340,130
	Kicker10.TimerEnabled = 0
	FlashersON
	RaiseBlock6
	vpmTimer.AddTimer 500, "DropBlock6 '"
	vpmTimer.AddTimer 750, "GiON '"
End Sub

Dim K9K10Count
Sub K9K10SFX() K9K10Count = Int(rnd*6)
	EndMusic
	If K9K10Count = 0 then PlaySound "Vents 3": End If
	If K9K10Count = 1 then PlaySound "Sulaco": End If
	If K9K10Count = 2 then PlaySound "Vents 4": End If
	If K9K10Count = 3 then PlaySound "Facehugger": End If
	If K9K10Count = 4 then PlaySound "Alien 14": End If
	If K9K10Count = 5 then PlaySound "Alien 13": End If
End Sub

Sub Kicker11_Hit() PlaySound "fx_kicker_enter"
	vpmTimer.AddTimer 500, "PlaySFXc '"
	Kicker11.TimerInterval = 3500
	Kicker11.TimerEnabled = 1
End Sub

Dim K11Count
Sub Kicker11_Timer() PlaySound "fx_kicker"
	K11Count = Int(rnd*3)
	If K11Count = 0 then Kicker11.Kick 0,25
	If K11Count = 1 then Kicker11.Kick 0,35
	If K11Count = 2 then Kicker11.Kick 0,45
	Kicker11.TimerEnabled = 0
	Flasher7ON
	PlaySFXa
End Sub

'*************************
' Controlled SoundFX
'*************************

Sub ScannerON()
	PlaySound "Motion Tracker"
	lightScanner.state = 2
End Sub

Sub ScannerOFF()
	StopSound "Motion Tracker"
	lightScanner.state = 0
End Sub

Sub DataLoadON
	PlaySound "DataLoad 4"
End Sub

Sub DataLoadOFF
	StopSound "DataLoad 4"
End Sub

Dim MM1Count
Sub MusicMod1()
	If MM1Count = 0 then PlayMusic "A-TRACK 3.mp3": End If
	If MM1Count = 1 then PlayMusic "A-TRACK 5.mp3": End If
	If MM1Count = 2 then PlayMusic "A-TRACK 7.mp3": End If
	MM1Count = (MM1Count+1) mod 3
End Sub

Sub StopMusicMod1()
	EndMusic "A-TRACK 3.mp3"
	EndMusic "A-TRACK 5.mp3"
	EndMusic "A-TRACK 7.mp3"
End Sub

Dim MM2Count
Sub MusicMod2()
	If MM2Count = 0 then PlayMusic "A-TRACK 1.mp3": End If
	If MM2Count = 1 then PlayMusic "A-TRACK 4.mp3": End If
	If MM2Count = 2 then PlayMusic "A-TRACK 9.mp3": End If
	MM2Count = (MM2Count+1) mod 3
End Sub

Sub StopMusicMod2()
	EndMusic "A-TRACK 1.mp3"
	EndMusic "A-TRACK 4.mp3"
	EndMusic "A-TRACK 9.mp3"
End Sub

Sub PlayTrack9() PlayMusic "A-TRACK 9.mp3": End Sub

Sub PlayTrack8() PlayMusic "A-TRACK 8.mp3": End Sub

Dim SensorCount
Sub SensorSFX
	If SensorCount = 0 then PlaySound "Ricochet3": End If
	If SensorCount = 1 then PlaySound "Ricochet4": End If
	SensorCount = (SensorCount+1) mod 2
End Sub

Dim RATCount
Sub RaiseAllTargetsSFX()
	If RATCount = 0 then PlaySound "Random 42": End If
	If RATCount = 1 then PlaySound "Hudson 30": End If
	If RATCount = 2 then PlaySound "Hudson 11": End If
	RATCount = Int(rnd*3)
End Sub

Sub PlaySFXa() PlaySound "Pressure": End Sub

Dim SFXbCount
Sub PlaySFXb()
	SFXbCount = Int(rnd*3)
	If SFXbCount = 0 then PlaySound "Ambient 2a": End If
	If SFXbCount = 1 then PlaySound "Ambient 2b": End If
	If SFXbCount = 2 then PlaySound "Ambient 2c": End If
	End Sub

Sub StopSFXb()
	StopSound "Ambient 2a"
	StopSound "Ambient 2b"
	StopSound "Ambient 2c"
End Sub

Sub PlaySFXc() PlaySound "Hardware3": End Sub

Sub PlaySFXd() PlaySound "Alien 10b": End Sub

Sub PlaySFXe() PlaySound "Explosion": End Sub 

Sub PlaySFXf() PlaySound "Snippet 7": End Sub

Sub StopSFXf() StopSound "Snippet 7": End Sub

Sub PlaySFXg() PlaySound "Target Neutralised": End Sub

Sub PlaySFXh() PlaySound "Snippet 9": End Sub

'*************************
' Special Walls / Targets
'*************************

Sub WallSensor1_Hit() SensorSFX: End Sub
Sub WallSensor2_Hit() SensorSFX: End Sub
Sub WallSensor3_Hit() SensorSFX: End Sub
Sub WallSensor4_Hit() SensorSFX: End Sub

Sub Wall8_Hit() PlaySound "fx_Metal_Touch", 0, 0.5, Pan(ActiveBall), 0: End Sub
Sub Wall100_Hit() PlaySound "Sentry gun 4": PlaySound "Ricochet3": End Sub

' * POSTS *

Sub	StopPosta_Hit() vpmTimer.AddTimer 1000,"DropPosta '": PlaySound "fx_metalhit", 0, 0.5, -0.1, 0.25: End Sub
Sub	StopPostb_Hit() vpmTimer.AddTimer 1000,"DropPostb '": PlaySound "fx_metalhit", 0, 0.5, 0.1, 0.25: End Sub
Sub StopPostc_Hit() vpmTimer.AddTimer 1000,"DropPostc '": PlaySound "fx_metalhit", 0, 1, 0, 0.25: End Sub

Sub RaisePosts()
	If StopPosta.Isdropped = True or StopPostb.IsDropped = True Then
	StopPosta.IsDropped = 0: PlaySound "fx_resetdrop": PlaySound "pneumatic b"
	StopPostb.IsDropped = 0: PlaySound "fx_resetdrop": PlaySound "pneumatic b"
	End If
End Sub

Sub RaisePostc() PlaySound "fx_resetdrop": PlaySound "pneumatic b"
	If StopPostc.IsDropped = True then StopPostc.IsDropped = False: End If
End Sub

Sub RaisePostd()
	If StopPostd.IsDropped = True then StopPostd.IsDropped = False: End If
End Sub

Sub DropPosta() PlaySound "fx_droptarget": PlaySound "pneumatic b"
	If StopPosta.IsDropped = False then StopPosta.IsDropped = True:End If
End Sub

Sub DropPostb() PlaySound "fx_droptarget": PlaySound "pneumatic b"
	If StopPostb.IsDropped = False then StopPostb.IsDropped = True:End If
End Sub

Sub DropPostc() PlaySound "fx_droptarget"
	If StopPostc.IsDropped = False then StopPostc.IsDropped = True:End If
End Sub

Sub DropPostd()
	If StopPostd.IsDropped = False then StopPostd.IsDropped = True:End If
End Sub

'* LEFT HOLD POST *

Sub LeftPostHold_Hit() PlaySound "fx_metalhit", 0, 0.5, -0.1, 0.25: End Sub
Sub RaiseFX() PlaySound "HUD Interface 1":End Sub

Sub RaiseLeftHold() PlaySound "fx_resetdrop": PlaySound "Alarm": PlaySound "pneumatic b": StopSound "Loader"
	If LeftPostHold.IsDropped = True then LeftPostHold.IsDropped = False: End If
	EndMusic
	GiOFF
	FFXd
	vpmTimer.AddTimer 3000, "PlaySFXh '"
	vpmTimer.AddTimer 4000, "BlueGiON '"
	vpmTimer.AddTimer 1500, "HolderSFX '"
	vpmTimer.AddTimer 5000, "RightRandomTargetRaise '"
	vpmTimer.AddTimer 10000, "BlueGiOFF '"	
End Sub

Sub DropLeftHold() PlaySound "fx_droptarget": PlaySound "pneumatic b": StopSound "Alarm"
	controlli5.state = 0
	controlli6.state = 1
	RandomTargetOffline
	FX9Countb = 0
	sw26Count = 0 
	Timer1.Enabled = True
	Timer3.Enabled = True
	If LeftPostHold.IsDropped = False then LeftPostHold.IsDropped = True: End If
	FX10Count=1
	FX11Count=1
	FX18Count=1
	FX19Count=1
End Sub

'* RIGHT HOLD POST *

Sub RightPostHold_Hit() PlaySound "fx_metalhit", 0, 0.5, 0.1, 0.25: End Sub

Sub RaiseRightHold() PlaySound "fx_resetdrop": PlaySound "Alarm": PlaySound "pneumatic b": StopSound "Loader"
	If RightPostHold.IsDropped = True then RightPostHold.IsDropped = False: End If
	EndMusic
	GiOFF
	FFXd
	vpmTimer.AddTimer 3000, "PlaySFXh '"
	vpmTimer.AddTimer 4000, "BlueGiON '"
	vpmTimer.AddTimer 1500, "HolderSFX '"
	vpmTimer.AddTimer 5000, "LeftRandomTargetRaise '"
	vpmTimer.AddTimer 10000, "BlueGiOFF '"
End Sub

Sub DropRightHold() PlaySound "fx_droptarget": PlaySound "pneumatic b": StopSound "Alarm"
	controlli5.state = 0
	controlli7.state = 1
	RandomTargetOffline
	FX9Countb = 0
	sw27Count = 0
	Timer1.Enabled = True
	Timer3.Enabled = True
	If RightPostHold.IsDropped = False then RightPostHold.IsDropped = True: End If
	FX10Count=1
	FX11Count=1
	FX18Count=1
	FX19Count=1
End Sub

Dim HolderCount
Sub HolderSFX()
	IF HolderCount = 0 then PlaySound "Random 32" :End If
	IF HolderCount = 1 then PlaySound "Random 22" :End If
	IF HolderCount = 2 then PlaySound "Sgt Apone 25" :End If
	HolderCount = (HolderCount+1) mod 3
End Sub

'* TIMER TIMERS *

Sub Timer1_Timer()
	Timer1.Enabled = False
	DATMod1
End Sub

Sub Timer2_Timer()
	Timer2.Enabled = False
End Sub

Sub Timer3_Timer()
	Timer3.Enabled = False
End Sub

'* WALL BLOCKS *

Sub RaiseBlock1()
	If WallBlock1.IsDropped = True then WallBlock1.IsDropped = False: End If
	SGli.state = 0
End Sub

Sub DropBlock1() PlaySound "Hudson 32"
	If WallBlock1.IsDropped = False then WallBlock1.IsDropped = True: End If
	vpmTimer.AddTimer 1000, "SentryLightON '"
	vpmTimer.AddTimer 3000, "SentryLightOFF '"
	vpmTimer.AddTimer 3500, "SGOnline '"
End Sub

Sub SGOnline() SGli.state = 1: PlaySound "HUD Interface 7": End Sub

Sub SGOffline() SGli.state = 0: PlaySound "HUD Interface 7": End Sub

Sub RaiseBlock2()
	If WallBlock2.IsDropped = True then WallBlock2.IsDropped = False: End If
End Sub

Sub DropBlock2()
	If WallBlock2.IsDropped = False then WallBlock2.IsDropped = True: End If
End Sub

Sub RaiseBlock3()
	If WallBlock3.IsDropped = True then WallBlock3.IsDropped = False: End If
End Sub

Sub DropBlock3()
	If WallBlock3.IsDropped = False then WallBlock3.IsDropped = True: End If
End Sub

Sub RaiseBlock4()
	If WallBlock4.IsDropped = True then WallBlock4.IsDropped = False: End If
End Sub

Sub DropBlock4()
	If WallBlock4.IsDropped = False then WallBlock4.IsDropped = True: End If
End Sub

Sub RaiseBlock5()
	If WallBlock5.IsDropped = True then WallBlock5.IsDropped = False: End If
End Sub

Sub DropBlock5()
	If WallBlock5.IsDropped = False then WallBlock5.IsDropped = True: End If
End Sub

Sub RaiseBlock6()
	If WallBlock6.IsDropped = True then WallBlock6.IsDropped = False: End If
End Sub

Sub DropBlock6()
	If WallBlock6.IsDropped = False then WallBlock6.IsDropped = True: End If
End Sub

'* TARGET X *

Sub TargetX_Hit() PlaySound "fx_droptarget"
	PlaySFXe
	Flasher5ON
	vpmTimer.AddTimer 500, "PlaySFXd '"
	vpmTimer.AddTimer 1000, "GreenGiFlash '"
	vpmTimer.AddTimer 1400, "GreenGiFlash '"
	vpmTimer.AddTimer 1800, "GreenGiFlash '"
	Controlli4.state = 1
	TargetX.IsDropped = True
	Controller.Switch(17) = 1
	Controller.Switch(18) = 1
	Controller.Switch(19) = 1
	Controller.Switch(20) = 1
End Sub

Sub RaiseTargetX() PlaySound "fx_resetdrop": PlaySound "Alien 6"
	Targetli.state = 1
	If TargetX.IsDropped = 1 then TargetX.IsDropped = 0: End If
End Sub

Dim TargetXCount
Sub DropTargetX() PlaySound "fx_droptarget"
	If TargetXCount = 0 then PlaySound "Alien 5": End If
	If TargetXCount = 1 then PlaySound "Alien 2": End If
	If TargetXCount = 2 then PlaySound "Alien 12": End If
	TargetXCount = (TargetXCount+1) mod 3
	If TargetX.IsDropped = 0 then TargetX.IsDropped = 1: End If
End Sub

'* RAISE ALL TARGETS *

Sub RaiseAllTargets() PlaySound "Alien 1"
	vpmTimer.AddTimer 500, "RaiseAllTargetsSFX '"
	If Target1.IsDropped = 1 then RaiseTarget1: End If
	If Target2.IsDropped = 1 then RaiseTarget2: End If
	If Target3.IsDropped = 1 then RaiseTarget3: End If
	If Target4.IsDropped = 1 then RaiseTarget4: End If
	If Target5.IsDropped = 1 then RaiseTarget5: End If
	If Target6.IsDropped = 1 then RaiseTarget6: End If
	If Target7.IsDropped = 1 then RaiseTarget7: End If
	If Target8.IsDropped = 1 then RaiseTarget8: End If
	If Target9.IsDropped = 1 then RaiseTarget9: End If
End Sub

Sub ReRaiseTargets()
		If Target1.IsDropped = 1 and Target2.IsDropped = 1 and Target3.IsDropped = 1 then PlaySound "Alien 6": RaiseTarget2: RaiseTarget3: End If
		If Target4.IsDropped = 1 and Target5.IsDropped = 1 and Target6.IsDropped = 1 then PlaySound "Alien 7": RaiseTarget4: RaiseTarget5: RaiseTarget6: End If
		If Target7.IsDropped = 1 and Target8.IsDropped = 1 and Target9.IsDropped = 1 then PlaySound "Alien 8": RaiseTarget7: RaiseTarget8: End If
End Sub

'* DROP ALL TARGETS *

Sub DropAllTargets()
	controlli6.state = 0
	controlli7.state = 0
	If Target1.IsDropped = 0 then DropTarget1: End If
	If Target2.IsDropped = 0 then DropTarget2: End If
	If Target3.IsDropped = 0 then DropTarget3: End If
	If Target4.IsDropped = 0 then DropTarget4: End If
	If Target5.IsDropped = 0 then DropTarget5: End If
	If Target6.IsDropped = 0 then DropTarget6: End If
	If Target7.IsDropped = 0 then DropTarget7: End If
	If Target8.IsDropped = 0 then DropTarget8: End If
	If Target9.IsDropped = 0 then DropTarget9: End If
End Sub

Sub DATMod1()
	DropAllTargets
	MusicMod2
	vpmTimer.AddTimer 500, "GiON '"
End Sub

Sub DATMod2() PlaySound "Alien 12": PlaySound "Snippet 4"
	DropAllTargets
End Sub

Sub DATMod3()
	DropAllTargets
End Sub

'* RAISE RANDOM TARGET *

Dim RaiseTargetCounta
Sub LeftRandomTargetRaise() PlaySound "Alien 6"
	RaiseTargetCounta = Int(rnd*6)
	If RaiseTargetCounta = 0 then RaiseTarget1: vpmTimer.AddTimer 500, "DropTarget1 '": vpmTimer.AddTimer 1000, "RaiseTarget1 '": End If
	If RaiseTargetCounta = 1 then RaiseTarget2: vpmTimer.AddTimer 500, "DropTarget2 '": vpmTimer.AddTimer 1000, "RaiseTarget2 '": End If
	If RaiseTargetCounta = 2 then RaiseTarget3: vpmTimer.AddTimer 500, "DropTarget3 '": vpmTimer.AddTimer 1000, "RaiseTarget3 '": End If
	If RaiseTargetCounta = 3 then RaiseTarget4: vpmTimer.AddTimer 500, "DropTarget4 '": vpmTimer.AddTimer 1000, "RaiseTarget4 '": End If
	If RaiseTargetCounta = 4 then RaiseTarget5: vpmTimer.AddTimer 500, "DropTarget5 '": vpmTimer.AddTimer 1000, "RaiseTarget5 '": End If
	If RaiseTargetCounta = 5 then RaiseTarget6: vpmTimer.AddTimer 500, "DropTarget6 '": vpmTimer.AddTimer 1000, "RaiseTarget6 '": End If
End Sub

Dim RaiseTargetCountb
Sub RightRandomTargetRaise() PlaySound "Alien 7"
	RaiseTargetCountb = Int(rnd*6)
	If RaiseTargetCountb = 0 then RaiseTarget4: vpmTimer.AddTimer 500, "DropTarget4 '": vpmTimer.AddTimer 1000, "RaiseTarget4 '": End If
	If RaiseTargetCountb = 1 then RaiseTarget5: vpmTimer.AddTimer 500, "DropTarget5 '": vpmTimer.AddTimer 1000, "RaiseTarget5 '": End If
	If RaiseTargetCountb = 2 then RaiseTarget6: vpmTimer.AddTimer 500, "DropTarget6 '": vpmTimer.AddTimer 1000, "RaiseTarget6 '": End If
	If RaiseTargetCountb = 3 then RaiseTarget7: vpmTimer.AddTimer 500, "DropTarget7 '": vpmTimer.AddTimer 1000, "RaiseTarget7 '": End If
	If RaiseTargetCountb = 4 then RaiseTarget8: vpmTimer.AddTimer 500, "DropTarget8 '": vpmTimer.AddTimer 1000, "RaiseTarget8 '": End If
	If RaiseTargetCountb = 5 then RaiseTarget9: vpmTimer.AddTimer 500, "DropTarget9 '": vpmTimer.AddTimer 1000, "RaiseTarget9 '": End If
End Sub

'* TARGETS HIT *

Sub Target1_Hit() PlaySound "fx_droptarget"
	DropTarget1
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet1
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 6"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target2_Hit() PlaySound "fx_droptarget"
	DropTarget2
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet2
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 7"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target3_Hit() PlaySound "fx_droptarget"
	DropTarget3
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet3
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 8"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target4_Hit() PlaySound "fx_droptarget"
	DropTarget4
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet4
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 9"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target5_Hit() PlaySound "fx_droptarget"
	DropTarget5
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet5
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 10"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target6_Hit() PlaySound "fx_droptarget"
	DropTarget6
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet6
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 9"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target7_Hit() PlaySound "fx_droptarget"
	DropTarget7
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet7
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 8"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target8_Hit() PlaySound "fx_droptarget"
	DropTarget8
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet8
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 7"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub Target9_Hit() PlaySound "fx_droptarget"
	DropTarget9
	vpmTimer.AddTimer 1200, "FFXe '"
	If controlli6.state = 1 or controlli7.state = 1 then
		TargetFX
		EnableMagnet9
		vpmtimer.AddTimer 6000, "DisableMagnets '"
	Else
		PlaySound "Alien 6"
		SFX23_Hit
	End If
	controlli6.state = 0
	controlli7.state = 0
End Sub

Sub TargetFX() 
	Timer1.Enabled = False
	FlashersON_mod
	PlaySFXe
	FX10Count=1
	FX11Count=1
	FX18Count=1
	FX19Count=1
	vpmTimer.AddTimer 500, "PlaySFXd '"
	vpmTimer.AddTimer 2100, "FFXe '"
	vpmTimer.AddTimer 3000, "FFXe '" 
	vpmTimer.AddTimer 2000, "PlaySFXg '"
	vpmTimer.AddTimer 4000, "PlaySFXf '" 
	vpmTimer.AddTimer 5500, "PlaySFXa '"
	vpmTimer.AddTimer 5500, "BlueGiOFF '" 
	vpmTimer.AddTimer 6000, "PlayTrack8 '"
	vpmTimer.AddTimer 6000, "GiON '"
End Sub
	

'* RAISE TARGETS *

Sub RaiseTarget1() PlaySound "fx_resetdrop"
	If Target1.IsDropped = 1 then Target1.IsDropped = 0: End If
End Sub

Sub RaiseTarget2() PlaySound "fx_resetdrop"
	If Target2.IsDropped = 1 then Target2.IsDropped = 0: End If
End Sub

Sub RaiseTarget3() PlaySound "fx_resetdrop"
	If Target3.IsDropped = 1 then Target3.IsDropped = 0: End If
	If T3Ghost.IsDropped = True then T3Ghost.IsDropped = False: End If
End Sub

Sub RaiseTarget4() PlaySound "fx_resetdrop"
	If Target4.IsDropped = 1 then Target4.IsDropped = 0: End If
	If T4Ghost.IsDropped = True then T4Ghost.IsDropped = False: End If
End Sub

Sub RaiseTarget5() PlaySound "fx_resetdrop"
	If Target5.IsDropped = 1 then Target5.IsDropped = 0: End If
	If T5Ghost.IsDropped = True then T5Ghost.IsDropped = False: End If
End Sub

Sub RaiseTarget6() PlaySound "fx_resetdrop"
	If Target6.IsDropped = 1 then Target6.IsDropped = 0: End If
	If T6Ghost.IsDropped = True then T6Ghost.IsDropped = False: End If
End Sub

Sub RaiseTarget7() PlaySound "fx_resetdrop"
	If Target7.IsDropped = 1 then Target7.IsDropped = 0: End If
End Sub

Sub RaiseTarget8() PlaySound "fx_resetdrop"
	If Target8.IsDropped = 1 then Target8.IsDropped = 0: End If
End Sub

Sub RaiseTarget9()
	If Target9.IsDropped = 1 then Target9.IsDropped = 0: End If
End Sub

'* DROP TARGETS *

Sub DropTarget1() PlaySound "fx_droptarget"
	If Target1.IsDropped = 0 then Target1.IsDropped = 1: End If
End Sub

Sub DropTarget2() PlaySound "fx_droptarget"
	If Target2.IsDropped = 0 then Target2.IsDropped = 1: End If
End Sub

Sub DropTarget3() PlaySound "fx_droptarget"
	If Target3.IsDropped = 0 then Target3.IsDropped = 1: End If
	If T3Ghost.IsDropped = False then T3Ghost.IsDropped = True: End If
End Sub

Sub DropTarget4() PlaySound "fx_droptarget"
	If Target4.IsDropped = 0 then Target4.IsDropped = 1: End If
	If T4Ghost.IsDropped = False then T4Ghost.IsDropped = True: End If
End Sub

Sub DropTarget5() PlaySound "fx_droptarget"
	If Target5.IsDropped = 0 then Target5.IsDropped = 1: End If
	If T5Ghost.IsDropped = False then T5Ghost.IsDropped = True: End If
End Sub

Sub DropTarget6() PlaySound "fx_droptarget"
	If Target6.IsDropped = 0 then Target6.IsDropped = 1: End If
	If T6Ghost.IsDropped = False then T6Ghost.IsDropped = True: End If
End Sub

Sub DropTarget7() PlaySound "fx_droptarget"
	If Target7.IsDropped = 0 then Target7.IsDropped = 1: End If
End Sub

Sub DropTarget8() PlaySound "fx_droptarget"
	If Target8.IsDropped = 0 then Target8.IsDropped = 1: End If
End Sub

Sub DropTarget9() PlaySound "fx_droptarget"
	If Target9.IsDropped = 0 then Target9.IsDropped = 1: End If
End Sub

'************
' Magnets
'************

Dim Magnet1
    Set Magnet1 = New cvpmMagnet
    With Magnet1
        .InitMagnet M1, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet1"
    End With

Sub EnableMagnet1
    Magnet1.MagnetOn = True
    Magnet1.GrabCenter = True
End Sub

Dim Magnet2
    Set Magnet2 = New cvpmMagnet
    With Magnet2
        .InitMagnet M2, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet2"
    End With
  
Sub EnableMagnet2
    Magnet2.MagnetOn = True
    Magnet2.GrabCenter = True
End Sub

Dim Magnet3
    Set Magnet3 = New cvpmMagnet
    With Magnet3
        .InitMagnet M3, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet3"
    End With

Sub EnableMagnet3
    Magnet3.MagnetOn = True
    Magnet3.GrabCenter = True
End Sub

Dim Magnet4
    Set Magnet4 = New cvpmMagnet
    With Magnet4
        .InitMagnet M4, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet4"
    End With
 
Sub EnableMagnet4
    Magnet4.MagnetOn = True
    Magnet4.GrabCenter = True
End Sub

Dim Magnet5
    Set Magnet5 = New cvpmMagnet
    With Magnet5
        .InitMagnet M5, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet5"
    End With
 
Sub EnableMagnet5
    Magnet5.MagnetOn = True
    Magnet5.GrabCenter = True
End Sub

Dim Magnet6
    Set Magnet6 = New cvpmMagnet
    With Magnet6
        .InitMagnet M6, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet6"
    End With
 
Sub EnableMagnet6
    Magnet6.MagnetOn = True
    Magnet6.GrabCenter = True
End Sub

Dim Magnet7
    Set Magnet7 = New cvpmMagnet
    With Magnet7
        .InitMagnet M7, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet7"
    End With
 
Sub EnableMagnet7
    Magnet7.MagnetOn = True
    Magnet7.GrabCenter = True
End Sub

Dim Magnet8
    Set Magnet8 = New cvpmMagnet
    With Magnet8
        .InitMagnet M8, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet8"
    End With
 
Sub EnableMagnet8
    Magnet8.MagnetOn = True
    Magnet8.GrabCenter = True
End Sub

Dim Magnet9
    Set Magnet9 = New cvpmMagnet
    With Magnet9
        .InitMagnet M9, 25
        .GrabCenter = False
        .MagnetOn = False
        .CreateEvents "Magnet9"
    End With
 
Sub EnableMagnet9
    Magnet9.MagnetOn = True
    Magnet9.GrabCenter = True
End Sub

Sub DisableMagnets()
    Magnet1.MagnetOn = False
    Magnet1.GrabCenter = False
	Magnet2.MagnetOn = False
    Magnet2.GrabCenter = False
	Magnet3.MagnetOn = False
    Magnet3.GrabCenter = False
	Magnet4.MagnetOn = False
    Magnet4.GrabCenter = False
	Magnet5.MagnetOn = False
    Magnet5.GrabCenter = False
	Magnet6.MagnetOn = False
    Magnet6.GrabCenter = False
	Magnet7.MagnetOn = False
    Magnet7.GrabCenter = False
	Magnet8.MagnetOn = False
    Magnet8.GrabCenter = False
	Magnet9.MagnetOn = False
    Magnet9.GrabCenter = False
End Sub

'************
' Gates
'************

Sub Gate1_Hit() PlaySound "fx_gate"
	FlashSequence1
	If liLock.state = 1 then PlaySound "APC": Timer2.Enabled = True: End If
	If WallBlock2.IsDropped = True then RaiseBlock2: vpmTimer.AddTimer 5000, "DropBlock2 '": End If
End Sub
	
Sub Gate2_Hit()
	If liLock.state = 0 then RaiseBlock1: RaiseBlock2: RaiseBlock5: End If
	Gate14Countb=0
	Gate14Countc=0
End Sub

Sub Gate3_Hit() PlaySound "fx_gate"
	ScannerON
	FX22Count=0
	Gate16Reset
	MusicMod2
	vpmTimer.AddTimer 1000, "DropTargetX '"
	If Controlli2.state = 0 then StopKicker7SFX: : GiON: End If
	If Controlli2.state = 1 then FlashersON: End If
End Sub

Sub Gate6_Hit() PlaySound "fx_gate": GiOFF: AlertGiOFF: SentryGunsON: End Sub

Sub Gate7_Hit() Playsound "fx_gate": GiON: SentryGunsOFF: End Sub

Dim Gate8Count
Sub Gate8_Hit() PlaySound "fx_gate"
	If SGli.state = 1 and li7.state = 1 then
	SGOffline
		If WallBlock1.IsDropped = True then WallBlock1.IsDropped = False: End If
	End If
	If SGli.state = 1 then 
		If Gate8Count = 0 then RaisePostc: PlaySound "HUD Interface 1": vpmTimer.AddTimer 2000, "DropPostc '": End If
		Gate8Count = (Gate8Count+1) mod 2
	End If
End Sub

Sub swGate8_Hit()
	If li21.state = 1 then
	PlaySound "HUD Interface 1"
	RaisePostc
	vpmTimer.AddTimer 2000, "DropPostc '"
	SGli.state = 1
		If WallBlock1.IsDropped = False then WallBlock1.IsDropped = True: End If
	End If
End Sub

Sub Gate8Reset() Gate8.Collidable = True: End Sub

Sub Gate9_Hit() PlaySound "fx_gate": StopSound "fx_metalrolling": PlaySound "Pneumatic": End Sub

Sub Gate11_Hit() PlaySound "fx_gate": PlaySound "fx_metalrolling": End Sub

Sub Gate12_Hit() PlaySound "fx_gate": PlaySound "Door": End Sub

Sub Gate16Reset() Gate16.Collidable = True: End Sub

Dim Gate14Counta, Gate14Countb, Gate14Countc
Sub Gate14_Hit()
	If Controlli2.state = 1 then
		SFX23_Hit
	Else
		If Gate14Counta = 5 then PlaySound "Loader": GiOFF: RaisePosts: FFXd: vpmTimer.AddTimer 900,"GiON '": End If
		If Gate14Counta = 11 then GiOFF: AlertGiON: RaisePosts: vpmTimer.AddTimer 16000,"AlertGiOFF '": vpmTimer.AddTimer 17000,"GiON '": End if 
		Gate14Counta = (Gate14Counta+1) mod 12

		If Gate14Countb = 2 then PlaySFXa: BlueGiFlashSequence: DropBlock2: DropBlock5: End If
		Gate14Countb = (Gate14Countb+1) mod 50

		If Gate14Countc = 6 and SGli.state = 0 then
		SGOnline
		If WallBlock1.IsDropped = False then WallBlock1.IsDropped = True: End If
		vpmTimer.AddTimer 1000, "PlaySFXa '"
		Gate14Countc = (Gate14Countc+1) mod 50
		End If
	End If
End Sub

Sub BlueGiFlashSequence()
	BlueGiFlash
	vpmTimer.AddTimer 400, "BlueGiFlash '"
	vpmTimer.AddTimer 800, "BlueGiFlash '"
	vpmTimer.AddTimer 1200, "BlueGiFlash '"
End Sub

Sub Gate15_Hit() ScannerOFF: End Sub

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "ALIENS 3.0 - Original 2022" & vbNewLine & "VPX table by Delta23"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 1 'VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
		.Games(cGameName).Settings.Value("sound") = 0
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With
	
    ' Nudging
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .Initsw 28, 30, 29, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 2
    End With
	
    ' Saucers
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw37, 37, 180, 20
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickZ = 1
        .KickForceVar = 2
    End With

    ' Drop targets
    Set dtBank = New cvpmDropTarget
    With dtBank
        .InitDrop Array(sw34, sw35, sw36), Array(34, 35, 36)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBank"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1: PlaySound "Combat Drop"
	Timer1.Enabled = False

	' Drop Walls
	StopPosta.IsDropped = 1
	StopPostb.IsDropped = 1
	StopPostc.IsDropped = 1
	StopPostd.IsDropped = 1
	WallBlock1.IsDropped = 1
	WallBlock2.IsDropped = 1
	WallBlock3.IsDropped = 1
	WallBlock4.IsDropped = 1
	WallBlock5.IsDropped = 1
	WallBlock6.IsDropped = 1
	LeftPostHold.IsDropped = 1
	RightPostHold.IsDropped = 1
	T3Ghost.IsDropped = 1
	T4Ghost.IsDropped = 1
	T5Ghost.IsDropped = 1
	T6Ghost.IsDropped = 1

	' Captive Ball
	CaptiveBall.CreateBall
	CaptiveBall.Kick 90, 1
	CaptiveBall.enabled = 0
End Sub

Sub table1_Paused: Controller.Pause = 1: End Sub
Sub table1_unPaused: Controller.Pause = 0: End Sub
Sub table1_exit: Controller.stop: End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = RightFlipperKey Then Controller.Switch(44) = 1
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = RightFlipperKey Then Controller.Switch(44) = 0
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep, LeftCount, RightCount

Sub LeftSlingShot_Slingshot
	If Controlli2.state = 1 then PulseRifle: PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05:End If
	If Controlli2.state = 0 then PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05:End If
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 32
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSling4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
	If Controlli2.state = 1 then PulseRifle: PlaySound SoundFX ("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05: End If
	If Controlli2.state = 0 then PlaySound SoundFX ("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05: End If
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 33
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Rubbers, sound is done in the collection

Sub sw39_Hit: vpmTimer.PulseSw 39: End Sub
Sub sw40_Hit: vpmTimer.PulseSw 40: End Sub
Sub sw42_Hit: vpmTimer.PulseSw 42: End Sub
Sub sw43_Hit: vpmTimer.PulseSw 43: End Sub

' Bumpers
Sub Bumper1_Hit()
	If Controlli2.state = 0 then
	vpmTimer.PulseSw 21: PlaySound SoundFX ("fx_bumper", DOFContactors), 0, 1, -0.05
	Else
	PulseRifle
	End If
End Sub
Sub Bumper2_Hit()
	If Controlli2.state = 0 then
	vpmTimer.PulseSw 23: PlaySound SoundFX ("fx_bumper", DOFContactors), 0, 1, -0.025
	Else
	PulseRifle
	End If
End Sub
Sub Bumper3_Hit()
	If Controlli2.state = 0 then
	vpmTimer.PulseSw 22: PlaySound SoundFX ("fx_bumper", DOFContactors), 0, 1, -0.05
	Else
	PulseRifle
	End If
End Sub
Sub SubBumper1_Hit:vpmTimer.PulseSw 21: PlaySound "Sentry gun 1": PlaySound SoundFX ("fx_bumper", DOFContactors), 0, 1, -0.05: SentryGunsFlashON: End Sub
Sub SubBumper2_Hit:vpmTimer.PulseSw 21: PlaySound "Sentry gun 1": PlaySound SoundFX ("fx_bumper", DOFContactors), 0, 1, -0.05: SentryGunsFlashON: End Sub
Sub SubBumper3_Hit:vpmTimer.PulseSw 22: PlaySound "Sentry gun 1": PlaySound SoundFX ("fx_bumper", DOFContactors), 0, 1, -0.05: SentryGunsFlashON: End Sub
Sub SubBumper4_Hit:vpmTimer.PulseSw 22: PlaySound "Sentry gun 1": PlaySound SoundFX ("fx_bumper", DOFContactors), 0, 1, -0.05: SentryGunsFlashON: End Sub

' Drain & Saucers

Sub DrainHold_Hit() Playsound "fx_drain": StopSound "Loader"
	DrainHold.TimerInterval = 5000
	DrainHold.TimerEnabled = 1
	Timer1.Enabled = False
	EndMusic 
	AlertGiOFF
	BlueGiOFF
	FX9Countb=0
	FX21Count=0
	FX10Count=1
	FX11Count=1
	FX18Count=1
	FX19Count=1
	RandomTargetOffline
	Controlli5.state = 0
	Controlli9.state = 0
	vpmTimer.AddTimer 1100, "GiOFF '"
		If Controlli2.state = 1 then
			MBLost
		else
			BLost
		End If
End Sub

Sub DrainHold_Timer
	DrainHold.Kick 180, 2
	DrainHold.TimerEnabled = 0
End Sub

Sub MBLost() PlaySound "Random 2"
	vpmTimer.AddTimer 6000, "DropBlock2 '"
	vpmTimer.AddTimer 6000, "DropBlock4 '"
	vpmTimer.AddTimer 6000, "DropBlock5 '"
	Controlli2.state = 0
	vpmTimer.AddTimer 1000, "DATMod2 '"
	vpmTimer.AddTimer 2000, "PlayTrack9 '"
	vpmTimer.AddTimer 5500, "GiON '"
	If StopPosta.IsDropped = 0 then vpmTimer.AddTimer 10000, "DropPosta '": End If
	If StopPostb.IsDropped = 0 then vpmTimer.AddTimer 10000, "DropPostb '": End If
End Sub	

Dim BLostCount
Sub BLost() DATMod3
	If Controlli3.state = 0 then
		If BLostCount = 0 then PlaySound "Snippet 2": End If
		If BLostCount = 1 then PlaySound "Snippet 1": End If
		BLostCount = (BLostCount+1) mod 2
		vpmTimer.AddTimer 5000, "DataLoadON '" 
		vpmTimer.AddTimer 1000, "PlaySFXb '" 
		If StopPosta.IsDropped = 0 then vpmTimer.AddTimer 3000, "DropPosta '": End If
		If StopPostb.IsDropped = 0 then vpmTimer.AddTimer 3000, "DropPostb '": End If
	End If
	If Controlli3.state = 1 then
	vpmTimer.AddTimer 1500, "StopSFXb '"
	vpmTimer.AddTimer 2000, "MusicMod1 '"
	Controlli3.state = 0
	End If
End Sub

Sub Drain_Hit() bsTrough.AddBall Me: End Sub

Sub sw37_Hit: PlaySound "Hydraulic Door": PlaySound "fx_kicker_enter", 0, 1, 0.05: bsSaucer.AddBall 0
	Controlli2.state = 1
	Controlli9.state = 1
	vpmTimer.AddTimer 500, "LockLight '"	
End Sub

Sub LockLight() liLock.state = 1: PlaySound "HUD Interface 7": End Sub

' Rollovers
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):PlaySound "fx_metalhit2", 0, 0.1, 0.05:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1: PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1: PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Dim sw26Count
Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "HUD Interface 2":PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	If Controlli2.state = 0 and sw26Count = 0 then
		If Controlli5.state = 1 then AlertGiOFF: RaiseLeftHold: vpmTimer.AddTimer 7000, "DropLeftHold '": sw26Count = (sw26Count+1) mod 2: End If
	End If
End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0: End Sub

Dim sw27Count
Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "HUD Interface 2":PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	If Controlli2.state = 0 and sw27Count = 0 then
		If Controlli5.state = 1 then AlertGiOFF: RaiseRightHold: vpmTimer.AddTimer 7000, "DropRightHold '": sw27Count = (sw27Count+1) mod 2: End If
	End If
End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "HUD Interface 2":PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "HUD Interface 2":PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySound "HUD Interface 2":PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySound "HUD Interface 2":PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

' Droptargets (sound only)
Sub sw34_Hit: PlaySound "Alien 3":PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw35_Hit: PlaySound "Alien 4":PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw36_Hit: PlaySound "Alien 10":PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

' Spinners
Sub sw9_Spin:vpmTimer.PulseSw 9:PlaySound "fx_spinner", 0, 1, -0.05:End Sub
Sub sw16_Spin:vpmTimer.PulseSw 16:PlaySound "fx_spinner", 0, 1, 0.05:End Sub

'Targets
Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySound "HUD Interface 7":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11:PlaySound "HUD Interface 7":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySound "HUD Interface 7":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySound "HUD Interface 7":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):Flasher6ON:End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySound "HUD Interface 7":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15:PlaySound "HUD Interface 7":PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

'*********
'Solenoids
'*********

Sub flArray()
	liArray = Array(F5,F5a,F6,F6a,F7,F7a)
	For a = 0 to 5
	liArray(a).State=0
	Next
End Sub

SolCallback(1) = "bsTrough.SolIn"   ' Outhole
SolCallback(2) = "bsTrough.SolOut"  ' Ball Release
SolCallback(3) = "bsSaucer.SolOut"  ' Multi-Ball Eject
SolCallback(4) = "dtBank.SolDropUp" ' Three Bank Reset
SolCallback(5) = "SetLamp 105,"     ' USC Marines Flasher
SolCallback(6) = "SetLamp 106," 	' Drop Target Flasher
SolCallback(7) = "SetLamp 107,"     ' Center Target Bank Flashers
SolCallback(8) = "SetLamp 108,"     ' Back Panel Flashers
SolCallback(11) = "SolGi"           ' General Illumination
'SolCallback(14)     = ""
SolCallback(15) = "vpmSolSound""Bleep 4""," ' Bell
'SolCallback(16)
'SolCallback(17)     = "" 				' Left Sling
'SolCallback(18)     = "" 				' Right Sling
'SolCallback(19)	  = ""  			' Left Jet Bumper
'SolCallback(20)	  = ""    			' Bottom Jet Bumper
'SolCallback(21)	  = ""    			' Right Jet Bumper
SolCallback(23) = "vpmNudge.SolGameOn"

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then 
		PlaySound SoundFX ("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.05
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
		If Controlli1.state = 1 and Controlli2.state = 0 then LeftFlipper2.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.05
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
		LeftFlipper2.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
		PlaySound SoundFX ("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.05
        RightFlipper.RotateToEnd
		RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.05
        RightFlipper.RotateToStart
		RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

'*****************
'   Gi Effects
'*****************

Sub SolGi(enabled)
   If Enabled Then
		PlaySound "fx_SolenoidOn", 0, 0.1
        GiOFF
    Else
		PlaySound "fx_SolenoidOff", 0, 0.1
        GiON
    End If
End Sub

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Sub PFflasher
	For each x in aPFlights
		x.duration 2,500,0
	Next
End Sub

Sub AlertGiON() PlaySound "Emergency alert"
	Timer1.Enabled = False
	Controlli8.state = 1
	For each x in aRedlights
		x.State = 2
	Next
End Sub

Sub AlertGiOFF() StopSound "Emergency alert"
	Controlli8.state = 0
	For each x in aRedlights
		x.State = 0
	Next
End Sub

Sub RedGiON
	For each x in aRedlights
		x.state = 1
	Next
End Sub

Sub RedGiOFF
	For each x in aRedlights
		x.State = 0
	Next
End Sub

Sub RedGiFlash
	For each x in aRedlights
		x.duration 2,250,0
	Next
End Sub

Sub GreenGiON
	For each x in aGreenlights
		x.state = 1
	Next
End Sub

Sub GreenGiOFF
	For each x in aGreenlights
		x.State = 0
	Next
End Sub

Sub GreenGiFlash
	For each x in aGreenlights
		x.duration 2,250,0
	Next
End Sub

Sub BlueGiON
	For each x in aBluelights
		x.State = 1
	Next
End Sub

Sub BlueGiOFF
	For each x in aBluelights
		x.State = 0
	Next
End Sub

Sub BlueGiFlash
	For each x in aBluelights
		x.duration 2,250,0
	Next
End Sub 

Sub FlashersON
	For each x in aFlashers
		x.duration 1,200,0
	Next
End Sub

Sub FlashersON_mod
	For each x in aFlashers
		x.duration 2,900,0
	Next
End Sub

Sub FlashersON_mod2
	For each x in aFlashers
		x.duration 1,500,0
	Next
End Sub

Sub Flasher5ON
	For each x in Flasher5
		x.duration 1,200,0
	Next
End Sub

Sub Flasher6ON
	For each x in Flasher6
		x.duration 1,200,0
	Next
End Sub

Sub Flasher7ON
	For each x in Flasher7
		x.duration 1,200,0
	Next
End Sub

Sub FlashSequence1
	vpmTimer.AddTimer 100, "Flasher7ON '"
	vpmTimer.AddTimer 200, "Flasher5ON '"
	vpmTimer.AddTimer 300, "Flasher6ON '"
End Sub

Sub FlashSequence2
	vpmTimer.AddTimer 100, "Flasher6ON '"
	vpmTimer.AddTimer 200, "Flasher5ON '"
	vpmTimer.AddTimer 300, "Flasher7ON '"
End Sub

Sub FlashSequence3
	vpmTimer.AddTimer 100, "Flasher5ON '"
	vpmTimer.AddTimer 200, "Flasher6ON '"
	vpmTimer.AddTimer 300, "Flasher7ON '"
End Sub

Sub GalleryGiON
	For each x in aGalleryLights
		x.State = 1
	Next
End Sub

Sub GalleryGiOFF
	For each x in aGalleryLights
		x.State = 0
	Next
End Sub

Sub TargetLightON
	If Targetli.state = 0 then Targetli.state = 1: End If
End Sub

Sub TargetLightOFF
	If Targetli.state = 1 then Targetli.state = 0: End If
End Sub

Sub SentryGunsON
	Consoleli.state = 1
	Sentryli.state = 1
End Sub

Sub SentryGunsOFF
	Consoleli.state = 0
	Sentryli.state = 0
End Sub

Sub SentryLightON() Sentryli.state = 2: End Sub

Sub SentryLightOFF() Sentryli.state = 0: End Sub

Sub SentryGunsFlashON
	For each x in SentryFlash
		x.duration 1,100,0
	Next
End Sub

Sub RandomTargetOnline
	PlaySound "HUD Interface 7"
	liA.state = 2
	liB.state = 2
	liC.state = 1
End Sub

Sub RandomTargetOffline
	liA.state = 0
	liB.state = 0
	liC.state = 0
End Sub

Sub F10Flash
	F10.duration 1,100,0 
End Sub
Sub F11Flash
	F11.duration 1,100,0
End Sub
Sub F12Flash
	F12.duration 1,100,0
End Sub
Sub F13Flash
	F13.duration 1,100,0
End Sub
Sub F14Flash
	F14.duration 1,100,0
End Sub
Sub F15Flash
	F15.duration 1,100,0
End Sub
Sub F16Flash
	F16.duration 1,100,0
End Sub
Sub F17Flash
	F17.duration 1,100,0
End Sub

Sub FFXa 'Clockwise
	vpmTimer.AddTimer 100, "F10Flash '"
	vpmTimer.AddTimer 100, "F11Flash '"
	vpmTimer.AddTimer 300, "F12Flash '"
	vpmTimer.AddTimer 300, "F13Flash '"
	vpmTimer.AddTimer 500, "F14Flash '"
	vpmTimer.AddTimer 500, "F15Flash '"
	vpmTimer.AddTimer 700, "F16Flash '"
	vpmTimer.AddTimer 700, "F17Flash '"
End Sub

Sub FFXb 'Down
	vpmTimer.AddTimer 100, "F11Flash '"
	vpmTimer.AddTimer 100, "F12Flash '"
	vpmTimer.AddTimer 300, "F10Flash '"
	vpmTimer.AddTimer 300, "F13Flash '"
	vpmTimer.AddTimer 500, "F17Flash '"
	vpmTimer.AddTimer 500, "F14Flash '"
	vpmTimer.AddTimer 700, "F16Flash '"
	vpmTimer.AddTimer 700, "F15Flash '"
End Sub

Sub FFXc 'Anti-Clockwise
	vpmTimer.AddTimer 100, "F13Flash '"
	vpmTimer.AddTimer 100, "F12Flash '"
	vpmTimer.AddTimer 300, "F11Flash '"
	vpmTimer.AddTimer 300, "F10Flash '"
	vpmTimer.AddTimer 500, "F17Flash '"
	vpmTimer.AddTimer 500, "F16Flash '"
	vpmTimer.AddTimer 700, "F15Flash '"
	vpmTimer.AddTimer 700, "F14Flash '"
End Sub	

Sub FFXd 'Up
	vpmTimer.AddTimer 100, "F15Flash '"
	vpmTimer.AddTimer 100, "F16Flash '"
	vpmTimer.AddTimer 300, "F14Flash '"
	vpmTimer.AddTimer 300, "F17Flash '"
	vpmTimer.AddTimer 500, "F10Flash '"
	vpmTimer.AddTimer 500, "F13Flash '"
	vpmTimer.AddTimer 700, "F11Flash '"
	vpmTimer.AddTimer 700, "F12Flash '"
End Sub	

Sub FFXe 'Random
	vpmTimer.AddTimer 100, "F10Flash '"
	vpmTimer.AddTimer 200, "F12Flash '"
	vpmTimer.AddTimer 300, "F11Flash '"
	vpmTimer.AddTimer 400, "F13Flash '"
	vpmTimer.AddTimer 500, "F17Flash '"
	vpmTimer.AddTimer 600, "F15Flash '"
	vpmTimer.AddTimer 700, "F16Flash '"
	vpmTimer.AddTimer 800, "F14Flash '"
End Sub
	
'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 ' lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0)) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    If VarHidden Then
        UpdateLeds
    End If
    UpdateLamps
    RollingUpdate
End Sub

Sub UpdateLamps()

    'backdrop lights
    NFadeT 1, li1, "Game Over"
    NFadeT 2, li2, "Match"
    NFadeT 3, li3, "TILT"
    NFadeTm 4, li4, "High Score"
    NFadeT 4, li4a, "To Date"
    NFadeT 5, li5, "Shoot Again"
    NFadeT 6, li6, "Ball in Play"

    ' playfield lights
    NFadeL 6, li6
    NFadeL 7, li7
    NFadeL 8, li8
    NFadeL 9, li9
    NFadeL 10, li10
    NFadeL 11, li11
    NFadeL 12, li12
    NFadeL 13, li13
    NFadeL 14, li14
    NFadeL 15, li15
    NFadeL 16, li16
    NFadeL 17, li17
    NFadeL 18, li18
    NFadeL 19, li19
    NFadeL 20, li20
    NFadeL 21, li21
    NFadeL 22, li22
    NFadeL 23, li23
    NFadeL 24, li24
    NFadeL 25, li25
    NFadeL 26, li26
    NFadeL 27, li27
    NFadeL 28, li28
    NFadeL 29, li29
    NFadeL 30, li30
    NFadeL 31, li31
    NFadeL 32, li32
    NFadeL 33, li33
    NFadeL 34, li34
    NFadeL 35, li35
    NFadeL 36, li36
    NFadeL 37, li37
    NFadeL 38, li38
    NFadeL 39, li39
    NFadeL 40, li40
    NFadeL 41, li41
    NFadeL 42, li42
    NFadeL 43, li43
    NFadeL 44, li44
    NFadeL 45, li45
    NFadeL 46, li46
    NFadeL 47, li47
    NFadeL 48, li48
    NFadeL 49, li49
    NFadeL 50, li50
    NFadeL 51, li51
    NFadeL 52, li52
    NFadeL 53, li53
    NFadeL 54, li54
    NFadeL 55, li55
	
    'flashers
    NFadeL 105, F5a
	NFadeL 105, F5
    NFadeL 106, F6a
    NFadeL 106, F6
	NFadeL 106, F7a
    NFadeL 106, F7
	NFadeL 106, F5a
    NFadeL 106, F5
    NFadeL 107, F7a
    NFadeL 107, F7
    Flashm 108, F8a
    Flash 108, F8b
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.5   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.25 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr)Then
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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr)Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr)-1
                If FlashRepeat(nr)Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr)Then FadingLevel(nr) = 4
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 6-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6

Set Digits(7) = b0
Set Digits(8) = b1
Set Digits(9) = b2
Set Digits(10) = b3
Set Digits(11) = b4
Set Digits(12) = b5
Set Digits(13) = b6

Set Digits(14) = c0
Set Digits(15) = c1
Set Digits(16) = c2
Set Digits(17) = c3
Set Digits(18) = c4
Set Digits(19) = c5
Set Digits(20) = c6

Set Digits(21) = d0
Set Digits(22) = d1
Set Digits(23) = d2
Set Digits(24) = d3
Set Digits(25) = d4
Set Digits(26) = d5
Set Digits(27) = d6

Set Digits(28) = e0
Set Digits(29) = e1
Set Digits(30) = e2
Set Digits(31) = e3

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Dim MetalHit
Sub aMetals_Hit(idx) MetalHit=Int(rnd*2)
	If MetalHit = 0 then PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End If
	If MetalHit = 1 then PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End If
End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRamps_Hit(idx):PlaySound "fx _metalrolling", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Dim TableWidth, TableHeight
TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng((ball.VelX*ball.VelX + ball.VelY*ball.VelY) / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp> 0 Then
        Dim pn2: pn2 = tmp*tmp: Dim pn4: pn4 = pn2*pn2: Dim pn8: pn8 = pn4*pn4
        Pan = Csng(pn8*pn2)
    Else
        Dim pnt: pnt = -tmp: Dim pnt2: pnt2 = pnt*pnt: Dim pnt4: pnt4 = pnt2*pnt2: Dim pnt8: pnt8 = pnt4*pnt4
        Pan = Csng(-(pnt8*pnt2))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX*ball.VelX) + (ball.VelY*ball.VelY)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        Dim af2: af2 = tmp*tmp: Dim af4: af4 = af2*af2: Dim af8: af8 = af4*af4
        AudioFade = Csng(af8*af2)
    Else
        Dim aft: aft = -tmp: Dim aft2: aft2 = aft*aft: Dim aft4: aft4 = aft2*aft2: Dim aft8: aft8 = aft4*aft4
        AudioFade = Csng(-(aft8*aft2))
    End If
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 6 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub
Dim BallRollStr
BallRollStr = Array("fx_ballrolling0", "fx_ballrolling1", "fx_ballrolling2", "fx_ballrolling3", "fx_ballrolling4", "fx_ballrolling5", "fx_ballrolling6")


Dim BallDrop
Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound(BallRollStr(b))
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)

		aBallShadow(b).X = BOT(b).X
		aBallShadow(b).Y = BOT(b).Y

        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 15000 'increase the pitch on a ramp or elevated surface
            End If
            rolling(b) = True
            PlaySound(BallRollStr(b)), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0
        Else
            If rolling(b) = True Then
                StopSound(BallRollStr(b))
                rolling(b) = False
            End If
        End If
			If BOT(b).VelZ < -2.3 and BOT(b).z < 56 and BOT(b).z > 30 Then 'height adjust for ball drop sounds
				PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 ', ABS(BOT(b).velz)/17, 0, 1, 0: End If
		End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity*velocity) / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

Sub SFX17_Init()
	
End Sub

Sub SFX17_Timer()
	
End Sub

Sub SFX17_Unhit()
	
End Sub