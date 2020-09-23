VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Space Quest"
   ClientHeight    =   4350
   ClientLeft      =   1965
   ClientTop       =   1875
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   Begin VB.Timer VolShow 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3000
      Top             =   2760
   End
   Begin VB.Timer VolChange 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   2760
   End
   Begin VB.Timer HealthRegen 
      Interval        =   1000
      Left            =   3600
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3960
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3360
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2520
      Top             =   1920
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'gets the key states Up or Down
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'gets input from the keyboard
Dim DI As DirectInput
Dim diDevice As DirectInputDevice

'used to create the direct sound buffers
Dim DS As DirectSound
Dim Perform As DirectMusicPerformance
Dim Loader As DirectMusicLoader
Dim Segment As DirectMusicSegment
Dim SegmentState As DirectMusicSegmentState
Dim VolumeCan As Boolean
'holds basic information about the sprites
Private Type Sprite
'holds the current X position
    X As Long
'holds the current Y position
    Y As Long
'holds the current X movement
    Xp As Long
'holds the current Y movement
    Yp As Long
'holds the color
    Color As Long
'holds the current phase
    Phase As Long
'determines whether the sprite is active
    Active As Boolean
'holds how many points the sprite is worth
    Worth As Long
'holds how much damage it will do
    Damage As Long
'determines how many frames pass between phases
    Spin As Long
End Type
'holds volume bitmap
Dim Vol As DirectDrawSurface7
'holds vol height, width and phases
Const VOLUMEWIDTH = 5
Const VOLUMEHEIGHT = 21
Const VOLUMEPHASES = 0
Const VOLUMECHANGE = 1
Const MAXVOLUME = 1000
Const MINVOLUME = -3000
'holds the background musics volume
Dim Volume As Integer
'maximum number of frames between phases
Const MAXSPIN = 5
'holds the current frame number
Dim Frame As Long
'holds the font height, width and phases
Const FONTWIDTH = 22
Const FONTHEIGHT = 22
Const FONTPHASES = 37
Dim Fonts(FONTPHASES) As Sprite
'holds the players score
Dim PlayerScore As Double
'holds the players health
Dim PlayerHealth As Long
'determines whether gun has reloaded
Dim CanShoot As Boolean
'holds whether the games is paused or not
Dim Paused As Boolean
'whether enough time has passed to unpause
Dim CanUnPause As Boolean
'holds the maximum number of bullets you can shoot
Const NUMBULLETS = 10
'holds bullets frame height
Const BULLETHEIGHT = 22
'holds bullets frame width
Const BULLETWIDTH = 10
'holds the number of bullet frames
Const BULLETPHASES = 8
'holds ships frame width
Const SHIPWIDTH = 67
'holds ships frame height
Const SHIPHEIGHT = 22
'holds ships number of phases
Const SHIPPHASES = 3
'holds where the gun shoots from
Const GUNX = 50
Const GUNY = 0
'contains the explosion height, width and
'number of frames
Const EXPLOSIONHEIGHT = 35
Const EXPLOSIONWIDTH = 35
Const EXPLOSIONPHASES = 19
'contains the asteroids height, width and
'number of frames
Const ASTEROIDHEIGHT = 25 '75
Const ASTEROIDWIDTH = 25 '75
Const ASTEROIDPHASES = 29
'maximum number of asteroids
Const NUMASTEROIDS = 25
'number of ship collision sounds can play
Const NUMSHIPHIT = 5
'holds the guns sound
Dim GunSound(NUMBULLETS - 1) As DirectSoundBuffer
'holds the explosion sound
Dim ExplosionSound(NUMASTEROIDS - 1) As DirectSoundBuffer
'holds the ship collision sound
Dim ShipHit(NUMSHIPHIT - 1) As DirectSoundBuffer
'holds the explosion sprite info
Dim Explosions(NUMASTEROIDS - 1) As Sprite
'holds the explosion bitmap
Dim ExplosionsS As DirectDrawSurface7
'holds the asteroids sprite info
Dim Asteroids(NUMASTEROIDS - 1) As Sprite
'holds the asteroids bitmap
Dim AsteroidsS As DirectDrawSurface7
'tells when program is done
Dim RUNNING As Boolean
'holds Bullets sprite info
Dim Bullets(NUMBULLETS - 1) As Sprite
'holds bullets bitmap
Dim BulletsS As DirectDrawSurface7
'holds ships position
Dim ShipX As Long
Dim ShipY As Long
'holds ships current frame
Dim ShipPhase As Long
'main directx object
Dim DX As New DirectX7
'main directdraw object
Dim DD As DirectDraw7
'screen that the user will see
Dim Primary As DirectDrawSurface7
'area in memory that drawing will take place on
Dim Backbuffer As DirectDrawSurface7
'area in memory that hold the ships phases
Dim Ship As DirectDrawSurface7
'area in memory that holds the font
Dim sFont As DirectDrawSurface7
'these just hold information about the surfaces
Dim DDSD1 As DDSURFACEDESC2
Dim DDSD2 As DDSURFACEDESC2
Dim DDSD3 As DDSURFACEDESC2
Dim DDCaps As DDSCAPS2
'holds the transparent color
Dim cKey As DDCOLORKEY
'holds the applications path
Dim aPath As String

Private Sub Form_Click()
'closes out the form
RUNNING = False
End Sub

Private Sub Form_Load()
VolumeCan = True
'create the direct input object
Set DI = DX.DirectInputCreate()
'creates the device object
Set diDevice = DI.CreateDevice("GUID_sysKeyboard")
'sets it to read the keyboard
diDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
'sets it to act with our application
diDevice.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND + DISCL_NONEXCLUSIVE
'allows reading of keys
diDevice.Acquire
'starts volume out at max
Volume = 100
'makes sure that the frame is 0
Frame = 0
'players sheild/health
PlayerHealth = 100
'sets the spacing between each font letter
SetFontWidths
'sets ships current x y postion
ShipX = 0
ShipY = 300
'gets the applications path
aPath = App.Path
'makes sure path ends with '\'
If Right(aPath, 1) <> "\" Then aPath = aPath & "\"
'creates the directdraw object
Set DD = DX.DirectDrawCreate("")
Set DS = DX.DirectSoundCreate("")
DS.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
'sets the drawing to take place on the form
DD.SetCooperativeLevel Me.hWnd, DDSCL_ALLOWMODEX + DDSCL_ALLOWREBOOT + DDSCL_FULLSCREEN + DDSCL_EXCLUSIVE
'changes the current screen mode
DD.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT
'holds info about primary surface
DDSD1.lFlags = DDSD_CAPS + DDSD_BACKBUFFERCOUNT
DDSD1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE + DDSCAPS_FLIP + DDSCAPS_COMPLEX + DDSCAPS_VIDEOMEMORY
'sets the number of buffers
DDSD1.lBackBufferCount = 1
'creates the primary surface from settings
Set Primary = DD.CreateSurface(DDSD1)
'sets info about backbuffer
DDCaps.lCaps = DDSCAPS_BACKBUFFER
'creates the backbuffer from info and attaches
'it to primary surface
Set Backbuffer = Primary.GetAttachedSurface(DDCaps)
'set the info for the ships surface
Backbuffer.GetSurfaceDesc DDSD1
'sets the font color
Backbuffer.SetForeColor RGB(255, 255, 0)
Backbuffer.SetFontBackColor vbWhite
Backbuffer.SetFontTransparency True
DDSD3.lFlags = DDSD_CAPS + DDSD_HEIGHT + DDSD_WIDTH
DDSD3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
DDSD3.lWidth = SHIPWIDTH * (SHIPPHASES + 1)
DDSD3.lHeight = SHIPHEIGHT
Set Ship = DD.CreateSurfaceFromFile(aPath & "shipa.bmp", DDSD3)
cKey.high = RGB(255, 255, 0)
cKey.low = RGB(255, 255, 0)
Ship.SetColorKey DDCKEY_SRCBLT, cKey
'sets the area to fill the backbuffers color
Dim I As Long
Dim bufferDesc As DSBUFFERDESC
Dim waveFormat As WAVEFORMATEX
bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
waveFormat.nFormatTag = WAVE_FORMAT_PCM
waveFormat.nChannels = 2
waveFormat.lSamplesPerSec = 22050
waveFormat.nBitsPerSample = 16
waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
For I = 0 To NUMBULLETS - 1
Set GunSound(I) = DS.CreateSoundBufferFromFile(aPath & "fire2.wav", bufferDesc, waveFormat)
Next I
For I = 0 To NUMSHIPHIT - 1
Set ShipHit(I) = DS.CreateSoundBufferFromFile(aPath & "hit.wav", bufferDesc, waveFormat)
Next I
Randomize Timer
For I = 0 To NUMASTEROIDS - 1
Set ExplosionSound(I) = DS.CreateSoundBufferFromFile(aPath & "exp" & (Int(Rnd * 4) + 1) & ".wav", bufferDesc, waveFormat)
Next I

Set Loader = DX.DirectMusicLoaderCreate()
Set Perform = DX.DirectMusicPerformanceCreate()
Call Perform.Init(Nothing, 0)
Perform.SetPort -1, 80
Call Perform.GetMasterAutoDownload
Loader.SetSearchDirectory aPath
Perform.SetMasterAutoDownload True
Set Segment = Loader.LoadSegment(aPath & "main2.mid")
Segment.SetStandardMidiFile
Segment.SetStartPoint (0)
Set SegmentState = Perform.PlaySegment(Segment, 0, 0)

DDSD3.lWidth = BULLETWIDTH * (BULLETPHASES + 1)
DDSD3.lHeight = BULLETHEIGHT
Set BulletsS = DD.CreateSurfaceFromFile(aPath & "b6.bmp", DDSD3)
BulletsS.SetColorKey DDCKEY_SRCBLT, cKey
DDSD3.lWidth = ASTEROIDWIDTH * (ASTEROIDPHASES + 1)
DDSD3.lHeight = ASTEROIDHEIGHT
Set AsteroidsS = DD.CreateSurfaceFromFile(aPath & "ast.bmp", DDSD3)
AsteroidsS.SetColorKey DDCKEY_SRCBLT, cKey
DDSD3.lWidth = EXPLOSIONWIDTH * (EXPLOSIONPHASES + 1)
DDSD3.lHeight = EXPLOSIONHEIGHT
Set ExplosionsS = DD.CreateSurfaceFromFile(aPath & "exp.bmp", DDSD3)
ExplosionsS.SetColorKey DDCKEY_SRCBLT, cKey
'sets the width of the font
DDSD3.lWidth = FONTWIDTH * (FONTPHASES + 1)
'sets the height of the font
DDSD3.lHeight = FONTHEIGHT
'loads in the bitmap
Set sFont = DD.CreateSurfaceFromFile(aPath & "text.bmp", DDSD3)
'sets the invisible color to white
sFont.SetColorKey DDCKEY_SRCBLT, cKey
'sets the width of the volume bars
DDSD3.lWidth = VOLUMEWIDTH * (VOLUMEPHASES + 1)
'sets the height of the volume bars
DDSD3.lHeight = VOLUMEHEIGHT
'loads the bitmap in
Set Vol = DD.CreateSurfaceFromFile(aPath & "vlm.bmp", DDSD3)
'sets the invisible color to white
Vol.SetColorKey DDCKEY_SRCBLT, cKey

Dim rSurf As RECT
rSurf.Top = 0
rSurf.Bottom = 600
rSurf.Left = 0
rSurf.Right = 800
RUNNING = True
Dim N As Long
If NUMBULLETS > NUMASTEROIDS Then
N = NUMBULLETS
Else
N = NUMASTEROIDS
End If
'sets the volume of the background music
Perform.SetMasterVolume 1000
'makes sure game starts unpaused
Paused = False
CanUnPause = False
Do
Frame = Frame + 1
If Frame > MAXSPIN Then Frame = 0
'CheckKeys
ReadKeys
DoEvents
'sets the background color to black
Backbuffer.BltColorFill rSurf, RGB(0, 0, 0)
'draws the ship
If PlayerHealth <= 100 Then ShipPhase = 0
If PlayerHealth <= 75 Then ShipPhase = 1
If PlayerHealth <= 50 Then ShipPhase = 2
If PlayerHealth <= 25 Then ShipPhase = 3
DrawShip ShipPhase, ShipX, ShipY
DoEvents
For I = 0 To N - 1
If Perform.IsPlaying(Segment, SegmentState) = False Then Segment.SetStartPoint (0): Perform.PlaySegment Segment, 0, 0

DoEvents
If I <= NUMBULLETS - 1 Then If Bullets(I).Active = True Then TestBulletCollision I
If I <= NUMBULLETS - 1 Then Call UpdateBullet(I, Bullets(I).Phase)
If I <= NUMASTEROIDS - 1 Then TestCollision I: UpdateAsteroids I: Explode I
Next I
On Error Resume Next
If PlayerScore < 0 Then PlayerScore = 0
If PlayerHealth < 0 Then PlayerHealth = 0
DrawHealth
DrawScore PlayerScore, 0, 0, 2
If Paused = True Then Call DrawString("Paused", 350, 289, 3)
If VolShow.Enabled = True Then DrawVolume 0, 0, 570
Call Primary.Flip(Nothing, DDFLIP_WAIT)
Loop Until RUNNING = False
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'restores the computers original settings
DD.RestoreAllSurfaces
'destroys the directdraw object
Set DD = Nothing
'destroys the directx object
Set DX = Nothing
End Sub

Public Sub DrawShip(ByVal Phase As Long, ByVal X As Long, ByVal Y As Long)
'draws the ship at the current location
'holds the area of the sprite to clip
Dim rMain As RECT
Dim rDest As RECT
rMain.Top = 0
rMain.Left = Phase * SHIPWIDTH
rMain.Right = rMain.Left + SHIPWIDTH
rMain.Bottom = SHIPHEIGHT
'copys the info to the buffer
Backbuffer.BltFast X, Y, Ship, rMain, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
End Sub

Public Sub ShootBullet(ByVal XDirect As Long, ByVal YDirect As Long, ByVal Col As Long)
'attempts to create a new bullet
Dim I As Long
For I = 0 To NUMBULLETS - 1
If Bullets(I).Active = False Then
Bullets(I).Active = True
Bullets(I).Xp = XDirect
Bullets(I).Yp = YDirect
Bullets(I).X = ShipX + GUNX
Bullets(I).Y = ShipY + GUNY
'Bullets(I).Color = Color '(I)
Randomize Timer
Bullets(I).Phase = Int(Rnd * (BULLETPHASES + 1))
GunSound(I).Play DSBPLAY_DEFAULT
Timer1.Enabled = True
Exit For
End If
Next I
End Sub

Public Sub UpdateBullet(ByVal Num As Long, ByVal Phase As Long)
'updates bullets
If Paused = False Then
Bullets(Num).X = Bullets(Num).X + Bullets(Num).Xp
Bullets(Num).Y = Bullets(Num).Y + Bullets(Num).Yp
Bullets(Num).Phase = Bullets(Num).Phase + 1
If Bullets(Num).X > 800 Then Bullets(Num).Active = False
If Bullets(Num).Y > 600 Then Bullets(Num).Active = False
If Bullets(Num).Phase > BULLETPHASES Then Bullets(Num).Phase = 0
End If
If Bullets(Num).Active = True Then
Dim rMain As RECT
rMain.Top = 0
rMain.Left = Phase * BULLETWIDTH
rMain.Right = rMain.Left + BULLETWIDTH
rMain.Bottom = BULLETHEIGHT
Backbuffer.BltFast Bullets(Num).X, Bullets(Num).Y, BulletsS, rMain, DDBLTFAST_SRCCOLORKEY
End If
End Sub

Public Sub CheckKeys()
'Static Coun
'Coun = Coun + 1
'If Coun <= 5 Then Exit Sub Else Coun = 6
'checks keys
If Paused = False Then
If GetAsyncKeyState(vbKeyDown) < 0 Then ShipY = ShipY + 5: If ShipY + SHIPHEIGHT > 600 Then ShipY = 600 - SHIPHEIGHT
If GetAsyncKeyState(vbKeyUp) < 0 Then ShipY = ShipY - 5: If ShipY < 0 Then ShipY = 0
If GetAsyncKeyState(vbKeyLeft) < 0 Then ShipX = ShipX - 5: If ShipX < 0 Then ShipX = 0
If GetAsyncKeyState(vbKeyRight) < 0 Then ShipX = ShipX + 5: If ShipX + SHIPWIDTH > 800 Then ShipX = 800 - SHIPWIDTH
If GetAsyncKeyState(vbKeyEscape) < 0 Then RUNNING = False
If GetAsyncKeyState(vbKeyLButton) < 0 Then RUNNING = False
If GetAsyncKeyState(vbKeySpace) < 0 Then If CanShoot = True Then ShootBullet 7, 0, RGB(0, 255, 0): CanShoot = False
If GetAsyncKeyState(vbKeyP) < 0 And CanUnPause = False < 0 Then Paused = True: Timer2.Enabled = True
If GetAsyncKeyState(vbKeyAdd) < 0 Then Volume = Volume + 50: If Volume > 1000 Then Volume = 1000: Perform.SetMasterVolume Volume
If GetAsyncKeyState(vbKeySubtract) < 0 Then Volume = Volume - 50: If Volume < 0 Then Volume = 0: Perform.SetMasterVolume Volume
Else
If GetAsyncKeyState(vbKeyP) < 0 And CanUnPause = True < 0 Then Paused = False: Timer3.Enabled = True
If GetAsyncKeyState(vbKeyEscape) < 0 Then RUNNING = False
End If
End Sub

Public Sub CreateAsteroid(ByVal XDirect As Long, ByVal YDirect As Long)
'creates a new random asteroid
'just holds the for loop
Dim I As Long
For I = 0 To NUMASTEROIDS - 1
'sets the randomize seed
Randomize Timer
'checks if asteroid is active
If Asteroids(I).Active = False Then
'makes asteroid active
Asteroids(I).Active = True
'sets the x direction movement
Asteroids(I).Xp = XDirect
'sets the y direction movement
Asteroids(I).Yp = YDirect
'sets the current x position
Asteroids(I).X = 800
'sets a randome y position
Asteroids(I).Y = Int(Rnd * (600 - ASTEROIDHEIGHT))
'sets a random points value
Asteroids(I).Worth = Int(Rnd * NUMASTEROIDS) + NUMASTEROIDS / 2
'sets a random damage amount
Asteroids(I).Damage = Int(Rnd * 10) + 1
'sets the asteroids phase back to 0
Asteroids(I).Phase = 0
'determines how fast asteroid should spin
Asteroids(I).Spin = Int(((Abs(XDirect) + Abs(YDirect)) / 9) * MAXSPIN)
'makes sure that a number is divided by 0
If Asteroids(I).Spin = MAXSPIN Then Asteroids(I).Spin = MAXSPIN - 1
Exit For
End If
Next I
End Sub

Public Sub UpdateAsteroids(ByVal Num As Long)
'updates asteroids
'doesn't want to update if it's paused
If Paused = False Then
'holds boolean that says if asteroid off screen
Dim AstMiss As Boolean
AstMiss = False
'updates the asteroids x location
Asteroids(Num).X = Asteroids(Num).X + Asteroids(Num).Xp
'updates the asteroids y location
Asteroids(Num).Y = Asteroids(Num).Y + Asteroids(Num).Yp
'updates the asteroids phase
If Frame Mod (MAXSPIN - Asteroids(Num).Spin) = 0 Then Asteroids(Num).Phase = Asteroids(Num).Phase + 1
'checks to see if asteroid has gone of right
If Asteroids(Num).X > 800 Then Asteroids(Num).Active = False: AstMiss = True
'checks to see if asteroid has gone of bottom
If Asteroids(Num).Y > 600 Then Asteroids(Num).Active = False: AstMiss = True
'checks to see if asteroid has gone of left
If Asteroids(Num).X + ASTEROIDWIDTH < 0 Then Asteroids(Num).Active = False: AstMiss = True
'checks to see if asteroid has gone of top
If Asteroids(Num).Y + ASTEROIDHEIGHT < 0 Then Asteroids(Num).Active = False: AstMiss = True
'if asteroid isn't active, we must take points
If AstMiss = True Then PlayerScore = PlayerScore - Int(Asteroids(Num).Worth / 5)
'makes sure that the phase is okay
If Asteroids(Num).Phase > ASTEROIDPHASES Then Asteroids(Num).Phase = 0
End If
'checks if asteroid is active
If Asteroids(Num).Active = True Then
'holds the place where bitmap will be drawn from
Dim rMain As RECT
'sets the top to 0
rMain.Top = 0
'finds the X postion of the picture
rMain.Left = Asteroids(Num).Phase * ASTEROIDWIDTH
'finds the right side of phase picture
rMain.Right = rMain.Left + ASTEROIDWIDTH
'sets the bottom of the picture
rMain.Bottom = ASTEROIDHEIGHT
'blits image to backbuffer
Backbuffer.BltFast Asteroids(Num).X, Asteroids(Num).Y, AsteroidsS, rMain, DDBLTFAST_SRCCOLORKEY
Else
'creates a new asteroid
Call CreateAsteroid(Int(Rnd * 5) - 6, Int(Rnd * 6) - 3)
End If
End Sub

Public Sub TestCollision(ByVal Num As Long)
'checks for collision between ship and asteroid
If Asteroids(Num).X < ShipX + SHIPWIDTH And Asteroids(Num).X + ASTEROIDWIDTH > ShipX And Asteroids(Num).Y < ShipY + SHIPHEIGHT And Asteroids(Num).Y + ASTEROIDHEIGHT > ShipY Then Asteroids(Num).Active = False: PlayShipHit: Explosions(Num).X = Asteroids(Num).X: Explosions(Num).Y = Asteroids(Num).Y: Explosions(Num).Xp = Asteroids(Num).Xp: Explosions(Num).Yp = Asteroids(Num).Yp: Explosions(Num).Active = True: ExplosionSound(Num).Play DSBPLAY_DEFAULT: PlayerScore = PlayerScore + Int(Asteroids(Num).Worth / 2): PlayerHealth = PlayerHealth - Asteroids(Num).Damage
End Sub

Public Sub TestBulletCollision(ByVal Num As Long)
'for the for loop
Dim I As Long
For I = 0 To NUMASTEROIDS - 1
'checks to see if asteroids have been shot
If Asteroids(I).X <= Bullets(Num).X + BULLETWIDTH And Asteroids(I).X + ASTEROIDWIDTH >= Bullets(Num).X And Asteroids(I).Y <= Bullets(Num).Y + BULLETHEIGHT And Asteroids(I).Y + ASTEROIDHEIGHT >= Bullets(Num).Y Then Asteroids(I).Active = False: Bullets(Num).Active = False: Explosions(I).Active = True: Explosions(I).X = Asteroids(I).X: Explosions(I).Y = Asteroids(I).Y: Explosions(I).Xp = Asteroids(I).Xp: Explosions(I).Yp = Asteroids(I).Yp: ExplosionSound(I).Play DSBPLAY_DEFAULT: PlayerScore = PlayerScore + Asteroids(I).Worth
Next I
End Sub

Public Sub Explode(ByVal Num As Long)
If Explosions(Num).Active = True Then
Dim rMain As RECT
If Paused = False Then
Explosions(Num).Phase = Explosions(Num).Phase + 1
If Explosions(Num).Phase = EXPLOSIONPHASES Then Explosions(Num).Phase = -1: Explosions(Num).Active = False: Exit Sub
Explosions(Num).X = Explosions(Num).X + Explosions(Num).Xp
Explosions(Num).Y = Explosions(Num).Y + Explosions(Num).Yp
End If
rMain.Top = 0
rMain.Left = EXPLOSIONWIDTH * Explosions(Num).Phase
rMain.Bottom = EXPLOSIONHEIGHT
rMain.Right = rMain.Left + EXPLOSIONWIDTH
Backbuffer.BltFast Explosions(Num).X, Explosions(Num).Y, ExplosionsS, rMain, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
End If
End Sub

Public Sub PlayShipHit()
Dim I As Long
Dim C As DSCURSORS
For I = 0 To NUMSHIPHIT
ShipHit(I).GetCurrentPosition C
If C.lPlay = 0 Then ShipHit(I).Play DSBPLAY_DEFAULT: Exit For
Next I
End Sub

Private Sub HealthRegen_Timer()
PlayerHealth = PlayerHealth + 1
If PlayerHealth > 100 Then PlayerHealth = 100
End Sub

Private Sub Timer1_Timer()
CanShoot = True
'Timer1.Enabled = False
End Sub

Public Sub DrawLetter(ByVal Phase As Long, ByVal X As Long, ByVal Y As Long)
Dim rMain As RECT
rMain.Top = 0
rMain.Left = FONTWIDTH * Phase
rMain.Right = rMain.Left + FONTWIDTH
rMain.Bottom = FONTHEIGHT
Backbuffer.BltFast X, Y, sFont, rMain, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
End Sub

Public Sub DrawVolume(ByVal Phase As Long, ByVal X As Long, ByVal Y As Long)
Dim rMain As RECT
rMain.Top = 0
rMain.Left = VOLUMEWIDTH * Phase
rMain.Right = rMain.Left + VOLUMEWIDTH
rMain.Bottom = VOLUMEHEIGHT
Dim N As Long
N = Volume
Dim B As Long
For B = 0 To N - 1
Backbuffer.BltFast X, Y, Vol, rMain, DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT
X = X + VOLUMEWIDTH + 2
Next B
End Sub

Public Sub DrawString(ByVal Str As String, ByVal X As Long, ByVal Y As Long, ByVal Space As Long)
Dim N As Long
Str = UCase(Str)
For I = 1 To Len(Str)
Select Case Mid(Str, I, 1)
Case "A": N = 0
Case "B": N = 1
Case "C": N = 2
Case "D": N = 3
Case "E": N = 4
Case "F": N = 5
Case "G": N = 6
Case "H": N = 7
Case "I": N = 8
Case "J": N = 9
Case "K": N = 10
Case "L": N = 11
Case "M": N = 12
Case "N": N = 13
Case "O": N = 14
Case "P": N = 15
Case "Q": N = 16
Case "R": N = 17
Case "S": N = 18
Case "T": N = 19
Case "U": N = 20
Case "V": N = 21
Case "W": N = 22
Case "X": N = 23
Case "Y": N = 24
Case "Z": N = 25
Case "1": N = 26
Case "2": N = 27
Case "3": N = 28
Case "4": N = 29
Case "5": N = 30
Case "6": N = 31
Case "7": N = 32
Case "8": N = 33
Case "9": N = 34
Case "0": N = 35
Case "!": N = 36
Case ":": N = 37
End Select
Call DrawLetter(N, X, Y)
X = X + Fonts(N).Worth + Space
Next I
End Sub
Public Sub DrawScore(ByVal Str As String, ByVal X As Long, ByVal Y As Long, ByVal Space As Long)
Call DrawLetter(18, X, Y)
X = X + Fonts(18).Worth + Space
Call DrawLetter(2, X, Y)
X = X + Fonts(2).Worth + Space
Call DrawLetter(14, X, Y)
X = X + Fonts(14).Worth + Space
Call DrawLetter(17, X, Y)
X = X + Fonts(17).Worth + Space
Call DrawLetter(4, X, Y)
X = X + Fonts(4).Worth + Space
Call DrawLetter(37, X, Y)
X = X + Fonts(37).Worth + Space
Dim N As Long
For I = 1 To Len(Str)
Select Case Mid(Str, I, 1)
Case "1": N = 26
Case "2": N = 27
Case "3": N = 28
Case "4": N = 29
Case "5": N = 30
Case "6": N = 31
Case "7": N = 32
Case "8": N = 33
Case "9": N = 34
Case "0": N = 35
End Select
Call DrawLetter(N, X, Y)
X = X + Fonts(N).Worth + Space
Next I
End Sub

Public Sub SetFontWidths()
'A
Fonts(0).Worth = 15
'B
Fonts(1).Worth = 10
'C
Fonts(2).Worth = 13
'D
Fonts(3).Worth = 14
'E
Fonts(4).Worth = 10
'F
Fonts(5).Worth = 9
'G
Fonts(6).Worth = 13
'H
Fonts(7).Worth = 11
'I
Fonts(8).Worth = 5
'J
Fonts(9).Worth = 5
'K
Fonts(10).Worth = 11
'L
Fonts(11).Worth = 10
'M
Fonts(12).Worth = 16
'N
Fonts(13).Worth = 13
'O
Fonts(14).Worth = 16
'P
Fonts(15).Worth = 10
'Q
Fonts(16).Worth = 15
'R
Fonts(17).Worth = 12
'S
Fonts(18).Worth = 10
'T
Fonts(19).Worth = 14
'U
Fonts(20).Worth = 17
'V
Fonts(21).Worth = 17
'W
Fonts(22).Worth = 22
'X
Fonts(23).Worth = 17
'Y
Fonts(24).Worth = 16
'Z
Fonts(25).Worth = 11
'1
Fonts(26).Worth = 5
'2
Fonts(27).Worth = 11
'3
Fonts(28).Worth = 11
'4
Fonts(29).Worth = 11
'5
Fonts(30).Worth = 10
'6
Fonts(31).Worth = 11
'7
Fonts(32).Worth = 11
'8
Fonts(33).Worth = 11
'9
Fonts(34).Worth = 11
'0
Fonts(35).Worth = 15
'!
Fonts(36).Worth = 5
':
Fonts(37).Worth = 5
End Sub

Public Sub DrawHealth()
Dim rMain As RECT
rMain.Top = 0
rMain.Bottom = 15
rMain.Left = 700
Backbuffer.SetForeColor RGB(255, 0, 0)
'Backbuffer.BltColorFill rMain, RGB(0, 0, 255)
Backbuffer.DrawBox 700, 0, 800, 15
rMain.Right = 700 + PlayerHealth
Backbuffer.SetForeColor vbBlue 'RGB(0, 255, 0)
Backbuffer.BltColorFill rMain, RGB(0, 255, 0)
End Sub

Public Sub ReadKeys()
'this sub checks the keys
Dim State As DIKEYBOARDSTATE
Call diDevice.GetDeviceStateKeyboard(State)
If Paused = False Then
If State.Key(DIK_DOWN) Then ShipY = ShipY + 5: If ShipY + SHIPHEIGHT > 600 Then ShipY = 600 - SHIPHEIGHT
If State.Key(DIK_UP) Then ShipY = ShipY - 5: If ShipY < 0 Then ShipY = 0
If State.Key(DIK_LEFT) Then ShipX = ShipX - 5: If ShipX < 0 Then ShipX = 0
If State.Key(DIK_RIGHT) Then ShipX = ShipX + 5: If ShipX + SHIPWIDTH > 800 Then ShipX = 800 - SHIPWIDTH
If State.Key(DIK_ESCAPE) Then RUNNING = False
If State.Key(DIK_PAUSE) And CanUnPause = False Then Paused = True: Timer2.Enabled = True
If State.Key(DIK_ADD) And VolumeCan = True Then
Volume = Volume + VOLUMECHANGE
If Volume > 100 Then Volume = 100
Perform.SetMasterVolume ((MAXVOLUME - MINVOLUME) * (Volume / 100) + MINVOLUME)
VolumeCan = False
VolChange.Enabled = True
VolShow.Enabled = True
End If
If State.Key(DIK_SUBTRACT) And VolumeCan = True Then
Volume = Volume - VOLUMECHANGE
VolumeCan = False
If Volume < 0 Then Volume = 0
Perform.SetMasterVolume ((MAXVOLUME - MINVOLUME) * (Volume / 100) + MINVOLUME)
VolumeCan = False
VolChange.Enabled = True
VolShow.Enabled = True
End If
If State.Key(DIK_SPACE) Then If CanShoot = True Then ShootBullet 7, 0, RGB(0, 255, 0): CanShoot = False
Else
If State.Key(DIK_PAUSE) And CanUnPause = True Then Paused = False: Timer3.Enabled = True
If State.Key(DIK_ESCAPE) Then RUNNING = False
End If
End Sub

Private Sub Timer2_Timer()
'makes it so that you can unpause (paused)
CanUnPause = True
'stops the timer from firing again until needed
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
'makes it so that you can't unpause (not paused)
CanUnPause = False
'stops the timer from firing again until needed
Timer3.Enabled = False
End Sub

Private Sub VolChange_Timer()
'makes it so you can change the volume
VolumeCan = True
'stops the timer from firing again until needed
VolChange.Enabled = False
End Sub

Private Sub VolShow_Timer()
'makes volume bar disappear
VolShow.Enabled = False
End Sub
