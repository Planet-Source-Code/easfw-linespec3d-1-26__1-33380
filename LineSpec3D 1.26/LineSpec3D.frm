VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   Caption         =   "LineSpec3D 1.26"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'LineSpec3D 1.26 - fluoats@hotmail.com - updated 4-23-2002
'Mostly interface cosmetics

'I am re-writing my antialias sub.  It could be weeks or
'months, depending on how my brain is.  Result will be
'faster, and complete lines.  Right now they don't do
'endpoints
'==============

'Now serving the pasta.

'These arrays hold 'backbuffer' data used to erase/draw
'up to 200000 pixels quickly

Dim savX(0 To 200000) As Long    'Filled
Dim savy(0 To 200000) As Long    'by
Dim savAlpha(0 To 200000) As Single 'antialias()

Dim savcL(0 To 200000) As Long  'Used
Dim savBG(0 To 200000) As Long  'by
Dim svBG(0 To 200000) As Long   'Sub
Dim svX(0 To 200000) As Long    'AnimWireframe()
Dim svY(0 To 200000) As Long    ' ...

Dim csn(0 To 200000) As Boolean 'AnimWireframe / antialias
Dim pixarray As Long, arraylen As Long

'Pixel processing in AnimWireframe()
Dim intensitybyt As Byte
Dim r2 As Integer, g2 As Long, b2 As Long
Dim r As Single, g As Single, b As Single
Dim bytR As Byte, bytG As Byte, bytB As Byte
Dim a As Single, s1 As Single, s2 As Single

'Line Info arrays -
'increase these if you want more points in any wireframe
Dim LnPt1(1 To 1001) As Integer, LnPt2(1 To 1001) As Integer
Dim px(1 To 1001) As Single, pY(1 To 1001) As Single
'pX and pY are used by AnimWireframe() and antialias()
Dim pointcount As Integer

'increase these if you want more lines
Dim opac(1 To 1000) As Byte 'Line Intensity
Dim cR(1 To 1000) As Byte
Dim cG(1 To 1000) As Byte
Dim cB(1 To 1000) As Byte
Dim ds(1 To 1000) As Byte 'drawstyle
Dim linecount As Integer

Dim modelselect As Byte 'Subs Form_Keydown and LineSpec()

'Used by many subs to determine how form layout goes
Dim modeselect As String
Dim shiftdown As Boolean

Dim fw As Long, fh As Long 'formwidth, formheight
Dim fw1 As Long, fw2 As Long
Dim sw As Long, sh As Long
Dim eye As Long, radius As Single, zdep As Single
Dim axi As Single, ayi As Single 'model rotation
Dim ay As Single, ax As Single, az As Single
Dim elap As Long, elap2 As Long

Dim breakloop As Boolean, resized As Boolean
Dim newbackground As Boolean, waving As Boolean

'Point data
Dim x3!(1 To 1001), y3!(1 To 1001), z3!(1 To 1001)
'scaled point data written to by ScaleModel(), used by AnimWireframe()
Dim x5!(1 To 1001), y5!(1 To 1001), z5!(1 To 1001)

''Temporary' model data
Dim pcTemp As Long, lcTemp As Long
Dim x3Temp!(1 To 1001), y3Temp!(1 To 1001), z3Temp!(1 To 1001)
Dim Pt1Temp(1 To 1001) As Integer, Pt2Temp(1 To 1001) As Integer
Dim rTemp(1 To 1000) As Byte
Dim gTemp(1 To 1000) As Byte
Dim bTemp(1 To 1000) As Byte
Dim opTemp(1 To 1000) As Byte
Dim dsTemp(1 To 1000) As Byte

Dim sr(0 To 128) As Single 'Form_Load() fills these 'look-up table'
Dim gs(0 To 255) As Single 'arrays which antialias() accesses

Dim drawselect As Byte

'"Virtual" controls
Dim vleft%(1 To 10), vtop%(1 To 10)
Dim vright%(1 To 10), vbot%(1 To 10)
Dim vmax(1 To 10) As Single
Dim vmin(1 To 10) As Single
Dim vval(1 To 10) As Single

'Mouse Control
Dim yInit As Integer, selectv As Byte
Dim xr As Integer, yr As Integer
Dim xr2 As Integer, yr2 As Integer
Dim ptspac As Single
Dim pressed As Boolean, skipped As Boolean

'Multi-Purpose
Dim BGR As Long, cL As Long
Dim intFN As Integer, sngAP As Single, sngX!, sngY!
Dim lngX&, lngY&, intN1%, intN2%
Dim bytW As Byte, sng1 As Single
Dim shape As Byte, pow As Single
Dim ns1 As Long, ns2 As Long, n2 As Long
Const pi As Single = 3.14159265
Const twopi = 2 * pi

Dim Fin As Boolean 'If true then Exits Do in AnimWireframe,
'then Unloads Me

'Background appearance - See Form_Load
Dim bwi As Boolean
Dim wildbackground As Boolean

'Horizontal Background Fade
Dim ditr As Byte, ditg As Byte, ditb As Byte
Dim h2 As Integer, kh As Integer
Dim rr As Byte, gg As Byte, bb As Byte
Dim dr As Byte, dg As Byte, db As Byte
Dim bg1 As Long, bg2 As Long, bg3 As Long
Dim bg4 As Long, bg5 As Long, bg6 As Long

'color shift in draw mode
Dim incr As Single, incg As Single, incb As Single
Dim outline2 As Long, outline As Long 'button and slider color

'File access
Dim FreeFileN As Integer
Const vc As String = ","
Dim lnfd As String 'LineFeed
Dim m As Long, strMult1 As String, Q As String
Dim bodystart As Long, tailstart As Long
Dim modelfound As Boolean
Dim iA As Boolean, iB As Boolean, iC As Boolean, iD As Boolean, iE As Boolean
Dim redC As Byte, redP As Byte 'P represents Previous
Dim grnC As Byte, grnP As Byte 'C represents Current
Dim bluC As Byte, bluP As Byte
Dim opC As Byte, opaP As Byte
Dim dsC As Byte, dsP As Byte
Dim Lnp1C As Long, Lnp1P As Long
Dim Lnp2C As Long, Lnp2P As Long
Dim qM As String, bytLen As Byte, lngLen As Long
Dim x3n As Single, y3n As Single, z3n As Single 'Temporary 'current point' info
Dim x32 As Single, y32 As Single, z32 As Single 'Temporary 'previous point' info
Dim pc1 As Integer 'pointcount
Dim linecoun1 As Integer 'linecount

'Line information
Dim LnPoint1 As Long, LnPoint2 As Long, lookofprev As Byte
Dim lineRed As Byte, lineGrn As Byte, lineBlu As Byte
Dim lineIntensity As Byte, linedrawstyle As Byte

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Form_GotFocus()
 b = ((BGR2& And &HFF0000) / &H10000) And &HFF
 g = ((BGR2& And &HFF00) / &H100) And &HFF
 r = BGR2& And &HFF
 vval(3) = r
 vval(4) = g
 vval(5) = b
 If modeselect = "Create" Then
  bytW = 10
  Form1.Line (vleft(bytW), vtop(bytW))-(vright(bytW), vbot(bytW)), BGR2, BF
  For lngY = 0 To 220 Step 1
  Form1.Line (fw1, lngY)-(fw1 - 9, lngY), RGB(r, g, b)
  Call colors
  Next lngY
 End If
End Sub

Private Sub Form_Load()
bg1 = 170: bg2 = 190: bg3 = 210 'Bytes - background 'dither' color
bg4 = 0: bg5 = 0: bg6 = 78

Form1.ScaleMode = vbPixels
Form1.FontSize = 10
Randomize

bwi = 0    'Boolean - 0 for light background
wildbackground = False 'Bok bok bok bok b'gah

sr(0) = 0: sr(128) = 0.5
For bytW = 1 To 127
sr(bytW) = (bytW / 128) ^ 2 / 2
 Next bytW

shape = 10 'byte used to apply a curve to 'brightness'
           'values in gs() array - used by antialias()

pow = 1 / (1 + shape / 255)
gs(0) = 0: gs(255) = 1
For bytW = 1 To 254
gs(bytW) = (bytW / 255) ^ pow
 Next bytW

vmax(3) = 255
vmax(4) = 255
vmax(5) = 255
vmax(6) = 15
vmax(7) = 15
vmax(8) = 15
vmax(9) = 250: vmin(9) = 50
vval(9) = 100

'Random staring rotation speed
axi = 0.012 * (Rnd - 0.5): ayi = 0.008 * (Rnd - 0.5)
ax = pi

modeselect = "View"
ptspac = 10

'Create file if none exists
Open "LSpecModels.txt" For Append As #1
Close #1

lnfd = Chr(10)
modelselect = 1
LineSpec
End Sub
Private Sub Form_Activate()
Call AnimWireframe
End Sub
Private Sub LineSpec()
pointcount = 0: linecount = 0

Select Case modelselect
Case 1
 Call PointDEF(-0.5, 0.5, -0.5)
 Call PointDEF(0.5, 0.5, -0.5)
 Call PointDEF(-0.5, -0.5, -0.5)
 Call PointDEF(0.5, -0.5, -0.5)
 Call PointDEF(-0.5, 0.5, 0.5)
 Call PointDEF(0.5, 0.5, 0.5)
 Call PointDEF(-0.5, -0.5, 0.5)
 Call PointDEF(0.5, -0.5, 0.5)
 
                     'red,green
 Call LineDEF(1, 2, , 255, 230)
 Call LineDEF(1, 3, 1) 'the ", 1" means "look of previous line"
 Call LineDEF(1, 5, 1)
 Call LineDEF(4, 2, , 255, 255, 255, 128) '128 for intensity
 Call LineDEF(4, 3, 1)
 Call LineDEF(4, 8, 1)
 Call LineDEF(7, 3, , 255, 255, 120, 50)
 Call LineDEF(7, 5, 1)
 Call LineDEF(7, 8, 1)
 Call LineDEF(6, 2, , 90, 80, 0, 200)
 Call LineDEF(6, 5, 1)
 Call LineDEF(6, 8, 1)
 
 'Extra line for style
 Call PointDEF(0.3, 0.3, 0.3)
 Call PointDEF(-0.3, 0.3, 0.3)
 Call LineDEF(9, 10, , , 255, , 150) 'A green line at 150 intensity
 
Case 2 'Type '2' on the keyboard to see this model at run-time
 'upper lip
 Call PointDEF(-0.91, -0.435, -0.5)
 Call PointDEF(-0.84, -0.404, -0.52)
 Call PointDEF(-0.65, -0.34, -0.6)
 Call PointDEF(-0.43, -0.276, -0.64)
 Call PointDEF(-0.19, -0.25, -0.65)
 Call PointDEF(0, -0.28, -0.65)
 Call PointDEF(0.19, -0.25, -0.65)
 Call PointDEF(0.43, -0.276, -0.64)
 Call PointDEF(0.65, -0.34, -0.6)
 Call PointDEF(0.84, -0.404, -0.52)
 Call PointDEF(0.91, -0.435, -0.5)
 Call LineDEF(1, 2, , 255, 120, 255, , 1) 'Drawstyle!  0 1 or 2
 Call LineDEF(2, 3, 1)
 Call LineDEF(3, 4, 1)
 Call LineDEF(4, 5, 1)
 Call LineDEF(5, 6, 1)
 Call LineDEF(6, 7, 1)
 Call LineDEF(7, 8, 1)
 Call LineDEF(8, 9, 1)
 Call LineDEF(9, 10, 1)
 Call LineDEF(10, 11, 1)
 
 'these are the coordinates for the lower lip
 Call PointDEF(-0.88, -0.46, -0.5) '
 Call PointDEF(-0.655, -0.56, -0.56)
 Call PointDEF(-0.35, -0.695, -0.59)
 Call PointDEF(0, -0.723, -0.59)
 Call PointDEF(0.35, -0.695, -0.59)
 Call PointDEF(0.655, -0.56, -0.56)
 Call PointDEF(0.88, -0.46, -0.5)
 Call LineDEF(1, 12, , 255, 120, 255)
 Call LineDEF(12, 13, 1)
 Call LineDEF(13, 14, 1)
 Call LineDEF(14, 15, 1)
 Call LineDEF(15, 16, 1)
 Call LineDEF(16, 17, 1)
 Call LineDEF(17, 18, 1)
 Call LineDEF(18, 11, 1)
 
 'Nose
 Call PointDEF(0, 0.78, -0.57)
 Call PointDEF(0, 0.01, -0.95)
 Call PointDEF(-0.21, -0.09, -0.7)
 Call PointDEF(0.21, -0.09, -0.7)
 Call LineDEF(19, 20, , 255, 176, 0, 192)
 Call LineDEF(20, 21, , 255, 176, 0, , 1)
 Call LineDEF(20, 22, 1)
 
 'This generates the eye circle
 lngX = pointcount: lngY = lngX + 1
 For sngAP = 0 To twopi Step twopi / 23
 lngX = lngX + 1
 Call PointDEF(0.14 * Sin(sngAP) - 0.56, 0.14 * Cos(sngAP) + 1.1, -0.5)
 If lngX < lngY + 23 Then Call LineDEF(lngX, lngX + 1, , 165, 128, 255)
 Next sngAP

Case 3 'Type '1' on the keyboard to see this model at run-time
 For sngX = -4.5 To 4.5 Step 1
 For sngY = -4.5 To 4.5 Step 1
 Call PointDEF(sngX, sngY)
 Next sngY
 Next sngX
 bytW = 80: lngY = 0
 For lngY = 0 To 80 Step 10
 For lngX = 1 To 9 Step 1
  cL = lngX + lngY
  Call LineDEF(cL, cL + 1, , bytW, bytW, 255 - bytW)
  Call LineDEF(cL, cL + 10, , , 255 - bytW)
  Call LineDEF(cL, cL + 11, , 255 - bytW, 255 - bytW, 192)
  bytW = bytW + 1
 Next lngX
 Next lngY
 
Case 4 'Type '4' on the keyboard
For ns1 = 1 To 200
Call PointDEF(Rnd - 0.5, Rnd - 0.5, Rnd - 0.5)
Next ns1
For ns1 = 2 To 200 Step 2
Call LineDEF(ns1, ns1 - 1, , 255 * Rnd, 255 * Rnd, 255 * Rnd, 255 * Rnd)
Next ns1

Case 5
End Select

ScaleModel
End Sub
Private Sub AnimWireframe()
Static backbuff() As Long
Static draw(0 To 300000) As Boolean, bcleansurf(0 To 1600, 0 To 1200) As Boolean
Static bclneapix(0 To 300000) As Boolean, backpxln() As Long
Static x1 As Single, y1 As Single
Static x2 As Single, y2 As Single
Static bytN1 As Byte, Framerate As Single
Static avgfr, framewait As Byte
Static x4 As Long, y4 As Long
Static X As Single, Y As Single, z As Single
Static vpd As Single  'vanishing-point distortion, or near-far distortion
Static csay As Single, snay As Single 'csay = cos(ay)
Static csaz As Single, snaz As Single
Static csax As Single, snax As Single
Static cswap As Integer
Static ax1 As Single, ay1 As Single, realpix As Long, pixeln As Long

'This part is used to initialize a bunch of things
'Some variables can be changed safely - they have comments
If newbackground Then
 Select Case fw
 Case Is <= 8: ReDim backbuff(0 To 8, 0 To 6): fw = 8: fh = 6
 Case Is <= 640: ReDim backbuff(0 To 640, 0 To 480): fw = 640: fh = 480: ReDim backpxln(0 To 640, 0 To 480)
 Case Is <= 800: ReDim backbuff(0 To 800, 0 To 600): fw = 800: fh = 600: ReDim backpxln(0 To 800, 0 To 600)
 Case Is <= 1024: ReDim backbuff(0 To 1024, 0 To 768): fw = 1024: fh = 768: ReDim backpxln(0 To 1024, 0 To 768)
 Case Is <= 1280: ReDim backbuff(0 To 1280, 0 To 1024): fw = 1280: fh = 1024: ReDim backpxln(0 To 1280, 0 To 1024)
 Case Is <= 1600: ReDim backbuff(0 To 1600, 0 To 1200): fw = 1600: fh = 1200: ReDim backpxln(0 To 1600, 0 To 1200)
 End Select

 'screen colorfade effect
 Select Case bwi
 Case True
 Form1.BackColor = vbBlack
 Case Else
 Form1.BackColor = vbBlack
 End Select
 If modeselect = "View" Then
 If wildbackground Then
  X = 9 'Rnd * 5 + 8
  Y = 100 '+ Rnd * 6000
  For sng1 = 0 To fw1
  For n2 = 0 To fw2
  cL = (12254 * Sin(n2 * sng1 / Y) + 12024 * Sin(sng1 / (X))) / 98
  If cL > 16777216 Then
  cL = 16777216
  ElseIf cL < 0 Then
  cL = 0
  End If
  r2 = cL& And &HFF
  If bwi Then If r2 < 120 Then r2 = 120
  r2 = 255 - r2 * 0.8
  b2 = 255 - r2
  g2 = 255 - b2
  'the If here produces the triangular background
  If sng1 And n2 Then
  b2 = 255 - b2
  End If
  If bwi Then
  cL = RGB(255 - r2, 255 - g2, 255 - b2)
  Else
  cL = RGB(r2, g2, b2)
  End If
  backbuff(sng1, n2) = cL&
  SetPixelV Form1.hdc, sng1, n2, cL
  Next n2
  Next sng1
 Else
  If bwi Then
  ditr = bg4: ditg = bg5: ditb = bg6
  Else
  ditr = bg1: ditg = bg2: ditb = bg3
  End If
  For h2 = 255 To 0 Step -1  'make this -2 for an interesting effect
  dr = ditr / 255 * h2
  dg = ditg / 255 * h2
  db = ditb / 255 * h2
  Select Case bwi
  Case 0: rr = 255 - dr: gg = 255 - dg: bb = 255 - db: kh = 741 - h2 * 3
  Case Else: rr = dr: gg = dg: bb = db: kh = 601 - h2 * 3
  End Select
  BGR = RGB(rr, gg, bb)
  For n2 = 0 To fw1 Step 1
  For cL = kh - 2 To kh Step 1
  If cL > -1 And cL < fh Then backbuff(n2, cL) = BGR
  Next cL
  If kh > 740 Then
  For cL = kh + 1 To fw2 Step 1
  Select Case bwi
  Case True
  If cL < fh Then backbuff(n2, cL) = vbBlack
  Case Else
  If cL < fh Then backbuff(n2, cL) = vbWhite
  End Select
  Next cL
  End If
  Next n2
  Next h2
 End If
 End If
 For ns1 = 1 To realpix 'clear coordinate array
 savX(ns1) = 0: savy(ns1) = 0
 svX(ns1) = 0: svY(ns1) = 0
 Next ns1
 elap = GetTickCount
 Form_Paint
 newbackground = False 'we are done with this section until a form resize
End If

'Here is the loop that animates
 'I do not actually use animation 'frames'. _
  I erase one pixel of old line then draw one pixel _
  of new line within a loop

Do While Not Fin
 DoEvents: If Fin Or breakloop Then Exit Do
 If modeselect = "View" Then
 
 'This part allows for constant rotation speed
 'regardless of cpu power or model complexity.
 'The multiplier (.05) scales the rotation speed
 
 elap2 = elap
 elap = GetTickCount
 BGR = elap - elap2: sngAP = BGR * 0.05
 
 'Framerate computer
 Select Case framewait
 Case 7
 avgfr = (avgfr + BGR) / 7
 Form1.Line (4, 4)-(175, 19), vbWhite, BF
 If avgfr <> 0 Then Framerate = Round(1000 / avgfr, 1)
 Form1.CurrentX = 6: Form1.CurrentY = 4: Form1.ForeColor = vbBlue
 Form1.Print "Framerate(FPS) :  "; Framerate
 framewait = 1: avgfr = BGR
 Case Else
 Form1.Line (vleft(1), vtop(1))-(vright(1), vbot(1)), outline2, B
 For bytW = 2 To 10 Step 1
 Select Case bytW
 Case 2 To 5, 9
 Form1.Line (vleft(bytW), vtop(bytW))-(vright(bytW), vbot(bytW)), outline2, B
 Case 6, 7, 8
 Form1.Line (vleft(bytW), vtop(bytW))-(vright(bytW), vbot(bytW)), outline, B
 End Select
 Next bytW
 If bwi Then
 Form1.ForeColor = RGB(175, 149, 9)
 Else
 Form1.ForeColor = RGB(40, 19, 90)
 End If
 Form1.CurrentX = vleft(2) + 3: Form1.CurrentY = vtop(2) + 2
 Form1.Print " Click to enter Draw Mode"
 framewait = framewait + 1: avgfr = avgfr + BGR
 End Select
 
 ay = ay + ayi * sngAP 'incrementing rotation
 ax = ax + axi * sngAP
 az = az + 0
 If ay > twopi Then
 ay = ay - twopi
 ElseIf ay < 0 Then ay = ay + twopi: End If
 If ax > twopi Then
 ax = ax - twopi
 ElseIf ax < 0 Then ax = ax + twopi: End If
 If az > twopi Then
 az = az - twopi
 ElseIf az < 0 Then az = az + twopi: End If
 snay = Sin(ay): csay = Cos(ay) 'precalc some things
 snax = Sin(ax): csax = Cos(ax)
 snaz = Sin(az): csaz = Cos(az)
 
 For kh% = 1 To pointcount Step 1 'performing Euler transform to rotate points
 z = z5(kh%) * csay - x5(kh%) * snay
 X = x5(kh%) * csay + z5(kh%) * snay
 Y = y5(kh%) * csax - z * snax 'Note: ScaleModel writes to the x5 y5 z5 arrays.
 z = z * csax + y5(kh%) * snax 'The 'non-destructed' values are in
 X = X * csaz - Y * snaz      'the x3 y3 z3 arrays.
 Y = Y * csaz + X * snaz
 z = radius * z
 vpd = radius * eye / (eye - z) 'vpd = vanishing-point distortion
 X = X * vpd + sw 'horizontal center
 Y = Y * vpd + sh
 px(kh%) = X: pY(kh%) = Y
 Next kh%
 
 If waving Then
 'Here is where we want to apply changes to the height of points
 ax1 = ax1 + 0.2: ay1 = ay1 + 0.2
 If ax1 > twopi Then
 ax1 = ax1 - twopi
 ElseIf ax1 < 0 Then ax1 = ax1 + twopi: End If
 If ay1 > twopi Then
 ay1 = ay1 - twopi
 ElseIf ay1 < 0 Then ay1 = ay1 + twopi: End If
 lngX = 1
  For sngX = -4.5 To 4.5 Step 1
  For sngY = -4.5 To 4.5 Step 1
  z5(lngX) = 0.1 * Sin(sngY + sngX + ax1 + ay1)
  lngX = lngX + 1
  Next sngY
  Next sngX
 End If

 'Calling antialias( LineNumber ) to generate x, y,
 'alpha values, and the number of pixels to be drawn
 pixarray& = 0
 For kh% = 1 To linecount%
 Call antialias(kh%)
 Next kh%
 
 realpix = 0
 
 'Puting all coordinate data within bounds of backbuff() array
 'Making sure that pixel at certain location is drawn once/frame
 For ns1& = 1 To pixarray Step 1
 If Not csn(ns1) Then
 x4& = savX&(ns1&): y4& = savy&(ns1&)
 If x4& > -1& And y4& > -1& And x4& < fw& And y4& < fh& Then
 draw(ns1&) = 1
 Select Case bcleansurf(x4&, y4&)
 Case False
 realpix& = realpix& + 1
 bcleansurf(x4&, y4&) = True
 bclneapix(ns1&) = True
 savBG(realpix&) = backbuff&(x4&, y4&)
 Case Else
 bclneapix(ns1&) = False
 End Select
 Else
 draw(ns1&) = 0: End If
 End If
 Next ns1&: pixeln = 0

 cswap = 0 'cswap is used to determine which line we are drawing
 'This heavy loop computes color over background based upon alpha that antialias() generates
 For ns1& = 1 To pixarray& Step 1
 x4& = savX&(ns1&): y4& = savy&(ns1&)
 Select Case csn(ns1&) 'at start of each new line, Sub antialias()
                       'sets csn(first pixel # of that line) = True
                       'For purpose of establishing new line color,
 Case True             'intensity and drawstyle
 cswap = cswap + 1
 If cswap > linecount Then cswap = 0
 intensitybyt = opac(cswap): drawselect = ds(cswap)
 r = cR(cswap): g = cG(cswap): b = cB(cswap)
 'savcL(ns1) = backbuff&(x4&, y4&)
 'The cR,cG,cB,opac,ds arrays are filled in Sub LineSpec
 csn(ns1&) = 0 'reset
 Case False
 Select Case draw(ns1&)
 Case True
 a! = savAlpha!(ns1&)
 BGR& = backbuff&(x4&, y4&)
 b2& = ((BGR& And &HFF0000) / &H10000) And &HFF
 g2& = ((BGR& And &HFF00) / &H100) And &HFF
 r2% = BGR& And &HFF
 s2! = (intensitybyt / 255) * a!
 Select Case drawselect
 Case 0: cL& = RGB(r2% - s2! * (r2% - r), g2& - s2! * (g2& - g), b2& - s2! * (b2& - b))
 Case 1: cL& = RGB(r2% - s2! * (r2% - a! * r), g2& - s2! * (g2& - a! * g), b2& - s2! * (b2& - a! * b))
 Case 2: cL& = RGB(r2% + s2! * Abs(r2% - r) ^ 0.85, g2& + s2! * Abs(g2& - g) ^ 0.85, b2& + s2! * Abs(b2& - b) ^ 0.85)
 End Select
  Select Case bclneapix(ns1&)
  Case True
  pixeln& = pixeln& + 1&
  savX&(pixeln&) = x4&: savy&(pixeln&) = y4&
  savcL&(pixeln&) = cL&
  backpxln&(x4&, y4&) = pixeln&
  Case Else
  savcL&(backpxln&(x4&, y4&)) = cL&
  End Select
  backbuff&(x4&, y4&) = cL&
  bcleansurf(x4&, y4&) = False
 End Select
 End Select
 Next ns1&

 If arraylen < realpix Then 'If draw pixel# > erase pixel#
  For ns1 = 1 To arraylen Step 1
  x4 = svX(ns1): y4 = svY(ns1): cL = svBG(ns1)
  If cL = backbuff(x4, y4) Then SetPixelV Form1.hdc, x4, y4, cL
  SetPixelV Form1.hdc, savX(ns1), savy(ns1), savcL(ns1)
  Next ns1
  For cL = ns1 To realpix Step 1
  SetPixelV Form1.hdc, savX(cL), savy(cL), savcL(cL)
  Next cL
 Else
  For ns1 = 1 To realpix Step 1
  x4 = svX(ns1): y4 = svY(ns1): cL = svBG(ns1)
  If cL = backbuff(x4, y4) Then SetPixelV Form1.hdc, x4, y4, cL
  SetPixelV Form1.hdc, savX(ns1), savy(ns1), savcL(ns1)
  Next ns1
  For cL = ns1 To arraylen Step 1
  x4 = svX(cL): y4 = svY(cL): BGR = svBG(cL)
  If BGR = backbuff(x4, y4) Then SetPixelV Form1.hdc, x4, y4, BGR
  Next cL
 End If
 
 'clean backbuffer
 For ns1 = 1 To realpix Step 1
 cL = savX(ns1): n2 = savy(ns1)
 backbuff(cL, n2) = savBG(ns1)
 svX(ns1) = cL: svY(ns1) = n2
 svBG(ns1) = savBG(ns1)
 Next ns1
 
 'Store this cycle's # of pixels drawn, used in next cycle to erase
 arraylen = realpix

 End If
Loop
 
If breakloop Then breakloop = 0: AnimWireframe
Unload Me

End Sub

' E S S E N T I A L - numbers in here not meant to be messed with
Private Sub antialias(LineN As Integer)
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim spx As Double, epx As Double
Dim spy As Double, epy As Double
Dim ax As Double, bx As Double, cx As Double, dx As Double
Dim ay As Double, by As Double, cy As Double, dy As Double
Dim ex As Double, ey As Double
Dim mp5 As Single, pp5 As Single
Dim rex As Integer, rey As Integer

Dim trz As Double, tri As Double
Dim lwris As Double, lwrun As Double
Dim zsl As Single
Dim slope As Double, lsope As Double
Dim midx As Double, midy As Double
Dim sl2 As Single
Dim distanc1 As Double, distanc2 As Double
Dim diagonal As Boolean
Dim a As Single
Dim st As Integer, lp1 As Integer, lp2 As Integer
Dim one As Single

pixarray = pixarray + 1  'Here, we are faking +1 to pixel count before
           'any pixel data on a new line is computed.
           
           'This value will be used in a loop within
           'AnimWireframe to determine at which pixel
           'number does a new line start.

csn(pixarray) = 1 'So, at this n in the array, AnimWireframe will
           'know when to change r,g,b and intensitybyt
           'values based upon specific line 'properties'
           'that are assigned to cR,cG,cB,opac arrays.
           
           'To see where this happens, Find " Case csn("

'Okay, time to get computing.
lp1 = LnPt1(LineN): lp2 = LnPt2(LineN)
x1 = px(lp1): y1 = pY(lp1)
x2 = px(lp2): y2 = pY(lp2)

If x1 < x2 Then
 epy = y2: epx = x2
 spy = y1: spx = x1
Else
spy = y2: spx = x2
epy = y1: epx = x1: End If

If epx = spx Or epy = spy Then
 diagonal = 0
 If epy > spy Then
  st = 1
 Else: st = -1: End If

Else: slope = (epy - spy) / (epx - spx): lsope = -1 / slope
diagonal = 1: End If

midx = 0.5 * lsope: midy = 0.5 * slope
sl2 = slope * slope: one = 1 / Sqr(1 + sl2)
lwris = 1 - (one - Abs(slope) + sl2 * one)
lwrun = lwris * Abs(lsope)
distanc1 = 0.5 * one
distanc2 = slope * distanc1
ax = spx - distanc1 - distanc2
ay = spy + distanc1 - distanc2
bx = epx + distanc1 - distanc2
by = epy + distanc1 + distanc2
cx = epx + distanc1 + distanc2
cy = epy - distanc1 + distanc2
dx = spx - distanc1 + distanc2
dy = spy - distanc1 - distanc2

one = 255 * (1 - 0.5 * lwris * lwrun)

If diagonal Then
If slope > 0 Then
If slope <= 1 Then
ey# = slope# * (Round(ax#) + 1.5 - ax#) + ay#
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1
tri# = pp5! - lwris#: trz# = pp5! - slope#: zsl! = mp5! - midy
For ex# = Round(ax#) + 1.5 To Round(bx#) - 1.5 Step 1
pixarray = pixarray + 1
 If ey# > tri# Then
  savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(one!)
 Else
  If ey# > trz# Then
  savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 + lsope# * sr(Int((pp5! - ey#) * 128))))
  Else: savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (ey# - zsl!)): End If
 End If
If ey# > trz# Then
 pp5! = pp5! + 1: tri# = tri# + 1: trz# = trz# + 1
 zsl! = zsl! + 1: rey% = rey% + 1: End If
ey = ey# + slope: Next ex#

ey# = cy# - slope# * (cx# - Round(cx#) + 1.5)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! + midy
tri# = mp5! + lwris: trz# = mp5! + slope
For ex# = Round(cx#) - 1.5 To Round(dx#) + 1.5 Step -1
If ey# > tri# Then
pixarray = pixarray + 1
 If ey# > trz# Then
 savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (zsl! - ey#))
 Else: savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 + lsope# * sr(Int((ey# - mp5!) * 128))))
 End If
End If
If ey# < trz# Then
 mp5! = mp5! - 1: tri# = tri# - 1: trz# = trz# - 1
 zsl! = zsl! - 1: rey% = rey% - 1: End If
ey# = ey# - slope: Next ex#

ex# = cx# + lsope * (cy# - Round(cy#) + 0.5)
For ey# = Round(cy#) - 0.5 To Round(dy#) + 1.5 Step -1
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (slope# * sr(Int((ex# - rex% + 0.5) * 128))))
ex# = ex# + lsope#: Next ey#

ex# = ax# - lsope# * (Round(ay#) + 0.5 - ay#)
For ey# = Round(ay) + 0.5 To Round(by) - 0.5 Step 1
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (slope# * sr(Int((rex% + 0.5 - ex#) * 128))))
ex# = ex# - lsope#: Next ey#

Else
ex# = dx# - lsope# * (Round(dy#) + 1.5 - dy#): rex% = Round(ex#): mp5! = rex% - 0.5
pp5! = mp5! + 1: tri# = pp5! - lwrun: trz# = pp5! + lsope: zsl! = mp5! + midx
For ey# = Round(dy#) + 1.5 To Round(cy#) - 0.5 Step 1
pixarray = pixarray + 1
 If ex# > tri# Then
  savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(one!)
 Else
  If ex# > trz# Then
   savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (1 - slope * sr(Int((pp5! - ex#) * 128))))
  Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (ex# - zsl!)): End If
 End If
If ex# > trz# Then
 tri# = tri# + 1: trz# = trz# + 1: pp5! = pp5! + 1
 zsl! = zsl! + 1: rex% = rex% + 1: End If
ex# = ex# - lsope#: Next ey#

ex# = bx# + lsope# * (by# - Round(by#) + 0.5): rex% = Round(ex#): mp5! = rex% - 0.5
pp5! = mp5! + 1: tri# = mp5! + lwrun: trz# = mp5! - lsope: zsl! = pp5! - midx#
For ey# = Round(by#) - 0.5 To Round(ay#) + 1.5 Step -1
If ex# > tri# Then
 pixarray = pixarray + 1
 If ex# < trz# Then
  savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (1 - slope# * sr(Int((ex# - mp5!) * 128))))
 Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (zsl! - ex#)): End If
End If
If ex# < trz# Then
 tri# = tri# - 1: trz# = trz# - 1: mp5! = mp5! - 1
 rex% = rex% - 1: zsl! = zsl! - 1: End If
ex# = ex# + lsope#: Next ey#

ey# = dy# + slope# * (Round(dx#) + 0.5 - dx#)
For ex# = Round(dx#) + 0.5 To Round(cx#) - 1.5 Step 1
 pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (-lsope# * sr(Int((rey% + 0.5 - ey#) * 128))))
ey# = ey# + slope#: Next ex#

ey# = ay# + slope# * (Round(ax#) + 1.5 - ax#)
For ex# = Round(ax#) + 1.5 To Round(bx#) - 0.5 Step 1
 pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (-lsope# * sr(Int((ey# - rey% + 0.5) * 128))))
ey# = ey# + slope: Next ex#
End If

Else
If slope > -1 Then
ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = pp5! - midy#
tri# = mp5! + lwris: trz# = mp5! - slope#
For ex# = Round(dx#) + 1.5 To Round(cx#) - 1.5 Step 1
pixarray = pixarray + 1
 If ey# < tri# Then
  savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(one!)
 Else
  If ey# < trz# Then
   savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 - lsope# * sr(Int((ey# - mp5!) * 128))))
  Else: savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (zsl! - ey#)): End If
 End If
If ey# < trz# Then
  mp5! = mp5! - 1: zsl! = zsl! - 1: tri# = tri# - 1
  rey% = rey% - 1: trz# = trz# - 1: End If
ey# = ey# + slope#: Next ex#

ey# = by# - slope# * (bx# - Round(bx#) + 1.5)
rey% = Round(ey#): mp5! = rey% - 0.5: pp5! = mp5! + 1: zsl! = mp5! + midy#
tri# = pp5! - lwris: trz# = pp5! + slope#
For ex# = Round(bx#) - 1.5 To Round(ax#) + 1.5 Step -1
If ey# < tri# Then
 pixarray = pixarray + 1
 If ey# < trz# Then
  savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (ey# - zsl!))
 Else
  savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (1 - lsope# * sr(Int((pp5! - ey#) * 128))))
 End If
End If
If ey# > trz# Then
  rey% = rey% + 1: pp5! = pp5! + 1: zsl! = zsl! + 1
  tri# = tri# + 1: trz# = trz# + 1: End If
ey# = ey# - slope#: Next ex#

ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5)
For ey# = Round(ay#) - 1.5 To Round(by#) + 0.5 Step -1
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (-slope * sr(Int((ex# - rex% + 0.5) * 128))))
ex# = ex# + 2 * midx#: Next ey#

ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#)
For ey# = Round(cy#) + 1.5 To Round(dy#) - 0.5
pixarray = pixarray + 1
 rex% = Round(ex#)
 savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (-slope# * sr(Int((rex% + 0.5 - ex#) * 128))))
ex# = ex# - 2 * midx#: Next ey#

Else
ex# = ax# + lsope# * (ay# - Round(ay#) + 1.5): rex% = Round(ex#): mp5! = rex% - 0.5
zsl! = mp5! - midx#: pp5! = mp5! + 1: tri# = pp5! - lwrun#: trz# = pp5! - lsope#
For ey# = Round(ay#) - 1.5 To Round(by#) + 1.5 Step -1
pixarray = pixarray + 1
 If ex# > tri# Then
  savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(one!)
 Else
  If ex# > trz# Then
   savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (1 + slope# * sr(Int((pp5! - ex#) * 128))))
  Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey#): savAlpha(pixarray) = gs(255 * (ex# - zsl!)): End If
 End If
If ex# > trz# Then
  tri# = tri# + 1: trz# = trz# + 1: pp5! = pp5! + 1
  rex% = rex% + 1: zsl! = zsl! + 1: End If
ex# = ex# + lsope#: Next ey#

ex# = cx# - lsope# * (Round(cy#) + 1.5 - cy#): rex% = Round(ex#): mp5! = rex% - 0.5
tri# = mp5! + lwrun#: trz# = mp5! + lsope#: zsl! = mp5! + 1 + midx#
For ey# = Round(cy#) + 1.5 To Round(dy#) - 1.5 Step 1
 If ex# > tri# Then
 pixarray = pixarray + 1
  If ex# < trz# Then
   savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (1 + slope * sr(Int((ex# - mp5!) * 128))))
  Else: savX(pixarray) = rex%: savy(pixarray) = Int(ey# + 1): savAlpha(pixarray) = gs(255 * (zsl! - ex#)): End If
 End If
If ex# < trz# Then
 tri# = tri# - 1: trz# = trz# - 1: zsl! = zsl! - 1
 rex% = rex% - 1: mp5! = mp5! - 1: End If
ex# = ex# - lsope#: Next ey#

ey# = by# - slope# * (bx# - Round(bx#) + 2.5)
For ex# = Round(bx#) - 2.5 To Round(ax#) + 0.5 Step -1
pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex# + 1): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (lsope# * sr(Int((ey# - rey% + 0.5) * 128))))
ey# = ey# - slope#: Next ex#

ey# = dy# + slope# * (Round(dx#) + 1.5 - dx#)
For ex# = Round(dx#) + 1.5 To Round(cx#) - 0.5 Step 1
pixarray = pixarray + 1
 rey% = Round(ey#)
 savX(pixarray) = Int(ex#): savy(pixarray) = rey%: savAlpha(pixarray) = gs(255 * (lsope# * sr(Int((rey% + 0.5 - ey#) * 128))))
ey# = ey# + slope#: Next ex#

End If
End If

Else
If epy = spy Then
 rey% = Round(ay): st% = rey% - 1: a! = ay - rey% + 0.5
For ex# = Round(ax) + 1.5 To Round(bx) - 0.5
 pixarray = pixarray + 1
 zsl! = gs(255! * a!): rex% = Int(ex#)
 savX(pixarray) = rex%: savy(pixarray) = rey%: savAlpha(pixarray) = zsl!
 pixarray = pixarray + 1
 savX(pixarray) = rex%: savy(pixarray) = st%: savAlpha(pixarray) = 1 - zsl!
Next ex#
ElseIf epx = spx Then
 rex% = Round(ax): st% = rex% - 1: a! = ax - rex% + 0.5
 If epy > spy Then
 trz = 1
 Else: trz = -1: End If
For ey# = Round(ay) - 0.5 To Round(by) + 1.5 Step trz
 pixarray = pixarray + 1
 zsl! = gs(255! * a!): rey% = Int(ey#)
 savX(pixarray) = rex%: savy(pixarray) = rey%: savAlpha(pixarray) = zsl!
 pixarray = pixarray + 1
 savX(pixarray) = st%: savy(pixarray) = rey%: savAlpha(pixarray) = 1 - zsl!
Next ey#: End If
End If
 
End Sub
Private Sub Form_DblClick()
pressed = 1
axi = axi * 0.4
ayi = ayi * 0.4
dblsngclickcommon
End Sub
Private Sub dblsngclickcommon()
 Select Case selectv
 Case 2
  Form1.CurrentX = vleft(2)
  Form1.CurrentY = vtop(2)
  If modeselect = "View" Then
  modeselect = "Create"
  Me.AutoRedraw = True
  pointcount = 0: linecount = 0
  zdep = 0
  b = ((BGR2& And &HFF0000) / &H10000) And &HFF
  g = ((BGR2& And &HFF00) / &H100) And &HFF
  r = BGR2& And &HFF

  Else
  ScaleModel
  modeselect = "View"
  Me.AutoRedraw = False
  ax = 0: ay = 0
  End If
  Form_Resize
 End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case 16
shiftdown = True

Case 48 To 57 'number keys
modelselect = KeyCode - 48
If modeselect = "View" Then
If shiftdown Then
WriteToFile
Else
LineSpec
Form1.ForeColor = vbWhite
Form1.CurrentX = 40: Form1.CurrentY = 20
'Print EncodeToString(modelselect)
LoadFromFile
End If
Else
If shiftdown Then
WriteToFile
Else
LoadFromFile
modeselect = "View"
Form_Resize
End If
End If
 
Case 65 'a
zdep = zdep - 10
If zdep < -90 Then zdep = -120

Case 69 'e
If modeselect = "View" Then
If wildbackground Then
wildbackground = False
Else
wildbackground = True
End If
breakloop = True: newbackground = True
End If

Case 90 'z
If modeselect = "View" Then
modelselect = 1: LineSpec
Else
zdep = zdep + 10
If zdep > 90 Then zdep = 120
End If

Case 87 'w
If waving Then
waving = False
Else
waving = True
End If

Case 88 'x
If modeselect = "View" Then modelselect = 2: LineSpec

Case 67 'c
If modeselect = "View" Then modelselect = 3: LineSpec

Case 86 'v
If modeselect = "View" Then modelselect = 4: LineSpec


Case 105, 73 'i
If bwi Then
bwi = 0
Else
bwi = 1
End If
newbackground = True: breakloop = True

Case 108, 76 ' (L)oad
'LoadFromFile
'ax = 0: ay = 0

Case 115, 83 ' (S)ave
'WriteToFile

Case 116, 84 ' (T)emporary
If shiftdown Then 'Storing
 pcTemp = pointcount: lcTemp = linecount
 For ns1 = 1 To pcTemp Step 1
 x3Temp(ns1) = x3(ns1): y3Temp(ns1) = y3(ns1): z3Temp(ns1) = z3(ns1)
 Next ns1
 For ns1 = 1 To lcTemp Step 1
 Pt1Temp(ns1) = LnPt1(ns1): Pt2Temp(ns1) = LnPt2(ns1)
 rTemp(ns1) = cR(ns1): gTemp(ns1) = cG(ns1): bTemp(ns1) = cB(ns1)
 opTemp(ns1) = opac(ns1): dsTemp(ns1) = ds(ns1)
 Next ns1
Else 'Retrieving
 pointcount = pcTemp: linecount = lcTemp
 For ns1 = 1 To pcTemp Step 1
 x3(ns1) = x3Temp(ns1): y3(ns1) = y3Temp(ns1): z3(ns1) = z3Temp(ns1)
 Next ns1
 For ns1 = 1 To lcTemp Step 1
 LnPt1(ns1) = Pt1Temp(ns1): LnPt2(ns1) = Pt2Temp(ns1)
 cR(ns1) = rTemp(ns1): cG(ns1) = gTemp(ns1): cB(ns1) = bTemp(ns1)
 opac(ns1) = opTemp(ns1): ds(ns1) = dsTemp(ns1)
 Next ns1
 ScaleModel
End If

Case 27: 'Esc
Fin = 1

End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 selectv = 0
 pressed = 1
 For bytW = 1 To 10
 Select Case X
 Case vleft(bytW) To vright(bytW)
 Select Case Y
 Case vtop(bytW) To vbot(bytW)
 selectv = bytW 'We've landed inside a control's dimensions
 yInit = 0
  End Select
   End Select
    Next bytW
 Select Case selectv
 Case 0
 Case 1
 Case 2
  dblsngclickcommon
 Case 10
 frmRainbow.Show
 End Select
 
 Select Case selectv
 Case 0
  xr = X
  yr = Y
  axi = axi * 0.4
  ayi = ayi * 0.4
 End Select

 Call Form_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pressed Then
If selectv > 0 Then
Select Case Y
Case Is <> yInit
yInit = Y: bytW = selectv
 Select Case Y
 Case vtop(bytW) To vbot(bytW)
 vval(bytW) = vmin(bytW) + (vmax(bytW) - vmin(bytW)) * (vbot(bytW) - Y) / (vbot(bytW) - vtop(bytW))
 Case Is < vtop(bytW)
 vval(bytW) = vmax(bytW)
 Case Is > vbot(bytW)
 vval(bytW) = vmin(bytW)
 End Select
 Select Case bytW
 Case 1
 shape = vval(bytW)
 pow = 1 / (1 + shape / 255)
 gs(0) = 0: gs(255) = 1
 For bytW = 1 To 254
 gs(bytW) = (bytW / 255) ^ pow
 Next bytW
 Case 3 To 8
 If modeselect = "Create" Then
 r = vval(3)
 g = vval(4)
 b = vval(5)
 BGR2 = RGB(r, g, b)
 incr = vval(6)
 incg = vval(7)
 incb = vval(8)
 Form1.Line (vleft(10), vtop(10))-(vright(10), vbot(10)), BGR2, BF
 For lngY = 0 To 220 Step 1
 Form1.Line (fw1, lngY)-(fw1 - 9, lngY), RGB(r, g, b)
 Call colors
 Next lngY
 End If
 Case 9
 radius = vval(9) * sw / 200
 End Select
 
 If Not wildbackground Then
 Select Case bytW
 Case 3
  If modeselect = "View" Then
  If bwi Then
  bg4 = vval(3)
  Else: bg1 = vval(3)
  End If:  End If
 Case 4
  If modeselect = "View" Then
  If bwi Then
  bg5 = vval(4)
  Else: bg2 = vval(4)
  End If: End If
 Case 5
  If modeselect = "View" Then
  If bwi Then
  bg6 = vval(5)
  Else: bg3 = vval(5)
  End If: End If
 End Select
 End If
 
End Select
Else
 If modeselect = "View" Then
 xr = xr - X
 If xr > 0 Then
  If xr > xr2 Then
   xr2 = xr
  End If
  ayi = ayi * 0.8
  If xr > 6 Then
   ayi = ayi + xr2 / 1000
  Else
   ayi = ayi + xr / 500: xr2 = xr2 - xr
  End If
 
 ElseIf xr < 0 Then
  If xr < xr2 Then
   xr2 = xr
  End If
  ayi = ayi * 0.8
  If xr < -6 Then
   ayi = ayi + xr2 / 1000
  Else
   ayi = ayi + xr / 500: xr2 = xr2 - xr
  End If
 End If
 
 yr = yr - Y
 If yr > 0 Then
  If yr > yr2 Then
   yr2 = yr
  End If
  axi = axi * 0.8
  If yr > 6 Then
   axi = axi + yr2 / 1000
  Else
   axi = axi + yr / 500: yr2 = yr2 - yr
  End If
 
 ElseIf yr < 0 Then
  If yr < yr2 Then
   yr2 = yr
  End If
  axi = axi * 0.8
  If yr < -6 Then
   axi = axi + yr2 / 1000
  Else
   axi = axi + yr / 500: yr2 = yr2 - yr
  End If
 End If
 
 xr = X
 yr = Y
 Else 'mode = "Create"
 If pointcount < 1001 Then
  If pointcount < 1 Then
   sngX = X: sngY = Y ': cL = pointcount
   Call PointDEF(sngX - sw, sngY - sh, zdep)
   'Form1.PSet (X, Y), vbBlack
  Else
   If Not skipped Then
   pow = X - sngX: sng1 = Y - sngY
   sngAP = Sqr(pow ^ 2 + sng1 ^ 2)
   If sngAP >= ptspac Then
    sngAP = ptspac / sngAP
    sngX = sngX + pow * sngAP
    sngY = sngY + sng1 * sngAP
    Call PointDEF(sngX - sw, sngY - sh, zdep)
     lngX = pointcount + skipped
     bytR = r: bytG = g: bytB = b
     Call LineDEF(lngX - 1, lngX, , bytR, bytG, bytB)
     Form1.Line (sngX, sngY)-(x3(lngX - 1) + sw, y3(lngX - 1) + sh), RGB(r, g, b)
     'Form1.PSet (sngX, sngY), vbBlack
   End If
   Else 'skipped
    sngX = X: sngY = Y ': cL = pointcount
    Call PointDEF(sngX - sw, sngY - sh, zdep)
    'Form1.PSet (X, Y), vbBlack
   End If
  End If
  End If
  skipped = 0
  Call colors
  
 End If
End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pressed = 0
Select Case selectv
Case 0, 2
If modeselect = "Create" Then
skipped = 1
'If cL = linecount Then pointcount = pointcount - 1
End If
Case 3 To 8
If modeselect = "Create" Then
r = vval(3)
g = vval(4)
b = vval(5)
incr = vval(6)
incg = vval(7)
incb = vval(8)
ElseIf selectv = 3 Or selectv = 4 Or selectv = 5 Then
If Not wildbackground Then breakloop = True: newbackground = True
End If
End Select
End Sub
Private Sub Form_Paint()
Static initclick As Boolean
 If Not wildbackground Then Form1.Cls
 If bwi Then
  If modeselect = "Create" Then
  Form1.BackColor = RGB(40, 20, 40)
  End If
 Else
  If modeselect = "Create" Then
  Form1.BackColor = RGB(245, 245, 205)
  Else
  If Not wildbackground Then
  Form1.BackColor = vbWhite
  End If
  End If
 End If
 
 If Not wildbackground Then
 'background colorfade
  For h2 = 255 To 0 Step -1
  dr = ditr / 255 * h2
  dg = ditg / 255 * h2
  db = ditb / 255 * h2
  Select Case bwi
  Case 0: rr = 255 - dr: gg = 255 - dg: bb = 255 - db: kh = 741 - h2 * 3
  Case Else: rr = dr: gg = dg: bb = db: kh = 601 - h2 * 3
  End Select
  BGR = RGB(rr, gg, bb)
  Form1.Line (0, kh)-(fw, kh), BGR, B
  Form1.Line (0, kh - 1)-(fw, kh - 1), BGR, B
  Form1.Line (0, kh - 2)-(fw, kh - 2), BGR, B
  Next h2
  End If
 
 If Not initclick Then
  initclick = 1
 '
 End If
 
 If bwi Then
  Form1.ForeColor = RGB(200, 175, 0)
  outline = RGB(70, 86, 90)
  outline2 = RGB(110, 255, 0)
 Else
  outline2 = RGB(131, 10, 216)
  Form1.ForeColor = RGB(20, 59, 90)
  outline = RGB(150, 100, 196)
 End If

 If modeselect = "Create" Then
 Form1.Cls
 Form1.CurrentX = vleft(2) + 4: Form1.CurrentY = vtop(2) + 2
 Form1.Print " Click to enter View Mode"
 Form1.CurrentX = 5: Form1.CurrentY = 0
 Form1.Print "Use mouse to draw.  Sliders at right change color"
 Form1.Print "  Draw Further=[a] / Draw Closer=[z]"
 Form1.Print "  To save temporarily, Press Shift-T"
 Form1.Print "  To view models on file, press a number key"
 Else
 If Not wildbackground Then
 Form1.ForeColor = vbWhite
 Else
 If bwi Then
 Form1.ForeColor = vbWhite
 Else
 Form1.ForeColor = RGB(95, 122, 0)
 End If
 End If
 If wildbackground Then
 Form1.CurrentY = 3
 Else
 Form1.CurrentY = -11
 End If
 Form1.Print Tab(28), "               Number keys load models from file." & " " & "[T]" & " Loads model from Temporary."
 Form1.Print Tab(28), "               Shift-{Number Key} saves to file"
 Form1.Print "Other Keys"
 If Not wildbackground Then
 Form1.Print " w e i"
 Form1.Print " z x c v"
 Else
 Form1.Print "         w e i"
 Form1.Print "        z x c v"
 End If
 End If

 'outlines of buttons and sliders
 For bytW = 1 To 10 Step 1
  Select Case bytW
  Case 2 To 8
'
  Form1.Line (vleft(bytW), vtop(bytW))-(vright(bytW), vbot(bytW)), outline2, B
  
  Case 10
  Form1.Line (vleft(bytW), vtop(bytW))-(vright(bytW), vbot(bytW)), BGR2, BF
 
  End Select
 Next bytW
   
End Sub
Private Sub Form_Resize()
fw = Form1.ScaleWidth: fw1 = fw
fh = Form1.ScaleHeight: fw2 = fh
sw = fw / 2
sh = fh / 2

vmax(1) = 255

vleft(2) = 3: vright(2) = vleft(2) + 145
vbot(2) = fh - 3: vtop(2) = vbot(2) - 18

'rgb sliders
vleft(3) = fw - 67: vright(3) = vleft(3) + 10: vtop(3) = fh - 73: vbot(3) = vtop(3) + 70
vleft(4) = vleft(3) + 22: vright(4) = vleft(4) + 10: vtop(4) = vtop(3): vbot(4) = vbot(3)
vleft(5) = vleft(4) + 22: vright(5) = vleft(5) + 10: vtop(5) = vtop(4): vbot(5) = vbot(4)
'rgb oscillation sliders
vleft(6) = vright(3) + 2: vright(6) = vleft(6) + 6: vtop(6) = vtop(3): vbot(6) = vbot(3)
vleft(7) = vright(4) + 2: vright(7) = vleft(7) + 6: vtop(7) = vtop(3): vbot(7) = vbot(3)
vleft(8) = vright(5) + 2: vright(8) = vleft(8) + 6: vtop(8) = vtop(3): vbot(8) = vbot(3)

If modeselect = "Create" Then
vleft(1) = 0: vright(1) = 0: vbot(1) = 0: vtop(1) = 0
vleft(9) = 0: vright(9) = 0: vbot(9) = 0: vtop(9) = 0
vright(10) = vleft(3) - 4: vbot(10) = vbot(3): vleft(10) = vright(10) - 8: vtop(10) = vbot(10) - 8
Else
vleft(1) = 4: vright(1) = vleft(1) + 8
vtop(1) = fh - 126: vbot(1) = vtop(1) + 100
vleft(9) = vright(1) + 4: vright(9) = vleft(9) + 8
vtop(9) = vtop(1): vbot(9) = vbot(1)
vleft(10) = 0: vright(10) = 0: vbot(10) = 0: vtop(10) = 0
End If

If fw > 0 Then eye = 1.6 * sw
radius = vval(9) * sw / 200: breakloop = True: newbackground = True
End Sub
Private Sub PointDEF(Optional ptX As Single, Optional ptY As Single, Optional ptZ As Single)
pointcount = pointcount + 1
x3(pointcount) = Round(ptX, 3): y3(pointcount) = Round(ptY, 3): z3(pointcount) = Round(ptZ, 3)
End Sub
Private Sub LineDEF(Optional LPoint1 As Long, Optional LPoint2 As Long, Optional SameFlavor As Byte, Optional LnRed As Byte, Optional LnGrn As Byte, Optional LnBlu As Byte, Optional LnIntensity As Byte, Optional LnDrawstyle As Byte)
Dim LCn As Integer
 linecount = linecount + 1
 LCn = linecount
 LnPt1(LCn) = LPoint1: LnPt2(LCn) = LPoint2
 If SameFlavor And LCn > 1 Then
 cR(LCn) = cR(LCn - 1)
 cG(LCn) = cG(LCn - 1)
 cB(LCn) = cB(LCn - 1)
 opac(LCn) = opac(LCn - 1)
 ds(LCn) = ds(LCn - 1)
 Else
 cR(LCn) = LnRed
 cG(LCn) = LnGrn
 cB(LCn) = LnBlu
 opac(LCn) = LnIntensity
 End If
 
 If LnIntensity = 0 And SameFlavor = 0 Then opac(LCn) = 255
 If LnDrawstyle < 3 And SameFlavor = 0 Then ds(LCn) = LnDrawstyle
End Sub
Private Sub ScaleModel()
Dim maxlength As Single
'For intFN = 1 To linecount Step 1
'intN1 = LnPt1(intFN): intN2 = LnPt2(intFN)
'sngAP = Sqr((x3(intN1) - x3(intN2)) ^ 2 + _
            (y3(intN1) - y3(intN2)) ^ 2 + _
            (z3(intN1) - z3(intN2)) ^ 2)
For intFN = 1 To pointcount Step 1
sngAP = Sqr(x3(intFN) ^ 2 + y3(intFN) ^ 2 + z3(intFN) ^ 2)
If sngAP > maxlength Then maxlength = sngAP
Next intFN
If maxlength >= 1 Then
For intFN = 1 To pointcount Step 1
x5(intFN) = x3(intFN) / maxlength
y5(intFN) = y3(intFN) / maxlength
z5(intFN) = z3(intFN) / maxlength
Next intFN
Else
For intFN = 1 To pointcount Step 1
x5(intFN) = x3(intFN) * maxlength
y5(intFN) = y3(intFN) * maxlength
z5(intFN) = z3(intFN) * maxlength
Next intFN
End If
End Sub
Private Sub colors()
 's1 = s1 + oscopa
 'If s1 > s1a Then
 's1 = s1a: oscopa = -oscopa
 'ElseIf s1 < 0 Then
 's1 = 0: oscopa = -oscopa
 'End If
 
 r = r + incr
 If r < 0 Then
 incr = -incr: r = 0
 ElseIf r > 255 Then
  incr = -incr: r = 255: End If
 g = g + incg
 If g < 0 Then
 incg = -incg: g = 0
 ElseIf g > 255 Then
  incg = -incg: g = 255: End If
 b = b + incb
 If b < 0 Then
 incb = -incb: b = 0
 ElseIf b > 255 Then
  incb = -incb: b = 255: End If
End Sub
Private Sub LoadFromFile()
Q = ReadBinFile("LSpecModels.txt")
Call BuildModelFromString(Q)
End Sub
Private Sub BuildModelFromString(strm1 As String)
Dim ElemVaL(1 To 7) As Single
Dim ElemValC(1 To 7) As Single
Dim sr1 As String

lngLen = Len(strm1)
modelfound = False: m = 1
Do While modelfound = False
If Mid(strm1, m, 1) = "M" Then
 m = m + 6
 If Mid(strm1, m, 1) = modelselect Then modelfound = True: m = m + 1
Else: m = m + 1: End If
If m >= lngLen Then Exit Do
Loop
m = m + 1
cL = Elem(strm1, m)
If modelfound And Mid(strm1, m + 9, 1) = "M" Then modelfound = False
If modelfound And Mid(strm1, m + 10, 1) = "M" Then modelfound = False
If modelfound And Mid(strm1, m + 10, 1) = "" Then modelfound = False

If modelfound And cL > 0 Then
pointcount = cL
m = m + 8 'skipping the 8 characters of " points"<cr>
x32 = Elem(strm1, m) / 1000
y32 = Elem(strm1, m) / 1000
z32 = Elem(strm1, m) / 1000: m = m + 1
x3(1) = x32: y3(1) = y32: z3(1) = z32

For cL = 2 To pointcount Step 1
If Mid$(strm1, m, 1) = vc Then
 x3n = x32: m = m + 1
 If Mid$(strm1, m, 1) = vc Then
  y3n = y32: m = m + 1
  If Mid$(strm1, m, 1) = lnfd Then
  z3n = z32: m = m + 1
  Else
  z3n = Elem(strm1, m) / 1000: m = m + 1
  End If
 ElseIf Mid$(strm1, m, 1) = lnfd Then
  y3n = y32: z3n = z32: m = m + 1
 Else
  y3n = Elem(strm1, m) / 1000
  If Mid$(strm1, m, 1) = vc Then
  z3n = Elem(strm1, m) / 1000: m = m + 1
  Else
  z3n = z32: m = m + 1
  End If
 End If
Else
 x3n = Elem(strm1, m) / 1000
 If Mid$(strm1, m, 1) = lnfd Then
  y3n = y32: z3n = z32: m = m + 1
 ElseIf Mid$(strm1, m + 1, 1) = vc Then
  y3n = y32: m = m + 1
  If Mid$(strm1, m + 1, 1) = lnfd Then
   z3n = z32: m = m + 1
  Else
   z3n = Elem(strm1, m) / 1000: m = m + 1
  End If
 Else
  y3n = Elem(strm1, m) / 1000
  If Mid$(strm1, m, 1) = lnfd Then
   z3n = z32: m = m + 1
  Else
   z3n = Elem(strm1, m) / 1000: m = m + 1
  End If
 End If
End If
x32 = x3n: y32 = y3n: z32 = z3n
x3(cL) = x3n: y3(cL) = y3n: z3(cL) = z3n
Next cL

Lnp1P = Elem(strm1, m): Lnp2P = Elem(strm1, m): redP = Elem(strm1, m)
grnP = Elem(strm1, m): bluP = Elem(strm1, m): opaP = Elem(strm1, m): dsP = Elem(strm1, m)
LnPt1(1) = Lnp1P: LnPt2(1) = Lnp2P: cR(1) = redP: cG(1) = grnP: cB(1) = bluP
opac(1) = opaP: ds(1) = dsP
ElemVaL(3) = redP: ElemVaL(4) = grnP: ElemVaL(5) = bluP
ElemVaL(6) = opaP: ElemVaL(7) = dsP: m = m + 1
modelfound = False: sr1 = Mid$(strm1, m + 2, 1)
If Mid$(strm1, m + 1, 1) = "M" Then modelfound = True
If sr1 = "o" Or sr1 = "" Then modelfound = True

BGR = 1: ns1 = 1 '
Do While Not modelfound And pointcount > 2 'Do while model isn't finished
'model is finished where 2 linefeeds occur

If Mid$(strm1, m, 1) = vc Then  'we've skipped
Lnp1C = Lnp1P: Lnp2C = Elem(strm1, m)
Else
Lnp1C = Elem(strm1, m)
If Mid$(strm1, m, 1) = vc Then
 If Mid$(strm1, m + 1, 1) = vc Then
 Lnp2C = Lnp2P: m = m + 1
 Else
 Lnp2C = Elem(strm1, m)
 End If
Else
 Lnp2C = Lnp2P
End If
End If

cL = 3
Do While cL < 8
If cL < 7 Then
 If Mid$(strm1, m, 1) = lnfd Then
  For BGR = cL To 7
  ElemValC(BGR) = ElemVaL(BGR)
  Next BGR
  m = m + 1: cL = 7
 ElseIf Mid$(strm1, m, 1) = vc Then
  If Mid$(strm1, m + 1, 1) = vc Then
  ElemValC(cL) = ElemVaL(cL): m = m + 1
  Else
  ElemValC(cL) = Elem(strm1, m)
  End If
 Else
  ElemValC(cL) = ElemVaL(cL): m = m + 1
 End If
Else
 If Mid$(strm1, m, 1) = vc Then
  ElemValC(7) = Elem(strm1, m)
 Else
  ElemValC(7) = ElemVaL(7)
 End If: m = m + 1
End If
 cL = cL + 1
Loop
ns1 = ns1 + 1 'line count

Lnp1P = Lnp1C: Lnp2P = Lnp2C
redP = ElemValC(3): grnP = ElemValC(4): bluP = ElemValC(5)
opaP = ElemValC(6): dsP = ElemValC(7)

LnPt1(ns1) = Lnp1C: LnPt2(ns1) = Lnp2C
cR(ns1) = redP: cG(ns1) = grnP: cB(ns1) = bluP
opac(ns1) = opaP: ds(ns1) = dsP

For cL = 3 To 7
ElemVaL(cL) = ElemValC(cL):
Next cL

modelfound = False
iA = False: iB = False: iC = False: iD = False
If Mid$(strm1, m - 1, 1) = lnfd Then iA = True
If Mid$(strm1, m + 2, 1) = "M" Then iB = True
If Mid$(strm1, m + 1, 1) = lnfd Then iC = True
If Mid$(strm1, m + 1, 1) = "M" Then iD = True

If iA And iB Then modelfound = True
If iB And iC Then modelfound = True
If Mid$(strm1, m, 1) = "M" Or Mid$(strm1, m, 1) = Chr(10) Then modelfound = True
If m > lngLen - 3 Then modelfound = True
Loop
linecount = ns1
ScaleModel
Else
linecount = 0: pointcount = 0
End If '"If modelfound"
End Sub
Private Sub WriteToFile()
modelfound = False

Q = ReadBinFile("LSpecModels.txt")

lngLen = Len(Q)
modelfound = False: m = 1
Do While modelfound = False
If Mid(Q, m, 1) = "M" Then
bodystart = m
If modelselect = 10 Then
 m = m + 8
 If Mid(Q, m, 1) = vc Then modelfound = True
Else
 m = m + 6
 If Mid(Q, m, 1) = modelselect Then modelfound = True: m = m + 1: newbackground = True
End If
Else: m = m + 1: End If
If m >= lngLen Then Exit Do
Loop

If modelfound Then
tailstart = lngLen: cL = m
Do While cL < lngLen
If Mid$(Q, cL, 1) = "M" Then tailstart = cL: modelfound = False
If modelfound = False Then Exit Do
cL = cL + 1
Loop
strMult1 = Left$(Q, bodystart - 1) & _
           EncodeToString(modelselect) & _
           Right$(Q, lngLen - tailstart + 3)
 Open "LSpecModels.txt" For Output As #1
 Print #1, Left$(strMult1, Len(strMult1) - 2)
 Close #1
Else
 Open "LSpecModels.txt" For Output As #1
 Print #1, Q & EncodeToString(modelselect)
 Close #1
End If
End Sub
Private Function EncodeToString(modelnum As Byte) As String
Dim strA2 As String, strB2 As String
Dim strA1 As String 'strA2 is in EncodeToString
Dim strB1 As String 'strB2 "  "  "

'Assembling a string representation of current model

strA1 = "Model" & Str(modelnum) & vc & Str(pointcount) & " points" & lnfd
' Example: "Model 1,5 points"

'To reduce string length, drop decimal points
'PointDEF() already has x3 y3 z3 arrays _
 rounded to nearest thousandth
If linecount > 0 Then
x32 = 1000 * x3(1): y32 = 1000 * y3(1): z32 = 1000 * z3(1)
strA1 = strA1 & x32 & vc & y32 & vc & z32 & lnfd
' Example: "4325,4562,-1000"

For ns1 = 2 To pointcount Step 1
x3n = 1000 * x3(ns1): y3n = 1000 * y3(ns1): z3n = 1000 * z3(ns1)
If x3n <> x32 Then strA2 = strA2 & x3n 'If x is different than last x then write x value
If y3n <> y32 Then 'If y is different than last, append ",yvalue"
 strA2 = strA2 & vc & y3n
 If z3n <> z32 Then  'If different z then write ",z value"
 strA2 = strA2 & vc & z3n
 End If
Else               'y has not changed so we're going to skip writing it
 If z3n <> z32 Then  'If z has changed, append ",,z value"
 strA2 = strA2 & vc & vc & z3n
 End If
End If
strA2 = strA2 & lnfd 'New Line
x32 = x3n: y32 = y3n: z32 = z3n 'x y z values are now 'previous values'
Next ns1
'strA2 might look like
' 2325,4562,-1000
' 2132,4502
' 2125,,-1235
' ,,-1352
' 2145

'first line values get copied into 'previous values' variables
Lnp1P = LnPt1(1): Lnp2P = LnPt2(1): redP = cR(1)
grnP = cG(1): bluP = cB(1): opaP = opac(1): dsP = ds(1)

'Head' of the Line data
strB1 = Lnp1P & vc & Lnp2P & vc & redP & vc & _
 grnP & vc & bluP & vc & opaP & vc & dsP & lnfd
 
For ns1 = 2 To linecount Step 1
Lnp1C = LnPt1(ns1): Lnp2C = LnPt2(ns1)
'As is often the case, if current line shares an endpoint
'with previous line, we can target that
If Lnp1C = Lnp2P Then      'If point1 = prev point2
strB2 = strB2 & Lnp2C & vc ' strB2 = "2,", where , represents point1
Lnp2P = Lnp1C: Lnp1C = Lnp2C: Lnp2C = Lnp2P
ElseIf Lnp1C = Lnp1P Then  'If point1 = prev point1 (cuz sometimes the points can be reversed)
strB2 = strB2 & vc & Lnp2C ' strB2 = ",2", where , represents point1
ElseIf Lnp2C = Lnp1P Then  'If point2 = prev point1
strB2 = strB2 & vc & Lnp1C ' strB2 = ",1"
Lnp2P = Lnp1C: Lnp1C = Lnp2C: Lnp2C = Lnp2P
ElseIf Lnp2C = Lnp2P Then  'If point2 = prev point2
strB2 = strB2 & Lnp1C & vc ' strB2 = "1,"
Else  'shares no endpoints with previous line
strB2 = strB2 & Lnp1C & vc & Lnp2C 'strB2 = "1,2"
End If
Lnp1P = Lnp1C: Lnp2P = Lnp2C

iA = 0: iB = 0: iC = 0: iD = 0: iE = 0
redC = cR(ns1): grnC = cG(ns1): bluC = cB(ns1): opC = opac(ns1): dsC = ds(ns1)
If redC = redP Then iA = True
If grnC = grnP Then iB = True
If bluC = bluP Then iC = True
If opC = opaP Then iD = True
If dsC = dsP Then iE = True

If iA And iB And iC And iD And iE Then
strB2 = strB2 & lnfd
Else
If iA Then
strB2 = strB2 & vc
Else: strB2 = strB2 & vc & redC: End If
If iB Then
strB2 = strB2 & vc
Else: strB2 = strB2 & vc & grnC: End If
If iC Then
strB2 = strB2 & vc
Else: strB2 = strB2 & vc & bluC: End If
If iD Then
strB2 = strB2 & vc
Else: strB2 = strB2 & vc & opC: End If
If Not iE Then strB2 = strB2 & vc & dsC
strB2 = strB2 & lnfd
End If
modelfound = False: cL = 1
Do While Not modelfound
If Mid$(strB2, Len(strB2) - cL, 1) = "," Then
cL = cL + 1
Else
modelfound = True
End If
Loop
strB2 = Left$(strB2, Len(strB2) - cL) & lnfd

If ns1 < linecount Then redP = redC: grnP = grnC: bluP = bluC: opaP = opC: dsP = dsC
Next ns1
End If 'linecount > 0
EncodeToString = strA1 & strA2 & strB1 & strB2
End Function
Private Function Elem(strm2 As String, m2 As Long) As Long
Dim str1 As String, str2 As String, str3 As String
str3 = Mid$(strm2, m2, 1)
If str3 = vc Or str3 = " " Or str3 = lnfd Then m2 = m2 + 1
Do While str2 <> vc And str2 <> " " And str2 <> lnfd
str1 = str1 & str2
str2 = Mid$(strm2, m2, 1)
m2 = m2 + 1
If m2 > lngLen Then Exit Do
Loop
m2 = m2 - 1
Elem = Val(str1)
End Function
Public Function ReadBinFile(ByVal FileName As String)
 Dim strFromFile As String
 Dim lngFileSize As Long
 lngFileSize = FileLen(FileName)
 strFromFile = String(lngFileSize, " ")
 Open FileName For Binary As #1
 Get #1, , strFromFile
 Close #1
 ReadBinFile = strFromFile
End Function
Private Sub Form_Unload(Cancel As Integer)
Fin = True
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 16
shiftdown = False
End Select
End Sub


