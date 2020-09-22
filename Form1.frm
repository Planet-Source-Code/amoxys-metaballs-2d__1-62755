VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gravitating Metadiscs"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox pctOutput 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   4335
      Left            =   120
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Metaball"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Metaball"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuDithered 
         Caption         =   "Dithered"
      End
      Begin VB.Menu mnuColoured 
         Caption         =   "Coloured"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Type typMetaball
   x As Single  'pos(x)
   y As Single  'pos(y)
   r As Single  'radius
   A As Single  'radius^2
   sx As Single 'speed(x)
   sy As Single 'speed(y)
End Type

Const LIMIT As Single = 2
Const RASTER As Long = 5 'only for coloured metaballs
Const GRAVITATION As Single = 0.01


Dim Ball() As typMetaball, c As Long
Dim OutputWidth As Single, OutputHeight As Single
Dim dPattern(0 To 3, 0 To 3) As Long
Dim bDither As Boolean
Dim bColour As Boolean
Dim bCancel As Boolean


Private Sub Form_Load()
Dim i As Long, x As Long, y As Long


'At first we generate a dithering pattern (0 to 15)
For x = 0 To 3
   For y = 0 To 3
      dPattern(x, y) = 8 * ((x + y) Mod 2) + _
                       4 * (y Mod 2) + _
                       2 * ((x \ 2 + y \ 2) Mod 2) + _
                           ((y \ 2) Mod 2)
   Next y
Next x


'Then we do some other stuff
Show

bDither = mnuDithered.Checked
bColour = mnuColoured.Checked

Randomize
c = -1
For i = 1 To 4
   addBall
Next i

'At last we start the loop
MainLoop
'and quit the application
End
End Sub

Public Sub addBall()
c = c + 1

ReDim Preserve Ball(0 To c)

Ball(c).x = Rnd * OutputWidth
Ball(c).y = Rnd * OutputHeight
Ball(c).r = 35 + Rnd * 60
Ball(c).A = Ball(c).r * Ball(c).r
Ball(c).sx = (1 + Rnd * 3) * IIf(Rnd < 0.5, -1, 1)
Ball(c).sy = (1 + Rnd * 3) * IIf(Rnd < 0.5, -1, 1)
End Sub

Public Sub removeBall()
If c >= 0 Then
   c = c - 1
   If c >= 0 Then
      ReDim Preserve Ball(c)
   End If
End If
End Sub

Public Sub MainLoop()
Do Until bCancel
   Do Until bCancel Or Not bDither Or bColour
     PaintDithered
     Animate
     DoEvents
   Loop
   Do Until bCancel Or bDither Or bColour
     Paint
     Animate
     DoEvents
   Loop
   Do Until bCancel Or bDither Or Not bColour
     PaintColoured
     Animate
     DoEvents
   Loop
Loop 'keeps the loop in the loop
End Sub

Private Sub Paint()
'This sub displays the metaballs undithered

Dim i As Long, f As Single, x As Single, y As Single
Dim dx As Single, dy As Single

pctOutput.Cls

For x = 0 To OutputWidth
   For y = 0 To OutputHeight
      
      f = 0
      For i = 0 To c
         dx = x - Ball(i).x
         dy = y - Ball(i).y
         f = f + Ball(i).A / (dx * dx + dy * dy + 1)  'add 1 to prevent division by zero
      Next i
      
      
      If f >= LIMIT Then SetPixel pctOutput.hdc, x, y, vbWhite
   Next y
Next x
End Sub

Private Sub PaintDithered()
'This sub displays the metaballs dithered

Dim i As Long, f As Single, x As Single, y As Single
Dim dx As Single, dy As Single

pctOutput.Cls
For x = 0 To OutputWidth
   For y = 0 To OutputHeight
      
      f = 0
      For i = 0 To c
         dx = x - Ball(i).x
         dy = y - Ball(i).y
         f = f + Ball(i).A / (dx * dx + dy * dy + 1) 'add 1 to prevent division by zero
      Next i
      f = -Log(1 / (f + 1)) 'add 1 to prevent division by zero
      
     
      If f <= 0 Then
         'draw a black pixel.. um, not
      ElseIf f > LIMIT Then
         'above limit
         SetPixel pctOutput.hdc, x, y, vbWhite
      ElseIf (f * 16 / LIMIT) > dPattern(x Mod 4, y Mod 4) Then
         SetPixel pctOutput.hdc, x, y, vbWhite
      End If
   Next y
Next x
End Sub

Private Sub PaintColoured()
Dim i As Long, f As Single, x As Single, y As Single
Dim dx As Single, dy As Single

pctOutput.Cls
For x = 0 To OutputWidth Step RASTER
   For y = 0 To OutputHeight Step RASTER
   
      f = 0
      For i = 0 To c
         dx = x - Ball(i).x
         dy = y - Ball(i).y
         f = f + Ball(i).A / (dx * dx + dy * dy + 1) 'add 1 to prevent division by zero
      Next i
      
      If f > LIMIT Then
         f = f - LIMIT
         pctOutput.Line (x, y)-Step(RASTER, RASTER), generateColour(f), BF
      End If
   Next y
Next x
End Sub

Public Function generateColour(ByVal f As Single) As Long
Dim r As Long, g As Long, b As Long

f = -Log(1 / f) * 256
If f >= 765 Then
   generateColour = vbWhite
   Exit Function
ElseIf f <= 0 Then
   Exit Function
End If

If f > 255 Then
   r = 255
   f = f - 255
   If f > 255 Then
      g = 255
      b = f
   Else
      g = f
   End If
Else
   r = f
End If


generateColour = RGB(r, g, b)
End Function

Public Sub Animate()
'This sub moves the metaballs by their old speed
'and calculates their new speed

Dim i As Long, j As Long
Dim dx As Single, dy As Single, dist As Single, accel As Single

For i = 0 To c
   Ball(i).x = Ball(i).x + Ball(i).sx
   Ball(i).y = Ball(i).y + Ball(i).sy
   
   If Ball(i).x < Ball(i).r Then
      Ball(i).sx = Abs(Ball(i).sx / 4)
      Ball(i).x = Ball(i).r
   ElseIf Ball(i).x + Ball(i).r > OutputWidth Then
      Ball(i).sx = -Abs(Ball(i).sx / 4)
      Ball(i).x = OutputWidth - Ball(i).r
   End If
   If Ball(i).y < Ball(i).r Then
      Ball(i).sy = Abs(Ball(i).sy / 4)
      Ball(i).y = Ball(i).r
   ElseIf Ball(i).y + Ball(i).r > OutputHeight Then
      Ball(i).sy = -Abs(Ball(i).sy / 4)
      Ball(i).y = OutputHeight - Ball(i).r
   End If
Next i

For i = 0 To c - 1
   For j = i + 1 To c
      dx = Ball(j).x - Ball(i).x
      dy = Ball(j).y - Ball(i).y
      dist = Sqr((dx) ^ 2 + (dy) ^ 2) + 1 'add 1 to prevent division by zero
      dx = dx / dist
      dy = dy / dist
      accel = GRAVITATION * Ball(j).A / dist
      Ball(i).sx = Ball(i).sx + accel * dx
      Ball(i).sy = Ball(i).sy + accel * dy
      accel = GRAVITATION * Ball(i).A / dist
      Ball(j).sx = Ball(j).sx - accel * dx
      Ball(j).sy = Ball(j).sy - accel * dy
   Next j
Next i
End Sub

Private Sub Form_Resize()
pctOutput.Move 0, 0, ScaleWidth, ScaleHeight
OutputWidth = pctOutput.ScaleWidth
OutputHeight = pctOutput.ScaleHeight
End Sub

Private Sub mnuAdd_Click()
addBall
End Sub

Private Sub mnuRemove_Click()
removeBall
End Sub

Private Sub mnuColoured_Click()
mnuColoured.Checked = Not mnuColoured.Checked
bColour = mnuColoured.Checked
If bColour Then
   mnuDithered.Checked = False
   bDither = False
End If
End Sub

Private Sub mnuDithered_Click()
mnuDithered.Checked = Not mnuDithered.Checked
bDither = mnuDithered.Checked
If bDither Then
   mnuColoured.Checked = False
   bColour = False
End If
End Sub

Private Sub pctOutput_Click()
'Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
bCancel = True
End Sub
