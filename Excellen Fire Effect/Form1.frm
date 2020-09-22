VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Burning Text"
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Beauty Of fire  inside text"
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   3075
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Fire Inside Text"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   2460
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Dividing Fire Zone"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   1845
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Beauty of Fire"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   1230
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fire on the Top of Text"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   3000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fire By Mistake"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   3690
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Excellent Fire Effect"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   615
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simple Fire Effect"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6000
      Left            =   0
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Please Wait ......"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1680
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   2595
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fire(2, 255) As Byte
Dim imagedata() As Byte
Dim buff() As Byte, brun(9) As Boolean, w As Integer, h As Integer
Dim fps As Long, kk As Integer

'remove it
Dim hit As Integer

Dim fx1 As Integer, fx2 As Integer
Dim newarr() As Byte
Private Sub Command1_Click()
stopall
brun(0) = True
runme
End Sub

Private Sub Command2_Click()
stopall
brun(1) = True
runme1
End Sub

Private Sub Command3_Click()
Label1.Visible = True
Picture1.BackColor = vbBlack
Picture1.CurrentX = (w - Picture1.TextWidth("PS CODE")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("PS CODE")) / 2
Picture1.Print "PS CODE"


Dim x As Integer, y As Integer, i As Integer, cl As Long
Dim hw As Long
hw = Picture1.hdc
'Picture1.Refresh
getbits
kk = 127

For y = 0 To h - 1
For x = 0 To w - 1
cl = GetPixel(hw, x, y)
If cl <> 0 Then
buff(x, y) = kk
imagedata(0, x, h - y) = fire(2, kk)
imagedata(1, x, h - y) = fire(1, kk)
imagedata(2, x, h - y) = fire(0, kk)
End If
Next
DoEvents
Next
'SavePicture Picture1.Image, App.Path & "\yu.bmp"
Label1.Visible = False
'setbits
stopall
brun(6) = True
runme2
End Sub

Private Sub Command4_Click()
Label1.Visible = True
Picture1.BackColor = vbBlack
Picture1.CurrentX = (w - Picture1.TextWidth("PS-CODE")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("PS-CODE")) / 2
Picture1.Print "PS CODE"
Dim x As Integer, y As Integer, i As Integer, cl As Long
Dim hw As Long
hw = Picture1.hdc
'Picture1.Refresh
getbits
kk = 127
hit = 0
For y = 0 To h - 1
For x = 0 To w - 1
cl = GetPixel(hw, x, y)
If cl <> 0 Then
If hit = 0 Then hit = y
buff(x, y) = kk
imagedata(0, x, h - y) = fire(2, kk)
imagedata(1, x, h - y) = fire(1, kk)
imagedata(2, x, h - y) = fire(0, kk)
End If
Next
DoEvents
Next
'SavePicture Picture1.Image, App.Path & "\yu.bmp"
Label1.Visible = False
'setbits
stopall
brun(7) = True
runme3

End Sub

Private Sub Command5_Click()
stopall
brun(2) = True
runme5

End Sub

Private Sub Command6_Click()
stopall
brun(3) = True
runme6

End Sub


Private Sub Command7_Click()
Label1.Visible = True
Picture1.BackColor = vbBlack
Picture1.FontSize = 60
Picture1.CurrentX = (w - Picture1.TextWidth("PS CODE")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("PS CODE"))
Picture1.Print "PS-CODE"

Picture1.CurrentX = (w - Picture1.TextWidth("ON")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("ON")) - 70
Picture1.Print "ON"

Picture1.CurrentX = (w - Picture1.TextWidth("BEAUTY")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("BEAUTY")) - 140
Picture1.Print "BEAUTY"

Dim x As Integer, y As Integer, i As Integer, cl As Long
Dim hw As Long
hw = Picture1.hdc
'Picture1.Refresh
getbits
kk = 127

For y = 0 To h - 1
For x = 0 To w - 1
cl = GetPixel(hw, x, y)
If cl <> 0 Then
buff(x, y) = kk
imagedata(0, x, h - y) = fire(2, kk)
imagedata(1, x, h - y) = fire(1, kk)
imagedata(2, x, h - y) = fire(0, kk)
End If
Next
DoEvents
Next
'SavePicture Picture1.Image, App.Path & "\yu.bmp"
Label1.Visible = False
'setbits

stopall
brun(4) = True
runme7
End Sub

Private Sub Command8_Click()
Label1.Visible = True
Picture1.BackColor = vbBlack
Picture1.FontSize = 60
Picture1.CurrentX = (w - Picture1.TextWidth("PS CODE")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("PS CODE"))
Picture1.Print "PS-CODE"

Picture1.CurrentX = (w - Picture1.TextWidth("ON")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("ON")) - 70
Picture1.Print "ON"

Picture1.CurrentX = (w - Picture1.TextWidth("BEAUTY")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("BEAUTY")) - 140
Picture1.Print "BEAUTY"

Dim x As Integer, y As Integer, i As Integer, cl As Long
Dim hw As Long
hw = Picture1.hdc
'Picture1.Refresh
getbits
kk = 127

For y = 0 To h - 1
For x = 0 To w - 1
cl = GetPixel(hw, x, y)
If cl <> 0 Then
buff(x, y) = kk
imagedata(0, x, h - y) = fire(2, kk)
imagedata(1, x, h - y) = fire(1, kk)
imagedata(2, x, h - y) = fire(0, kk)
End If
Next
DoEvents
Next
'SavePicture Picture1.Image, App.Path & "\yu.bmp"
Label1.Visible = False
'setbits
stopall
brun(5) = True
runme8

End Sub

Private Sub Command9_Click()
Label1.Visible = True
Picture1.BackColor = vbBlack
Picture1.FontSize = 60
Picture1.CurrentX = (w - Picture1.TextWidth("PS CODE")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("PS CODE"))
Picture1.Print "PS-CODE"

Picture1.CurrentX = (w - Picture1.TextWidth("~ON~")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("ON")) - 80
Picture1.Print "~ON~"

Picture1.CurrentX = (w - Picture1.TextWidth("BEAUTY")) / 2
Picture1.CurrentY = (h - Picture1.TextHeight("BEAUTY")) - 160
Picture1.Print "BEAUTY"

fx1 = 0
fx2 = 0

Dim x As Integer, y As Integer, i As Integer, cl As Long
Dim hw As Long
hw = Picture1.hdc
'Picture1.Refresh
getbits
kk = 127

For y = 0 To h - 1
For x = 0 To w - 1
cl = GetPixel(hw, x, y)
If cl <> 0 Then
If (fx1 = 0) Then fx1 = y
fx2 = y
buff(x, y) = kk
imagedata(0, x, h - y) = fire(2, kk)
imagedata(1, x, h - y) = fire(1, kk)
imagedata(2, x, h - y) = fire(0, kk)
End If
Next
DoEvents
Next
'SavePicture Picture1.Image, App.Path & "\yu.bmp"
Label1.Visible = False
'setbits
stopall
brun(8) = True
runme9
End Sub

Private Sub Form_Load()
MsgBox "Please Compile for speed" & vbCrLf & "Compile EXE file is much faster"
Me.Show
w = Picture1.ScaleWidth
h = Picture1.ScaleHeight

getbits
initialize
stopall
Me.Show
fps = 1
End Sub
Sub setbits()
Dim bm As BITMAP
Dim bmi As BITMAPINFO

bmi.bmiHeader.biSize = 40
bmi.bmiHeader.biPlanes = 1
bmi.bmiHeader.biBitCount = 24
bmi.bmiHeader.biCompression = 0

bmi.bmiHeader.biWidth = w
bmi.bmiHeader.biHeight = h

'SetDIBits Picture1.hdc, Picture1.Image, 0, 256, imagedata(0, 0), bmi, 0
StretchDIBits Picture1.hdc, 0, 0, w, h, 0, 0, w, h, imagedata(0, 0, 0), bmi, 0, vbSrcCopy
Picture1.Refresh

End Sub

Sub getbits()
Dim bm As BITMAP
Dim bmi As BITMAPINFO

bmi.bmiHeader.biSize = 40
bmi.bmiHeader.biPlanes = 1
bmi.bmiHeader.biBitCount = 24
bmi.bmiHeader.biCompression = 0

Dim bmlen As Long
bmlen = Len(bm)



GetObject Picture1.Image, bmlen, bm

ReDim imagedata(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
ReDim buff(bm.bmWidth - 1, bm.bmHeight - 1)

GetDIBits Picture1.hdc, Picture1.Image, 0, bm.bmHeight, imagedata(0, 0, 0), bmi, 0


End Sub
Sub initialize()
Dim i As Integer
For i = 0 To 63
fire(0, i) = i * 4
fire(1, i) = 0
fire(2, i) = 0
Next
For i = 64 To 127

fire(0, i) = 255
fire(1, i) = (i - 64) * 4
fire(2, i) = 0

Next

For i = 128 To 191

fire(0, i) = 255
fire(1, i) = 255
fire(2, i) = (i - 128) * 4


Next
For i = 192 To 255

fire(0, i) = (i - 192) * 4
fire(1, i) = (i - 192) * 4
fire(2, i) = (i - 192) * 4

Next

End Sub
Sub runme()
Dim x As Integer, y As Integer
Dim av As Integer

Do While brun(0)
Randomize
For y = (h - 4) To (h - 1)
For x = 0 To w - 1
buff(x, y) = 255 * Rnd
Next
Next

For y = (h - 3) To 2 Step -1
For x = 1 To (w - 2)
av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1) + buff(x, y + 2) + buff(x, y + 1)
av = av / 5

If av >= 3 Then av = av - 2
buff(x, y) = av

imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
Next
Next
DoEvents
setbits
fps = fps + 1
Loop

End Sub

Sub runme1()
Dim x As Integer, y As Integer
Dim av As Integer

Do While brun(1)
Randomize
For y = (h - 4) To (h - 1)
For x = 0 To w - 1
buff(x, y) = 255 * Rnd
Next
Next

For y = (h - 3) To 2 Step -1
For x = 1 To (w - 2)
'av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1) + buff(x, y + 2) + buff(x, y + 1)
av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1)

av = av / 3

If av >= 3 Then
av = av - 2
Else
 av = Rnd * 255
 End If

buff(x, y) = av
imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
Next
Next
DoEvents
setbits
fps = fps + 1
Loop

End Sub

Sub runme2()
Dim x As Integer, y As Integer
Dim av As Integer
Dim t As Integer

Dim k1 As Integer, k2 As Integer
k1 = 1
k2 = -1
Do While brun(6)
Randomize
'For y = 200 To 210
'For x = 0 To w - 1
'buff(x, y) = 255 * Rnd
'Next
'Next

'For y = h - 3 To 10 Step -1
For y = h + 2 * k2 To (-k2) Step k2
For x = 2 * k1 To (w - 2 * k1) Step k1
'av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1) + buff(x, y + 2)

'av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1) + buff(x, y + 2) + buff(x, y + 1)
av = 0 + buff(x + k1, y) + buff(x - k1, y) + buff(x, y + k2) + buff(x, y - k2)
av = av / 4
If av >= 3 Then
av = av - Rnd * 15 + 10
'ElseIf av > 1 Then
'av = 255 * Rnd
End If
If av < 0 Then
av = 0
ElseIf av > 255 Then
av = 255
ElseIf av = kk Then
av = av - 4
End If


't = Rnd * 20 - 10
't = fire(2, av) - t
If t < 0 Then t = 0
If buff(x, y) <> kk Then
buff(x, y) = av
Else
av = 0
End If
'ElseIf buff(x, y - 1) = 127 Then
' buff(x, y) = Rnd * 255
'Else
 '   av = 127
'End If


imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
'End If
Next
Next
DoEvents
setbits
fps = fps + 1
Loop

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
stopall
Timer1.Enabled = False
End

End Sub

Private Sub Timer1_Timer()
Me.Caption = "FPS " & fps
fps = 0
End Sub
Sub runme3()
Dim x As Integer, y As Integer
Dim av As Integer
Dim t As Integer

Do While brun(7)
Randomize
For y = hit To hit + 1
For x = 0 To w - 2
  If buff(x, y) <> 0 Then buff(x, y) = 255 * Rnd
Next
Next


'For y = h - 3 To 10 Step -1
For y = hit To 10 Step -1
For x = 1 To w - 2
av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1) + buff(x, y + 2)

'av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1) + buff(x, y + 2) + buff(x, y + 1)

av = av / 4
If av >= 3 Then
av = av - 3
End If
'ElseIf av > 1 Then
'av = 255 * Rnd
If av < 0 Then
av = 0
ElseIf av > 255 Then
av = 255
ElseIf av = kk Then
av = av - 4
End If


't = Rnd * 20 - 10
't = fire(2, av) - t

If buff(x, y) <> kk Then
buff(x, y) = av
Else
av = kk
End If
'ElseIf buff(x, y - 1) = 127 Then
' buff(x, y) = Rnd * 255
'Else
 '   av = 127
'End If

If buff(x, y) <> kk Then
imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
End If
Next
Next
DoEvents
setbits
fps = fps + 1
Loop

End Sub

Sub runme5()
Dim x As Integer, y As Integer
Dim av As Integer

Do While brun(2)
Randomize
For y = (h - 4) To (h - 1)
For x = 0 To w - 1
buff(x, y) = 255 * Rnd
Next
Next

For y = (h - 3) To 2 Step -1
For x = 1 To (w - 2)

av = 0 + buff(x, y - 1) + buff(x, y + 1)
av = av / 2
If av >= 3 Then
av = av - 2
Else
av = Rnd * 255
End If


buff(x, y) = av
imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
Next
Next
DoEvents
setbits
fps = fps + 1
Loop

End Sub

Sub runme6()
Dim x As Integer, y As Integer
Dim av As Integer
Dim a As Integer, b As Integer, c As Long
Do While brun(3)
Randomize
For y = (h - 4) To (h - 1)
For x = 0 To w - 1
buff(x, y) = 255 * Rnd
Next
Next

For y = (h - 3) To 2 Step -1
For x = 1 To (w - 7)

av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1) + buff(x, y + 2) + buff(x, y + 1)
av = av / 5
If av >= 3 Then
av = av - 2
Else
av = Rnd * 255
End If
c = 1
c = c * y * w
c = x + c - h / 5
a = c Mod w
b = c \ h
buff(a, b) = av
imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
Next
Next
DoEvents
setbits
fps = fps + 1
Loop

End Sub

Sub runme7()
Dim x As Integer, y As Integer
Dim av As Integer

Do While brun(4)
Randomize
For y = (h - 4) To (h - 1)
For x = 0 To w - 1
buff(x, y) = 255 * Rnd
Next
Next

For y = (h - 3) To 2 Step -1
For x = 1 To (w - 7)

av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1)
av = av / 3
If av >= 3 Then
av = av - 1
Else
av = Rnd * 255
End If

buff(x, y) = av
If imagedata(2, x, h - y) <> 0 Then

imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
End If
Next
Next
DoEvents
setbits
fps = fps + 1
Loop

End Sub


Sub runme8()
Dim x As Integer, y As Integer
Dim av As Integer

Do While brun(5)
Randomize
For y = (h - 4) To (h - 1)
For x = 0 To w - 1
buff(x, y) = 255 * Rnd
Next
Next

For y = (h - 3) To 2 Step -1
For x = 1 To (w - 2)

av = 0 + buff(x, y - 1) + buff(x, y + 1)
av = av / 2
If av >= 3 Then
av = av - 2
Else
av = Rnd * 255
End If


buff(x, y) = av
If imagedata(2, x, h - y) <> 0 Then
imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
End If
Next
Next

DoEvents
setbits
fps = fps + 1
Loop


End Sub

Sub runme9()
Dim x As Integer, y As Integer
Dim av As Integer

ReDim newarr(w, h)

Dim i As Integer
i = 0

For y = fx1 To fx2
For x = 0 To w - 1
newarr(x, y) = buff(x, y)
DoEvents
Next
Next

Do While brun(8)
Randomize
For y = fx1 To fx2
For x = 0 To w - 1
If newarr(x, y) <> 0 Then buff(x, y) = 255 * Rnd
Next
Next

For y = fx2 To 2 Step -1
For x = 1 To (w - 2)

av = 0 + buff(x - 1, y + 1) + buff(x + 1, y + 1) + buff(x, y + 1)
av = av / 3

av = av - 2
If av <= 0 Then av = 0
'If av >= 3 Then av = av - 2

If newarr(x, y) <> 127 Then
buff(x, y) = av
'If imagedata(2, x, h - y) <> 0 Then
imagedata(0, x, h - y) = fire(2, av)
imagedata(1, x, h - y) = fire(1, av)
imagedata(2, x, h - y) = fire(0, av)
End If
Next
Next

DoEvents
setbits
If i = 0 Then SavePicture Picture1.Image, App.Path & "\ss.bmp"
i = i + 1
fps = fps + 1
Loop


End Sub

Sub stopall()
Dim i As Integer
For i = 0 To 9
brun(i) = False
Next

End Sub
