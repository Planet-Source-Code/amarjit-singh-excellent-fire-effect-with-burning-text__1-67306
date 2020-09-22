VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   LinkTopic       =   "Form2"
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2520
      Top             =   3000
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      Height          =   6240
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   6240
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   7
         X1              =   0
         X2              =   6240
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   6
         X1              =   0
         X2              =   6240
         Y1              =   6100
         Y2              =   6100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   5
         X1              =   0
         X2              =   6240
         Y1              =   140
         Y2              =   140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   4
         X1              =   0
         X2              =   6240
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   3
         X1              =   6080
         X2              =   6080
         Y1              =   0
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   2
         X1              =   6100
         X2              =   6100
         Y1              =   0
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   1
         X1              =   80
         X2              =   80
         Y1              =   0
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   0
         X1              =   100
         X2              =   100
         Y1              =   0
         Y2              =   6240
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Me
Form1.Show
End Sub
