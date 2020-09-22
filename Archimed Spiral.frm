VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Archimed spiral"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5040
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4800
      Left            =   120
      ScaleHeight     =   4740
      ScaleMode       =   0  'User
      ScaleWidth      =   5664.986
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Pi = 3.14159263

Dim T As Single     'starting point of spiral [0 to infinity]
Dim A As Single     'increment size - spiral trail width

Dim R As Single
Dim Xs As Single:       Dim Ys As Single    'picturebox start point
Dim X0 As Single:       Dim Y0 As Single
Dim X1 As Single:       Dim Y1 As Single

Dim P As Variant

Private Function Draw()
        R = A * T:
        X1 = R * Cos(T) * 2000:    Y1 = R * Sin(T) * 2000
        
        Picture1.Line (Xs + X0, Ys + Y0)-(Xs + X1, Ys + Y1)
        Picture1.Refresh
        'Debug.Print Format(T, "00.0") & "   " & Format(x0, "#.00") & "  " & Format(y0, "#.00") _
                      & "   " & Format(x1, "#.00") & "    " & Format(y1, "#.00")
        X0 = X1:                    Y0 = Y1
End Function

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True:          Timer1.Interval = 1:
    T = 0.5:                        A = 0.02
    Xs = 2700:                      Ys = 2300
    X0 = 0:                         Y0 = 0      'start point of line
End Sub

Private Sub Timer1_Timer()
    Call Draw
    T = T + A
    'stop condition expresed in number of 0 to 180° and 180 to 360° intervals
    If T > (19 * Pi) Then Timer1.Enabled = False
End Sub
