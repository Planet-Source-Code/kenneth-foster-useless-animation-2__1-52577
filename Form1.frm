VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image5 
      Height          =   450
      Index           =   0
      Left            =   6330
      Picture         =   "Form1.frx":0000
      Top             =   2685
      Width           =   405
   End
   Begin VB.Image Image4 
      Height          =   450
      Index           =   0
      Left            =   5670
      Picture         =   "Form1.frx":054F
      Top             =   2685
      Width           =   405
   End
   Begin VB.Image Image3 
      Height          =   450
      Index           =   0
      Left            =   165
      Picture         =   "Form1.frx":0A9E
      Top             =   2670
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   3
      Left            =   1815
      Picture         =   "Form1.frx":0FED
      Top             =   3465
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   2
      Left            =   1260
      Picture         =   "Form1.frx":153C
      Top             =   3420
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   1
      Left            =   675
      Picture         =   "Form1.frx":1A8E
      Top             =   3420
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   450
      Index           =   0
      Left            =   825
      Picture         =   "Form1.frx":1B64
      Top             =   2670
      Width           =   405
   End
   Begin VB.Label Label3 
      Caption         =   "Use arrow keys for controls."
      Height          =   240
      Left            =   1335
      TabIndex        =   1
      Top             =   2715
      Width           =   2130
   End
   Begin VB.Label Label2 
      Caption         =   "Press spacebar to release box."
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   2955
      Width           =   2265
   End
   Begin VB.Line Line3 
      X1              =   374
      X2              =   451
      Y1              =   178
      Y2              =   178
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   91
      Y1              =   177
      Y2              =   177
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   540
      Left            =   795
      Top             =   2130
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   70
      X2              =   70
      Y1              =   38
      Y2              =   86
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   885
      Picture         =   "Form1.frx":20B3
      Top             =   90
      Width           =   5430
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dd As Boolean  'for control logic
Const x As Integer = 5
Dim y As Integer

Private Sub Form_Load()
dd = False

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case vbKeySpace
 'makes sure box is in right place to release cable
 If Line1.X2 > 53 And Line1.X2 < 394 Then Exit Sub
 If Line1.Y2 <> 147 Then Exit Sub
 'cable release
   Line1.Y2 = Shape1.Top - 10
   dd = True   'disables vbkeyright
   
Case vbKeyRight
  If dd = True Then Exit Sub
  Line1.X1 = Line1.X1 + x
  Line1.X2 = Line1.X2 + x
    If Line1.Y2 >= Shape1.Top Then
       Shape1.Left = Line1.X2
       Shape1.Left = Shape1.Left - (12 + x)
    End If
    'right limit
    If Line1.X1 >= 395 Then
       Line1.X1 = 395
       Line1.X2 = 395
       If Shape1.Left = 53 Then Exit Sub
       Shape1.Left = 378
     End If
     
Case vbKeyLeft
  If Line1.X1 <= 70 Then Exit Sub
  
  Line1.X1 = Line1.X1 - (x * 2)   'faster return
  Line1.X2 = Line1.X2 - (x * 2)  'faster return
    If Line1.Y2 >= Shape1.Top Then
       Shape1.Left = Line1.X2 - 3
       Shape1.Left = Shape1.Left - (x * 3)
    End If
    'move box to right
    If Shape1.Top = 142 Then
       If Shape1.Left = 53 Then Exit Sub
       Shape1.Left = Shape1.Left + x
       y = y + 1
       If y = 4 Then y = 1
       Image2(0).Picture = Image2(y).Picture
       Image3(0).Picture = Image2(y).Picture
       Image4(0).Picture = Image2(y).Picture
       Image5(0).Picture = Image2(y).Picture
    End If
    If Shape1.Left = 458 Then
        Shape1.Left = -30
        If Shape1.BackColor = &H40C0& Then
            Shape1.BackColor = &H80C0FF
        Else
            Shape1.BackColor = &H40C0&
        End If
    End If
    ' left limit
    If Line1.X1 <= 70 Then
       Line1.X1 = 70
       Line1.X2 = 70
       Shape1.Left = 53
       dd = False 'enables arrow keys
    End If
    
Case vbKeyDown
' prevents box from jumping to line(Y2) on return
  If dd = True Then
    Line1.Y2 = 138
    Exit Sub
  End If
   'lower limit
    If Line1.Y2 >= 130 Then
        Line1.Y2 = 142
        Shape1.Top = 130
        dd = False  'enables  arrow keys
     End If
  'lower box
   Line1.Y2 = Line1.Y2 + x
     If Line1.Y2 >= Shape1.Top Then
       Shape1.Top = Line1.Y2
       Shape1.Top = Shape1.Top - x
     End If

Case vbKeyUp
  'upper limit
     If Line1.Y2 <= 70 Then
        Line1.Y2 = 70
     End If
  'raise box
   Line1.Y2 = Line1.Y2 - x
      If Line1.Y2 >= Shape1.Top Then
        Shape1.Top = Line1.Y2
        Shape1.Top = Shape1.Top - x
      End If
 End Select

End Sub


