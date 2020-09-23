VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "riple"
   ClientHeight    =   390
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   1590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   25
      Left            =   9960
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   9960
      Top             =   1080
   End
   Begin VB.PictureBox Display 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   420
      ScaleHeight     =   3375
      ScaleWidth      =   8595
      TabIndex        =   1
      Top             =   3420
      Width           =   8595
   End
   Begin VB.PictureBox Buffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   540
      ScaleHeight     =   3015
      ScaleWidth      =   8535
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9600
      Y1              =   3180
      Y2              =   3180
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
    
End Sub

Private Sub Form_Load()
    'Me.Display.BackColor = RGB(200, 225, 255)
    Me.Buffer.Width = 1600 * 15
    Me.Display.Width = 1600 * 15
    Me.Display.Height = 200 * 15
    Me.Display.Move 0, Screen.Height - Me.Display.Height
    SetParent Me.Display.hwnd, 0
    DumpToWindow Buffer
    Buffer.Refresh
End Sub

Private Sub Timer1_Timer()
    DumpToWindow Buffer
    Buffer.Refresh
End Sub

Private Sub Timer2_Timer()
    Ripple Buffer, Display
End Sub
