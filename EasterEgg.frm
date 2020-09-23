VERSION 5.00
Begin VB.Form EasterE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EasterEgg Terminator..."
   ClientHeight    =   3756
   ClientLeft      =   2376
   ClientTop       =   1896
   ClientWidth     =   5184
   Icon            =   "EasterEgg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3756
   ScaleWidth      =   5184
   Begin VB.PictureBox pctEE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3012
      Left            =   960
      ScaleHeight     =   3012
      ScaleWidth      =   3492
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Timer Coloring 
      Interval        =   1
      Left            =   120
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   2640
   End
End
Attribute VB_Name = "EasterE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const Tp1 = 600
Private Const Tp2 = 1320
Private Const Lft = 960

Dim pTop As Integer
Dim Secret, Moveme As Integer
Dim i
Dim NL As String
Dim Shw As Boolean

Public Sub EasterEgg()
    Timer1.Interval = 100
    Timer1.Enabled = True
    NL = Chr(13) & Chr(10)
    pctEE.Cls
    pTop = pctEE.Height
End Sub

Private Sub Coloring_Timer()
Static Turn, nex, color, stat, f, g
    f = f + 0.000005
    If f = 0.000005 Then Turn = True: stat = "ascend"
    If Turn = True Then
        nex = nex + 1
        If nex >= 4 Then nex = 1
        If nex = 1 Then color = "Red"
        If nex = 2 Then color = "Green"
        If nex = 3 Then color = "Blue"
        stat = "ascend"
        Turn = False
    End If
    If stat = "ascend" Then
        g = g + 1
        If g = 255 Then stat = "decend"
    ElseIf stat = "decend" Then
        g = g - 1
        If g = 0 Then Turn = True
    End If
    'If color = "Red" Then pctEE.ForeColor = RGB(g, 0, 0)
    'If color = "Green" Then pctEE.ForeColor = RGB(0, g, 0)
    'If color = "Blue" Then pctEE.ForeColor = RGB(0, 0, g)
    If Shw = False Then
        pctEE.CurrentX = 120: pctEE.CurrentY = pTop
        pctEE.Print "Program created by" & NL: pctEE.CurrentX = 120: pctEE.Print "tmo (programmer)" & NL: pctEE.CurrentX = 120: pctEE.Print "tmo Software (1999)" & NL: pctEE.CurrentX = 120: pctEE.Print "General Idea by" & NL: pctEE.CurrentX = 120: pctEE.Print "Mazze"
        Exit Sub
    End If
    pctEE.CurrentX = 120: pctEE.CurrentY = 360
    pctEE.Print App.ProductName: pctEE.CurrentX = 120: pctEE.Print "Version " & App.Major & "." & App.Minor & "." & App.Revision & NL: pctEE.CurrentX = 120: pctEE.Print "Copyright (C) 10-10-1999" & NL: pctEE.CurrentX = 120: pctEE.Print "FREEWARE": pctEE.CurrentX = 120: pctEE.Print "(Press Spacebar For Credits.)"
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    If Shw = True Then
        Shw = False
        EasterEgg
    End If
End Sub

Private Sub Form_Load()
    CF
    i = True
    Moveme = 50
    NL = Chr(13) & Chr(10)
    Shw = True
    Show
End Sub
Sub CF()
    Left = (Screen.Width \ 2) - (Width \ 2)
    Top = (Screen.Height \ 2) - (Height \ 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    EasterE.Hide
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    Secret = Secret + KeyAscii
    If Secret = 32 Then
        Label2.Top = Picture1.Height
        EasterEgg
    End If
End Sub

Private Sub pctEE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    If Shw = True Then
        Shw = False
        EasterEgg
    End If
End Sub

Private Sub Timer1_Timer()
Static j
    If i = True Then pctEE.Cls: pTop = pTop - Moveme
    If pTop < 150 Then
        Moveme = -Moveme
        i = False
        j = j + 1
        If j >= 5 Then i = True: j = 0
    End If
    If pTop > pctEE.Height Then
        Moveme = -Moveme
        Timer1.Enabled = False
        pTop = pctEE.Height
        Shw = True
        Secret = 0
    End If
End Sub


