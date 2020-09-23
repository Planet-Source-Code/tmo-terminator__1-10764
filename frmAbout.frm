VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Terminator"
   ClientHeight    =   2676
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3504
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2676
   ScaleWidth      =   3504
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "for updates and News."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "http://hem1.passagen.se/tmo"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmAbout.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Visit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "This is FREEWARE."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "by"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "tmo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      MouseIcon       =   "frmAbout.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "mazze"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Created  to"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1020
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    RemoveCancel Me
    Label3.ForeColor = vbBlue
    Label7.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
    Dim lRet As Long
    Dim sText As String
    sText = "mailto:tmo@hem1.passagen.se"
    lRet = shellexecute(hwnd, "open", sText, vbNull, vbNull, SW_SHOWNORMAL)
    If lRet >= 0 And lRet <= 32 Then
        MsgBox "Error!! Can't open Your Mailprog..."
    End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.ToolTipText = "tmo@hem1.passagen.se"
End Sub

Private Sub Label7_Click()
   Dim lRet As Long
    Dim sText As String
    sText = "http://hem1.passagen.se/tmo"
    lRet = shellexecute(hwnd, "open", sText, vbNull, vbNull, SW_SHOWNORMAL)
    If lRet >= 0 And lRet <= 32 Then
        MsgBox "Error!! Can't open Your Mailprog..."
    End If
End Sub
