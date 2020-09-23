VERSION 5.00
Begin VB.Form frmRemove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Programs"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmRemove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numberofprogz As String
Dim progtoremove As String
Dim settings As String
Dim index As Integer
Dim terminatelist(50) As String

Private Sub cmdRemove_Click()
   Dim ret, i As Integer
   Dim strString As String
   If Val(numberofprogz) <= 0 Then
      numberofprogz = Val("0")
      List1.Enabled = False
      Exit Sub
   End If
   ret = MsgBox("Do You want to remove" + vbCr + vbCr + progtoremove + vbCr + vbCr + "from Your terminate list?", vbOKCancel, "Remove...")
   ' programet skall inte tas bort...
   If ret = 2 Then Exit Sub
   'Bort skall det...
   List1.RemoveItem index
   'Bort ur registret oxå...
   Call DeleteKey(HKEY_CURRENT_USER, "Software\tmo\Terminator")
   Call SaveDword(HKEY_CURRENT_USER, TEXT, "Dword", settings)
   numberofprogz = Val(numberofprogz) - Val("1")
   Call SaveDword(HKEY_CURRENT_USER, TEXT, "NumberOfProgz", numberofprogz)
   For i = 0 To numberofprogz
      'If i > 0 Then
         strString = List1.List(i)
         Call savestring(HKEY_CURRENT_USER, TEXT, "prog" & i + 1, strString)
      'End If
   Next i
   Label4 = "You have " + numberofprogz + " programs in Your list."
   cmdRemove.Enabled = False
End Sub

Private Sub Command2_Click()
   Unload Me
   lstProcess.Show
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim strString As String
   List1.Clear
   numberofprogz = getdword(HKEY_CURRENT_USER, TEXT, "NumberOfProgz")
   If numberofprogz = Val("0") Then
      List1.Enabled = False
   End If
   If numberofprogz <> Val("0") Then
      List1.Enabled = True
   End If
   For i = 0 To Val(numberofprogz)
      If i > 0 Then
         strString = getstring(HKEY_CURRENT_USER, TEXT, "prog" & i)
         List1.AddItem strString
      End If
   Next i
   cmdRemove.Enabled = False
   settings = getdword(HKEY_CURRENT_USER, TEXT, "Dword")
   Label3 = "Click on the file You want to remove from your terminate list."
   Label4 = "You have " + numberofprogz + " programs in Your list."
End Sub

Private Sub List1_Click()
   progtoremove = List1
   cmdRemove.Enabled = True
   index = List1.ListIndex
End Sub
