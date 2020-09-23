VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Programs"
   ClientHeight    =   5040
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5340
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1608
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   5052
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numberofprogz As String
Dim terminatelist(50) As String

Private Sub cmdAdd_Click()
   Dim i, ret As Integer
   Dim strString As String
   ' kolla så inte filen redan finns
   strString = Dir1 + "\" + File1
   For i = 0 To Val(numberofprogz)
      If strString = terminatelist(i) Then
         ret = MsgBox("This file already exist in Your terminate list!", vbExclamation, "Already exist")
         Exit Sub
      End If
   Next i
   'kolla så att det inte är mer än 16...
   If Val(numberofprogz) + Val("1") > 16 Then
      ret = MsgBox("You are limited to 16 programs...", vbExclamation, "To many programs")
      Exit Sub
   End If
   'nope... lägg till den...
   numberofprogz = numberofprogz + Val("1")
   Call SaveDword(HKEY_CURRENT_USER, TEXT, "NumberOfProgz", numberofprogz)
   terminatelist(Val(numberofprogz)) = Dir1 + "\" + File1
   Call savestring(HKEY_CURRENT_USER, TEXT, "prog" & Val(numberofprogz), terminatelist(Val(numberofprogz)))
   Label3 = "You have " + numberofprogz + " programs in Your terminate list."
   cmdAdd.Enabled = False
End Sub

Private Sub Command2_Click()
   lstProcess.Show
   Unload Me
End Sub

Private Sub Dir1_Change()
   File1 = Dir1
End Sub

Private Sub Drive1_Change()
    Dir1 = Drive1
End Sub

Private Sub File1_Click()
   Label1 = Dir1 + "\" + File1
   cmdAdd.Enabled = True
End Sub

Private Sub Form_Load()
   Dim i As Integer
   numberofprogz = getdword(HKEY_CURRENT_USER, TEXT, "NumberOfProgz")
   If numberofprogz = "0" Then
      numberofprogz = CLng("0")
      Call SaveDword(HKEY_CURRENT_USER, TEXT, "NumberOfProgz", numberofprogz)
   End If
   If numberofprogz <> "0" Then
      For i = 0 To Val(numberofprogz)
         terminatelist(i) = getstring(HKEY_CURRENT_USER, TEXT, "prog" & i)
      Next i
   End If
   Label1 = "No program has been choosen!"
   Label2 = "Click on the Program You want to add to terminate list."
   Label3 = "You have " + numberofprogz + " programs in Your terminate list."
   cmdAdd.Enabled = False
End Sub
