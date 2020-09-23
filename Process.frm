VERSION 5.00
Begin VB.Form lstProcess 
   Caption         =   "Terminator"
   ClientHeight    =   2952
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   8496
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Process.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   708
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   8160
      Top             =   2760
   End
   Begin VB.ListBox lstProcess 
      Height          =   2784
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   6852
   End
   Begin VB.Frame Frame1 
      Caption         =   "Terminate"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   39
         Top             =   1320
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   4800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   38
         Top             =   2280
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   15
         Left            =   3240
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
         Top             =   2280
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   1680
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   34
         Top             =   240
         Width           =   240
         Visible         =   0   'False
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   1680
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   4800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   3240
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   1680
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   4800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   3240
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   1680
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   4800
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         Top             =   840
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   3240
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         Top             =   840
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1680
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         Top             =   840
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   16
         Left            =   5160
         TabIndex        =   36
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   15
         Left            =   3600
         TabIndex        =   35
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   20
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   13
         Left            =   480
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   11
         Left            =   3600
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   4200
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   492
      Left            =   7200
      TabIndex        =   1
      Top             =   2160
      Width           =   1212
   End
   Begin VB.CommandButton cmdKill 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7200
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Program"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Re&move Program"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuTerappz 
         Caption         =   "&Terminate appz"
      End
      Begin VB.Menu mnuTerone 
         Caption         =   "T&erminate one..."
      End
   End
   Begin VB.Menu mnuHjalp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "H&elp"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Te&rminator"
      End
   End
End
Attribute VB_Name = "lstProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Dim activeappz(100) As String
Dim removeappz(16) As String
Dim numberofprogz As String
Dim easter As String

Dim counter, j, C As Integer

Dim blinkthread(16) As Boolean
Dim terminate As Boolean
Dim shutoff As Boolean
Dim blink As Boolean
Dim killappz As Boolean

Private Sub cmdQuit_Click()
   If shutoff = False Then Exit Sub
    ' Stäng ner programmet...
    Dim GotoVal, Gointo
    GotoVal = Me.Height / 2
    Me.Caption = "Shuting down..."
    For Gointo = 1 To GotoVal
    DoEvents
        Me.Height = Me.Height - 200
        Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 11 Then Exit For
    Next Gointo
    Me.Caption = "Bye.."
    Me.Height = 30
    GotoVal = Me.Width / 2
    For Gointo = 1 To GotoVal
        DoEvents
            Me.Width = Me.Width - 100
            Me.Left = (Screen.Width - Me.Width) \ 2
            If Me.Width <= 11 Then End
    Next Gointo
    End
End Sub

Private Sub cmdKill_Click()
    If killappz = True Then
        Dim hProcess As Long
        hProcess = OpenProcess(&H1F0FFF, 1, lstProcess.ItemData(lstProcess.ListIndex))
        TerminateProcess hProcess, 0
        Call Update
        easter = ""
    End If
    If killappz = False Then
        cmdQuit.Enabled = False
        terminate = True
        Timer2.Interval = 0
        easter = ""
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    easter = easter + Chr(KeyCode)
    If Len(easter) > 3 Then easter = ""
    If LCase(easter) = "tmo" Then
        EasterE.Show
        easter = ""
    End If
End Sub

Private Sub Form_Load()
    Dim i, l As Integer
    Dim lngDword, strInput As String
    Label1 = ""
    Label3 = ""
    easter = ""
    
    Timer2.Interval = 200
    
    shutoff = True
    blink = False
    
    For i = 0 To 16
      Label2(i).Visible = False
      Picture1(i).Visible = False
      blinkthread(i) = False
    Next i
    
    'Dessa appzen skall TAS bort...
    numberofprogz = getdword(HKEY_CURRENT_USER, TEXT, "NumberOfProgz")
    For i = 1 To Val(numberofprogz)
      removeappz(i) = getstring(HKEY_CURRENT_USER, TEXT, "prog" & i)
    Next i
    
    Frame1.Caption = "Terminate [" & numberofprogz & " of 16]"
    
    ' Lägg till 0 tkn för utfyllnad...
    ' strängen måste vara 260 tkn lång.
    For l = 1 To Val(numberofprogz)
        For i = Len(removeappz(l)) To 259
            removeappz(l) = removeappz(l) + Chr(0)
        Next i
    Next l
    
    ' Läs in appzen i vektorn
    Call Update
    
    'Kolla inställningar i registret...
    lngDword = getdword(HKEY_CURRENT_USER, TEXT, "Dword")
    If lngDword = Chr(48) Then
        
        ' Det fanns inga inställningar. Sätt Terminate appz som default.
        mnuTerappz.Checked = True
        strInput = "1"
        lngDword = CLng(strInput)
        Call SaveDword(HKEY_CURRENT_USER, TEXT, "Dword", lngDword)
    End If
    
    ' Sätt ut en bock och fixa fönstret...
    'mnuTerappz
    If lngDword = CLng("1") Then
        mnuTerappz.Checked = True
        cmdKill.Caption = "Terminate appz"
        killappz = False
        lstProcess.Visible = False
        Frame1.Visible = True
        Call GetIcon
    ' mnuTerone
    ElseIf lngDword = CLng("2") Then
        mnuTerone.Checked = True
        cmdKill.Caption = "Kill"
        Frame1.Visible = False
        lstProcess.Top = 8
        lstProcess.Left = 8
        lstProcess.Visible = True
        killappz = True
    End If
End Sub

Private Function Update()
    Dim count, i As Integer
    Dim hSnapShot As Long, nProcess As Long
    Dim uProcess As PROCESSENTRY32
    lstProcess.Clear
    hSnapShot = CreateToolhelpSnapshot(2, 0)
    uProcess.dwSize = LenB(uProcess)
    nProcess = Process32First(hSnapShot, uProcess)
    Do While nProcess
        lstProcess.AddItem uProcess.szExeFile
        ' Läs alla aktiva filer i en vektor...
        activeappz(count) = UCase(uProcess.szExeFile)
        lstProcess.ItemData(lstProcess.NewIndex) = uProcess.th32ProcessID
        nProcess = Process32Next(hSnapShot, uProcess)
        For i = 1 To Val(numberofprogz)
            If UCase(activeappz(count)) = UCase(removeappz(i)) Then
               blinkthread(i) = True
            End If
         Next i
         count = count + 1
    Loop
    counter = count
    CloseHandle hSnapShot
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call cmdQuit_Click
End Sub

Private Sub Label2_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Label2(index).ToolTipText = removeappz(index)
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAdd_Click()
   frmAdd.Show
   shutoff = False
   Unload Me
End Sub

Private Sub mnuExit_Click()
    cmdQuit.Value = True
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show
End Sub

Private Sub mnuRemove_Click()
   frmRemove.Show
   shutoff = False
   Unload Me
End Sub

Private Sub mnuTerappz_Click()
    Dim lngDword As String
    mnuTerappz.Checked = True
    mnuTerone.Checked = False
    cmdKill.Caption = "Terminate appz"
    killappz = False
    lstProcess.Visible = False
    Frame1.Visible = True
    lngDword = CLng("1")
    Call SaveDword(HKEY_CURRENT_USER, TEXT, "Dword", lngDword)
    Call GetIcon
End Sub

Private Sub mnuTerone_Click()
    Dim lngDword As String
    mnuTerone.Checked = True
    mnuTerappz.Checked = False
    Frame1.Visible = False
    lstProcess.Top = 8
    lstProcess.Left = 8
    lstProcess.Visible = True
    cmdKill.Caption = "Kill"
    killappz = True
    lngDword = CLng("2")
    Call SaveDword(HKEY_CURRENT_USER, TEXT, "Dword", lngDword)
End Sub

Private Sub Timer1_Timer()
    Dim l, i As Integer
    Dim hProcess As Long
    
    For i = 1 To Val(numberofprogz)
      Label2(i).Visible = True
      Picture1(i).Visible = True
    Next i
    
    If terminate = True Then
        Label1 = "Terminate active process thread...."
        Label3 = activeappz(j)
        For l = 1 To Val(numberofprogz)
            If activeappz(j) = UCase(removeappz(l)) Then
                 C = C + 1
                'Döda processen...
                hProcess = OpenProcess(&H1F0FFF, 1, lstProcess.ItemData(j)) 'removeappz(l)
                TerminateProcess hProcess, 0
                Label2(l).Enabled = False
            End If
        Next l
        j = j + 1
        If j > counter Then
            terminate = False
            Label1 = "Done...   " & C & " of " & Val(numberofprogz) & " appz was terminated!"
            j = 0
            C = 0
            For i = 1 To Val(numberofprogz)
               Label2(i).ForeColor = vbRed
            Next i
        cmdQuit.Enabled = True
        End If
    End If
End Sub

Public Function GetIcon()
    'dim working variables...
    '.. the handle to the system image list
    '.. the file name to get icon from
    '.. the file name filter
    Dim hImgSmall As Long
    Dim hImgLarge As Long
    Dim fName As String
    Dim fnFilter As String
    Dim r As Long
    Dim i As Integer
    'a little error handling to trap a cancel
    On Local Error GoTo cmdLoadErrorHandler
    For i = 1 To Val(numberofprogz)
    
    fName$ = removeappz(i)
    'get the system icons associated with that file
    hImgSmall& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), _
    BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    'hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), _
    BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    'fill in the labels with the image's file data
    Label2(i) = Left$(shinfo.szDisplayName, _
    InStr(shinfo.szDisplayName, Chr$(0)) - 1)
    'Info2 = Left$(shinfo.szTypeName, _
    InStr(shinfo.szTypeName, Chr$(0)) - 1)
    'set set the pictureboxes to receive the icons.
    'Their size must be 16x16 pixels (240x240 twips), for the small i
    '     con, and
    '32x32 pixels (480x480 twips) for the large icon, with no 3d or b
    '     order.
    'Clear any existing image
    Picture1(i).Picture = LoadPicture()
    Picture1(i).AutoRedraw = True
    'pixLarge.Picture = LoadPicture()
    'pixLarge.AutoRedraw = True
    'draw the associated icons into the pictureboxes
    r& = ImageList_Draw(hImgSmall&, shinfo.iIcon, Picture1(i).hDC, 0, 0, ILD_TRANSPARENT)
    'r& = ImageList_Draw(hImgLarge&, shinfo.iIcon, pixLarge.hDC, 0, 0, ILD_TRANSPARENT)
    'realize the images by assigning it's image property
    '(where the icon was drawn) to the actual picture property
    Picture1(i).Picture = Picture1(i).Image
    'pixLarge.Picture = pixLarge.Image
    'Uncomment out the following code to save to the current path
    'Note that the background colour of the icon saved will
    'be the background colour of the pixSmall control.
    ' SavePicture pixSmall, "testSmall.bmp"
    ' SavePicture pixLarge, "testLarge.bmp"
    Next i
    Exit Function
cmdLoadErrorHandler:
    Exit Function
End Function

Private Sub Timer2_Timer()
   Dim i As Integer
   If blink = False Then
      For i = 1 To 16
         If blinkthread(i) = True Then
            Label2(i).ForeColor = &H187A08
         End If
      Next i
      blink = True
      Exit Sub
    End If
    If blink = True Then
      For i = 1 To 16
         If blinkthread(i) = True Then
            Label2(i).ForeColor = vbBlack
         End If
      Next i
      blink = False
    End If
End Sub
