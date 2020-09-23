Attribute VB_Name = "RemoveX"
'
' Funktionen tar bort möljigheterna att stänga av
' applikationen från 'x' i systemmenyn...
' Använd denna modul och från Form_Load()
' skriv RemoveCancelMenuItem Me
'

Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public Sub RemoveCancel(frm As Form)
    Dim hSysMenu As Long
    ' get the system menu for this form
    hSysMenu = GetSystemMenu(frm.hWnd, 0)
    ' remove the close item
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    ' remove the separator the was over the close item
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub
