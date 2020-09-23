Attribute VB_Name = "Internetlink"
Option Explicit

' Internetlänkar

Public Declare Function shellexecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As _
String, ByVal lpDirectory As String, ByVal nShowCmd As _
Long) As Long

Private Const SW_SHOWNORMAL = 1

