Attribute VB_Name = "Registry"
' -----------------
' ADVAPI32
' -----------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

' Registry API prototypes

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

' Own Constants

Public Const TEXT As String = "Software\Tmo\Terminator"


Public Sub savekey(hKey As Long, strPath As String)
Dim keyhand&
r = RegCreateKey(hKey, strPath, keyhand&)
r = RegCloseKey(keyhand&)
End Sub

Public Function getstring(hKey As Long, strPath As String, strValue As String)

Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(hKey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            getstring = Left$(strBuf, intZeroPos - 1)
        Else
            getstring = strBuf
        End If
    End If
End If
End Function


Public Sub savestring(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(hKey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub


Function getdword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long

r = RegOpenKey(hKey, strPath, keyhand)

 ' Get length/data type
lDataBufSize = 4
    
lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        getdword = lBuf
    End If
'Else
'    Call errlog("GetDWORD-" & strPath, False)
End If

r = RegCloseKey(keyhand)
    
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function

Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
Dim r As Long
r = RegDeleteKey(hKey, strKey)
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long
r = RegOpenKey(hKey, strPath, keyhand)
r = RegDeleteValue(keyhand, strValue)
r = RegCloseKey(keyhand)
End Function

'''''''''''''''''''''
'''''''''''''''''''''

'   Så här används funktionerna

'Delete the Dword value
'   Call DeleteValue(HKEY_CURRENT_USER, TEXT, "Dword")

'Delete the Key
'   Call DeleteKey(HKEY_CURRENT_USER, "Software\tmo")

'Delete the string value
'   Call DeleteValue(HKEY_CURRENT_USER, TEXT, "String")

'Get a DWord out the Registry
'   Dim lngDword As String
'   lngDword = getdword(HKEY_CURRENT_USER, TEXT, "Dword")
'   MsgBox lngDword

'Get a String out the Registry
'   Dim strString As String
'   strString = getstring(HKEY_CURRENT_USER, TEXT, "String")
'   MsgBox strString

'Save DWord
'   Dim lngDword As Long
'   Dim strInput As String
'   Ask for String
'   strInput = InputBox("Please type in a string to save to the Registry.")
'   Check to see if the user pressed cancel
'   If strInput = "" Then Exit Sub
'   get long value from this
'   lngDword = CLng(strInput)
'   Edit Registry
'   Call SaveDword(HKEY_CURRENT_USER, TEXT, "Dword", lngDword)

'Save String
'   Dim strString As String
'   Ask for String
'   strString = InputBox("Please type in a string to save to the Registry.")
'   Check to see if the user pressed cancel
'   If strString = "" Then Exit Sub
'   Edit Registry
'   Call savestring(HKEY_CURRENT_USER, TEXT, "String", strString)





