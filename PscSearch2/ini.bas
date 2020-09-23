Attribute VB_Name = "ini"
Option Compare Text
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function AppPath() As String
AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
End Function
Public Sub PutIni(iniFile As String, iniHead As String, iniKey As String, iniVal As String)
Dim IniFileName As String
IniFileName = AppPath & iniFile
ret = WritePrivateProfileString(iniHead, iniKey, iniVal, IniFileName)
If ret = 0 Then
  Beep
End If
End Sub

Public Function GetIni(iniFile As String, iniHead As String, iniKey As String, iniDefault As String) As String
Dim IniFileName As String
IniFileName = AppPath & iniFile
Dim Temp As String ' set temp string to be as long as the longest value you need to save.
Temp = "                                                                                                                                                                                                                     "
ret = GetPrivateProfileString(iniHead, iniKey, iniDefault, Temp, Len(Temp), IniFileName)
GetIni = Trim(Temp)
GetIni = Left(GetIni, Len(GetIni) - 1)
End Function
