Attribute VB_Name = "MRegistry"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1 ' Unicode nul terminated String
    Public Const REG_BINARY = 3 ' Binary
    Public Const REG_DWORD = 4 ' 32-bit number

Public Sub savekey(Hkey As Long, strPath As String)
  Dim keyhand&
  r = RegCreateKey(Hkey, strPath, keyhand&)
  r = RegCloseKey(keyhand&)
End Sub

Public Function GetRegistryKey(strValue As String, Optional strLocation As String, Optional vBinary As Boolean, Optional vHKey As Integer)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    Dim strPath As String
    If strLocation = "" Then
      strPath = registryLocation
    Else
      strPath = strLocation
    End If
    
    If vBinary Then
      If vHKey = 1 Then
        r = RegOpenKey(HKEY_CURRENT_USER, strPath, keyhand)
      Else
        r = RegOpenKey(HKEY_LOCAL_MACHINE, strPath, keyhand)
      End If
      lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

      If lValueType = REG_BINARY Then
          strBuf = String(lDataBufSize, " ")
          lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

          If lResult = ERROR_SUCCESS Then
              Dim tempStr$, i!
              For i = 1 To Len(strBuf)
                tempStr = tempStr & Format$(Hex(Asc(Mid$(strBuf, i, 1))), "00") & " "
              Next i
              tempStr = Trim(tempStr)
              
              GetRegistryKey = tempStr
          End If
      End If
    Else
      If vHKey = 1 Then
        r = RegOpenKey(HKEY_CURRENT_USER, strPath, keyhand)
      Else
        r = RegOpenKey(HKEY_LOCAL_MACHINE, strPath, keyhand)
      End If
      lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

      If lValueType = REG_SZ Then
          strBuf = String(lDataBufSize, " ")
          lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

          If lResult = ERROR_SUCCESS Then
              intZeroPos = InStr(strBuf, Chr$(0))

              If intZeroPos > 0 Then
                  GetRegistryKey = Left$(strBuf, intZeroPos - 1)
              Else
                  GetRegistryKey = strBuf
              End If
          End If
      End If
    End If
End Function

Public Sub SetRegistryKey(strValue As String, strdata As Variant, Optional strLocation As String, Optional vBinary As Boolean, Optional vHKey As Integer)
    Dim keyhand As Long
    Dim r As Long
    
    Dim strPath As String
    If strLocation = "" Then
      strPath = registryLocation
    Else
      strPath = strLocation
    End If
    
    If vHKey = 1 Then
      r = RegCreateKey(HKEY_CURRENT_USER, strPath, keyhand)
    Else
      r = RegCreateKey(HKEY_LOCAL_MACHINE, strPath, keyhand)
    End If
    
    If vBinary Then
      r = RegSetValueEx(keyhand, strValue, 0&, REG_BINARY, ByVal CStr(strdata & Chr$(0)), Len(strdata))
    Else
      r = RegSetValueEx(keyhand, strValue, 0&, REG_SZ, ByVal CStr(strdata & Chr$(0)), Len(strdata))
    End If
    r = RegCloseKey(keyhand)
End Sub

Public Function DeleteKey(ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(HKEY_LOCAL_MACHINE, strKey)
End Function

Public Function DeleteValue(ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(HKEY_LOCAL_MACHINE, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

