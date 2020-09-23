Attribute VB_Name = "modRegistry"
' Part of this code is ©2001 By NYxZ {nyxz_d2@hotmail.com}
' and to be freely used for anything than commercial use
Option Explicit


Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Attribute RegCloseKey.VB_UserMemId = 1879048224
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Attribute RegCreateKey.VB_UserMemId = 1879048256
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Attribute RegDeleteKey.VB_UserMemId = 1879048292
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Attribute RegDeleteValue.VB_UserMemId = 1879048328
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Attribute RegEnumKey.VB_UserMemId = 1879048364
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Attribute RegQueryValueEx.VB_UserMemId = 1879048432
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Attribute RegSetValueEx.VB_UserMemId = 1879048472

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_SZ = 1            ' Unicode nul terminated string '
Public Const REG_DWORD = 4         ' 32-bit number                 '
Public Const ERROR_SUCCESS = 0&

Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
    Dim lRegResult As Long

    lRegResult = RegDeleteKey(hKey, strPath)

End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim hCurKey As Long

    RegOpenKey hKey, strPath, hCurKey

    RegDeleteValue hCurKey, strValue

    RegCloseKey hCurKey

End Sub

Public Function GetRegString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long

    ' Set up default value '
    If Not IsEmpty(Default) Then
        GetRegString = Default
    Else
        GetRegString = vbNullString
    End If

    ' Open the key and get length of string '
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then

        If lValueType = REG_SZ Then
            ' initialise string buffer and retrieve string '
            strBuffer = String$(lDataBufferSize, Chr$(32))
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)

            ' format string '
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetRegString = Left$(strBuffer, intZeroPos - 1)
            Else
                GetRegString = strBuffer
            End If

        End If

    Else
    ' there is a problem '
        Save2Log lRegResult, "GetRegString()"
    End If

    lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveRegString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

    If lRegResult <> ERROR_SUCCESS Then
    ' there is a problem '
        Save2Log lRegResult, "SaveRegString()"
    End If

    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetRegLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

    Dim lRegResult As Long
    Dim lValueType As Long
    Dim lBuffer As Long
    Dim lDataBufferSize As Long
    Dim hCurKey As Long

    ' Set up default value '
    If Not IsEmpty(Default) Then
        GetRegLong = Default
    Else
        GetRegLong = 0
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4            ' 4 bytes = 32 bits = long '

    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then

        If lValueType = REG_DWORD Then
            GetRegLong = lBuffer
        End If

    Else
    ' there is a problem '
        Save2Log lRegResult, "GetRegLong()"
    End If

    lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveRegLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal Ldata As Long)
    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)

    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, Ldata, 4)

    If lRegResult <> ERROR_SUCCESS Then
    ' there is a problem '
        Save2Log lRegResult, "SaveRegLong()"
    End If

    lRegResult = RegCloseKey(hCurKey)
End Sub


Public Function CountRegKeys(hKey As Long, strPath As String) As Variant
' Returns: an count of all keys '

    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long

    lCounter = 0

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)

    Do

        ' initialise buffers (longest possible length=255) '
        lDataBufferSize = 255
        strBuffer = String$(lDataBufferSize, Chr$(32))
        lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

        If lRegResult = ERROR_SUCCESS Then

            lCounter = lCounter + 1

        Else
            Exit Do
        End If
    Loop

    CountRegKeys = lCounter
End Function

Public Function GetRegKey(hKey As Long, strPath As String, RegKey) As Variant
' Returns: an array in a variant of strings '

    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim strNames() As String

    lCounter = 0

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)

    Do

        ' initialise buffers (longest possible length=255) '
        lDataBufferSize = 255
        strBuffer = String$(lDataBufferSize, Chr$(32))
        lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

        If lRegResult = ERROR_SUCCESS Then

            ' tidy up string and save it '
            ReDim Preserve strNames(lCounter) As String
            If RegKey = lCounter Then
                GetRegKey = strBuffer
                Exit Do
            Else
                lCounter = lCounter + 1
            End If
        Else
            Exit Do
        End If
    Loop

End Function

