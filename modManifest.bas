Attribute VB_Name = "modManifest"
' Part of this code is ©2001 By NYxZ {nyxz_d2@hotmail.com}
' and to be freely used for anything than commercial use
Option Explicit

Type RegInfo
    BaseAdr As String
    ManifestAdr As String
    NTLayer As String
End Type

Public RegInf As RegInfo

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Type ManiFest
    s1 As String
    s2 As String
End Type
Public MF As ManiFest

Public sPath As String


Public Sub Save2Log(sError As Long, sFunction As String)

    Dim ff As Integer
    ff = FreeFile
With App
    Open .Path & "\ErrorLog.log" For Append As #ff
    Print #ff, vbNewLine
    Print #ff, "--WXPMC_Error-------------------"
    Print #ff, "Version     : " & .Major & "." & .Minor & "." & .Revision
End With
    Print #ff, "Date        : " & Date$
    Print #ff, "Time        : " & Time$
    Print #ff, "Function    : " & sFunction
    Print #ff, "Error       : " & CStr(sError)
    Print #ff, "--------------------------------"
    Close #ff

    MsgBox "Function " & sFunction & " Causes Error : " & sError & vbCr & "The error has been logged.", vbCritical, frmManifester.Caption & " Error!"

End Sub


Public Function GetFilePath(FileName As String, Optional IncludeDrive As Boolean = True) As String
' returns full path. drive can be excluded if needed
    If (IncludeDrive) Then
        GetFilePath = FileName
    Else
        GetFilePath = Right$(FileName, Len(FileName) - 3)
    End If
  Dim i As Integer
'  GetFilePath = FileName     ' Just in case there is no "\" in the file
  For i = 1 To Len(FileName)
    If Mid$(FileName, Len(FileName) - i, 1) = "\" Then
      GetFilePath = Mid$(FileName, 1, Len(FileName) - (i + 1))
      Exit For
    End If
  Next
End Function

Public Function ShortFileName(ByVal sFileName As String) As String
    Dim i As Long
    For i = 0 To Len(sFileName)
        If Left$(Right$(sFileName, i), 1) = "\" Then
            ShortFileName = Right$(sFileName, i - 1)
            i = Len(sFileName)
        End If
    Next i
End Function

'Public Function MultipleChr(ChrValue As Long, Amount As Long) As String
'
'    MultipleChr = vbNullString
'
'    If ChrValue < 0 Or Amount < 1 Then Exit Function
'
'    Select Case Amount
'    Case Is > 1
'        Dim nLoop As Long
'        For nLoop = 1 To Amount
'            MultipleChr = MultipleChr & Chr$(ChrValue)
'        Next nLoop
'    Case 1
'        MultipleChr = Chr$(ChrValue)
'    End Select
'End Function

'Public Function BuildManifest(Optional Part1 As Boolean = False) As String
'    Dim ManiFest(171) As Byte
'    Dim i As Integer
'    For i = 0 To 171
'        Select Case i
'        Case 5, 19, 36, 66, 107, 151 To 154
'            ManiFest(i) = 32
'        Case 0, 57, 132
'            ManiFest(i) = 60
'        Case 1, 53
'            ManiFest(i) = 63
'        Case 2, 67
'            ManiFest(i) = 120
'        Case 3, 62, 68, 82, 86, 98, 102, 108
'            ManiFest(i) = 109
'        Case 4, 43, 64, 69
'            ManiFest(4) = 108
'        Case 6
'            ManiFest(i) = 118
'        Case 7, 20, 46, 50, 61, 81, 113, 117, 136
'            ManiFest(i) = 101
'        Case 8
'            ManiFest(i) = 114
'        Case 9, 37, 51, 59, 60, 71, 78, 91, 101, 114
'            ManiFest(i) = 115
'        Case 10
'            ManiFest(i) = 105
'        Case 11, 23, 44, 90, 92, 97, 121
'            ManiFest(i) = 111
'        Case 12, 21, 26, 40, 45, 70, 76, 110, 122
'            ManiFest(i) = 110
'        Case 13, 28, 47, 72, 123, 162
'            ManiFest(i) = 61
'        Case 14, 18, 29, 35, 48, 52, 73, 106, 124, 128, 163, 171
'            ManiFest(i) = 34
'        Case 15, 105, 125, 164
'            ManiFest(i) = 49
'        Case 16, 103, 126, 165, 167, 169
'            ManiFest(i) = 46
'        Case 17, 127, 166, 168, 170
'            ManiFest(i) = 48
'        Case 22, 76, 88, 96
'            ManiFest(i) = 99
'        Case 24, 41, 142
'            ManiFest(i) = 100
'        Case 25, 87, 111, 120
'            ManiFest(i) = 105
'        Case 27
'            ManiFest(i) = 103
'        Case 30
'            ManiFest(i) = 85
'        Case 31
'            ManiFest(i) = 84
'        Case 32
'            ManiFest(i) = 70
'        Case 33, 85, 95
'            ManiFest(i) = 45
'        Case 34
'            ManiFest(i) = 56
'        Case 38, 94, 115
'            ManiFest(i) = 116
'        Case 39, 42, 58, 83, 100, 109, 133
'            ManiFest(i) = 97
'        Case 49, 65
'            ManiFest(i) = 121
'        Case 54, 129
'            ManiFest(i) = 62
'        Case 55, 130, 149 ', 172
'            ManiFest(i) = 13
'        Case 56, 150 ', 173
'            ManiFest(i) = 10
'        Case 63
'            ManiFest(i) = 98
'        Case 77, 99
'            ManiFest(i) = 58
'        Case 93, 112, 131
'            ManiFest(i) = 102
'        End Select
'    Next i
'ManiFest(141) = 73
'ManiFest(116) = 86
'ManiFest(138) = 98
'ManiFest(143) = 101
'ManiFest(156) = 101
'ManiFest(80) = 104
'ManiFest(146) = 105
'ManiFest(159) = 105
'ManiFest(139) = 108
'ManiFest(137) = 109
'ManiFest(144) = 110
'ManiFest(161) = 110
'ManiFest(160) = 111
'ManiFest(75) = 114
'ManiFest(89) = 114
'ManiFest(118) = 114
'ManiFest(157) = 114
'ManiFest(84) = 115
'ManiFest(119) = 115
'ManiFest(134) = 115
'ManiFest(135) = 115
'ManiFest(158) = 115
'ManiFest(145) = 116
'ManiFest(147) = 116
'ManiFest(74) = 117
'ManiFest(104) = 118
'ManiFest(155) = 118
'ManiFest(140) = 121
'ManiFest(148) = 121

'Chr$ (32 32 32 32 112 114 111 99 101 115 115 111 114 65 114 99 104 105 116 101 99 116 117 114 101 61 34 88 56 54 34 13 10 32 32 32 32 110 97 109 101 61 34) ' & 13 10)



'    .s2 = 32 32 32 32 116 121 112 101 61 34 119 105 110 51 50 34 13 10 47 62 13 10 60 100 101 115 99 114 105 112 116 105 111 110 62 87 105 110 100 111 119 115 69 120 101 99 117 116 97 98 108 101 60 47 100 101 115 99 114 105 112 116 105 111 110 62 13 10 60 100 101 112 101 110 100 101 110 99 121 62 13 10 32 32 32 32 60)
'    .s2 = .s2 & 100)
'    .s2 = .s2 & 101 112 101 110 100 101 110 116 65 115 115 101 109 98 108 121 62 13 10 32 32 32 32 32 32 32 32 60 97 115 115 101 109 98 108 121 73 100 101 110 116 105 116 121 13 10 32 32 32 32 32 32 32 32 32 32 32 32 116 121 112 101 61 34 119 105 110 51 50 34 13 10 32 32 32 32 32 32 32 32 32 32 32 32 110 97 109 101)
'    .s2 = .s2 & 61 34 77 105 99 114 111 115 111 102 116 46 87 105 110 100 111 119 115 46 67 111 109 109 111 110 45 67 111 110 116 114 111 108 115 34 13 10 32 32 32 32 32 32 32 32 32 32 32 32 118 101 114 115 105 111 110 61 34 54 46 48 46 48 46 48 34 13 10 32 32 32 32 32 32 32 32 32 32 32 32 112 114 111 99 101 115 115)
'    .s2 = .s2 & 111 114 65 114 99 104 105 116 101 99 116 117 114 101 61 34 88 56 54 34 13 10 32 32 32 32 32 32 32 32 32 32 32 32 112 117 98 108 105 99 75 101 121 84 111 107 101 110 61 34 54 53 57 53 98 54 52 49 52 52 99 99 102 49 100 102 34 13 10 32 32 32 32 32 32 32 32 32 32 32 32 108 97 110 103 117 97 103 101)
'    .s2 = .s2 & 61 34 42 34 13 10 32 32 32 32 32 32 32 32 47 62 13 10 32 32 32 32 60 47 100 101 112 101 110 100 101 110 116 65 115 115 101 109 98 108 121 62 13 10 60 47 100 101 112 101 110 100 101 110 99 121 62 13 10 60 47 97 115 115 101 109 98 108 121 62) ' & 13 10)





'End Function
