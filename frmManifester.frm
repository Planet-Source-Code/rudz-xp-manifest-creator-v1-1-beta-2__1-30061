VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManifester 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinXP Manifest Creator v1.1 Beta 2"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frmManifester.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   2
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   5295
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdRestore 
         Caption         =   "&Restore"
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Top             =   2400
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstvRestore 
         CausesValidation=   0   'False
         Height          =   2175
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3836
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Application Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Application Path"
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   1
      Left            =   360
      ScaleHeight     =   2775
      ScaleWidth      =   5175
      TabIndex        =   5
      Top             =   720
      Width           =   5175
      Begin VB.Frame Frame2 
         Caption         =   "After manifest"
         Height          =   1455
         Left            =   0
         TabIndex        =   11
         Top             =   840
         Width           =   4695
         Begin VB.CheckBox Check2 
            Caption         =   "E&xit Manifest Creator"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox chkOpenFolder 
            Caption         =   "Open Folder"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Run Program"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Program FileName"
         Height          =   615
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   4695
         Begin VB.TextBox tbxApp 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "The program file to be manifestet"
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   ".."
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   290
         Width           =   375
      End
      Begin VB.CommandButton cmdManifest 
         Caption         =   "&Manifest"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   3
      Left            =   360
      ScaleHeight     =   2775
      ScaleWidth      =   5175
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Frame Frame3 
         Caption         =   "Shell integration"
         Height          =   1215
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   5055
         Begin VB.CheckBox chkCascade 
            Caption         =   "Cascade menu items"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Integrate with shell"
            Height          =   315
            Left            =   240
            TabIndex        =   3
            ToolTipText     =   "Add a cascades explorer 'right-click' feature to enable/disable manifistation of the selected file"
            Top             =   360
            Width           =   1695
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6376
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Style           =   2
      TabFixedWidth   =   2972
      HotTracking     =   -1  'True
      TabMinWidth     =   1235
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Create"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmManifester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Part of this code is ©2001 By NYxZ {nyxz_d2@hotmail.com}
' and to be freely used for anything than commercial use
' Original Concept by : Hectotized

Option Explicit

Private fso As New FileSystemObject

'Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub Check3_Click()

    SaveRegLong HKEY_LOCAL_MACHINE, RegInf.BaseAdr, "ContxMenu", Check3.Value
    Dim sCommandString(1) As String
    sCommandString(0) = "exefile\shell\Disable Visual Styles"
    sCommandString(1) = "exefile\shell\Enable Visual Styles"
    If Check3.Value = 1 Then
With App
        SaveRegString HKEY_CLASSES_ROOT, sCommandString(0) & "\Command", vbNullString, .Path & "\" & .EXEName & ".exe /D%L\" & Chr$(34)
        SaveRegString HKEY_CLASSES_ROOT, sCommandString(1) & "\Command", vbNullString, .Path & "\" & .EXEName & ".exe /E%L\" & Chr$(34)
End With
    Else
        DeleteKey HKEY_CLASSES_ROOT, sCommandString(0) & "\Command"
        DeleteKey HKEY_CLASSES_ROOT, sCommandString(0)
        DeleteKey HKEY_CLASSES_ROOT, sCommandString(1) & "\Command"
        DeleteKey HKEY_CLASSES_ROOT, sCommandString(1)
    End If
    Erase sCommandString()
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show , Me
End Sub

Private Sub cmdBrowse_Click()


    sPath = FILE_DIALOG(Me, False, "Select Executeable", "Executeable (*.Exe)|*.exe", , "*.exe", "c:\")
    If LenB(sPath) <> 0 Then tbxApp.Text = FILE_TITLE_ONLY(sPath)

'    With CommonDialog1
'        .DialogTitle = "Select Executeable"
'        .CancelError = False
'        .FileName = vbNullString
'        .InitDir = "C:\"
'        .Filter = "Executeable (*.Exe)|*.exe"
'        .MaxFileSize = 32000
'        .ShowOpen
'        tbxApp.Text = .FileTitle
'        sPath = .FileName
'    End With

'    If LenB(tbxApp.Text) = 0 Then
'        cmdManifest.Enabled = False
'    Else
'        cmdManifest.Enabled = True
'    End If

End Sub

Private Sub cmdManifest_Click()
    If LenB(tbxApp.Text) = 0 Or LenB(sPath) = 0 Then
        MsgBox "Critical error ocoured!", vbCritical, Me.Caption
        Exit Sub
    End If


    Dim sManifestFile As String
    sManifestFile = sPath & ".manifest"

With fso
    If (.FileExists(sManifestFile)) Then .DeleteFile (sManifestFile)
End With

    Dim ff As Integer

    ff = FreeFile
    Open sManifestFile For Output As #ff
        Print #ff, MF.s1 & ShortFileName(sPath) & Chr$(34)   ' tbx1.Text & ssFileName & Chr$(34)
        Print #ff, MF.s2 ' tbx2.Text
'        Print #ff, Manifest.Mani(0) & tbxApp.Text & Chr$(34) & vbCrLf & Manifest.Mani(1)
    Close #ff

    'Set HIDDEN attribute to .manifest file
    SetAttr sManifestFile, vbHidden
With tbxApp
    ' save to manifest settings '
    SaveRegString HKEY_LOCAL_MACHINE, RegInf.ManifestAdr & .Text, "FileName", .Text
    SaveRegString HKEY_LOCAL_MACHINE, RegInf.ManifestAdr & .Text, "FilePath", sPath

    ' save to win reg '
    SaveRegString HKEY_CURRENT_USER, RegInf.NTLayer, sPath, "WIN2000"

    ' load info to listview '
    lstvRestore.ListItems.Add(, , Left$(.Text, Len(.Text) - 4)).SubItems(1) = sPath

With Me
    If Check1.Value = 1 Then
        ShellExecute .hwnd, vbNullString, sPath, vbNullString, "C:\", SW_SHOWNORMAL
    End If
    If chkOpenFolder.Value = 1 Then
        ShellExecute .hwnd, vbNullString, GetFilePath(sPath), vbNullString, "C:\", SW_SHOWNORMAL
    End If
End With
    .Text = vbNullString
End With
    sPath = vbNullString
    cmdManifest.Enabled = False

    cmdRestore.Enabled = True
    If Check2.Value = 1 Then
        On Error Resume Next
        Unload frmAbout
        Unload Me
    End If
End Sub

Private Sub cmdRestore_Click()
    On Error Resume Next

With lstvRestore.SelectedItem
    If (.Selected) Then
        ' Delete  from registry '
        DeleteKey HKEY_LOCAL_MACHINE, RegInf.ManifestAdr & .Text & ".exe"
        DeleteValue HKEY_CURRENT_USER, RegInf.NTLayer, .SubItems(1)

        ' delete manifest '
        fso.DeleteFile (.SubItems(1) & ".manifest")

        ' Delete from listview '
        lstvRestore.ListItems.Remove (.Index)

        'check to see if list is empty and disable restor button if true
        If lstvRestore.ListItems.Count = 0 Then cmdRestore.Enabled = False
    Else
        MsgBox "Please Select Application To Restore"
    End If
End With

End Sub

Private Sub Form_Load()
    On Error Resume Next
'With fso
'    If (.FileExists(App.Path & "\start.exe")) Then .DeleteFile (App.Path & "\start.exe")
'End With

'    DeleteKey HKEY_CLASSES_ROOT, "exefile\shell\Open Manifest Creator\Command"
'    DeleteKey HKEY_CLASSES_ROOT, "exefile\shell\Open Manifest Creator"


With MF
    .s1 = Chr$(60) & Chr$(63) & Chr$(120) & Chr$(109) & Chr$(108) & Chr$(32) & Chr$(118) & Chr$(101) & Chr$(114) & Chr$(115) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(61) & Chr$(34) & Chr$(49) & Chr$(46) & Chr$(48) & Chr$(34) & Chr$(32) & Chr$(101) & Chr$(110) & Chr$(99) & Chr$(111) & Chr$(100) & Chr$(105) & Chr$(110) & Chr$(103) & Chr$(61) & Chr$(34) & Chr$(85) & Chr$(84) & Chr$(70) & Chr$(45) & Chr$(56) & Chr$(34) & Chr$(32) & Chr$(115) & Chr$(116) & Chr$(97) & Chr$(110) & Chr$(100) & Chr$(97) & Chr$(108) & Chr$(111) & Chr$(110) & Chr$(101) & Chr$(61) & Chr$(34) & Chr$(121) & Chr$(101) & Chr$(115) & Chr$(34) & Chr$(63) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(60) & Chr$(97) & Chr$(115) & Chr$(115) & Chr$(101) & Chr$(109) & Chr$(98) & Chr$(108) & Chr$(121) & Chr$(32) & Chr$(120) & Chr$(109) & Chr$(108) & Chr$(110) & Chr$(115) & Chr$(61) & Chr$(34) & Chr$(117) & Chr$(114) & Chr$(110) & Chr$(58) & Chr$(115) & Chr$(99) & Chr$(104) & Chr$(101) & Chr$(109) & Chr$(97) & Chr$(115) & Chr$(45) & Chr$(109)
    .s1 = .s1 & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(45) & Chr$(99) & Chr$(111) & Chr$(109) & Chr$(58) & Chr$(97) & Chr$(115) & Chr$(109) & Chr$(46) & Chr$(118) & Chr$(49) & Chr$(34) & Chr$(32) & Chr$(109) & Chr$(97) & Chr$(110) & Chr$(105) & Chr$(102) & Chr$(101) & Chr$(115) & Chr$(116) & Chr$(86) & Chr$(101) & Chr$(114) & Chr$(115) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(61) & Chr$(34) & Chr$(49) & Chr$(46) & Chr$(48) & Chr$(34) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(60) & Chr$(97) & Chr$(115) & Chr$(115) & Chr$(101) & Chr$(109) & Chr$(98) & Chr$(108) & Chr$(121) & Chr$(73) & Chr$(100) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(105) & Chr$(116) & Chr$(121) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(118) & Chr$(101) & Chr$(114) & Chr$(115) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(61) & Chr$(34) & Chr$(49) & Chr$(46) & Chr$(48) & Chr$(46) & Chr$(48) & Chr$(46) & Chr$(48) & Chr$(34) & Chr$(13) & Chr$(10)
    .s1 = .s1 & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(112) & Chr$(114) & Chr$(111) & Chr$(99) & Chr$(101) & Chr$(115) & Chr$(115) & Chr$(111) & Chr$(114) & Chr$(65) & Chr$(114) & Chr$(99) & Chr$(104) & Chr$(105) & Chr$(116) & Chr$(101) & Chr$(99) & Chr$(116) & Chr$(117) & Chr$(114) & Chr$(101) & Chr$(61) & Chr$(34) & Chr$(88) & Chr$(56) & Chr$(54) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(110) & Chr$(97) & Chr$(109) & Chr$(101) & Chr$(61) & Chr$(34) ' & Chr$(13) & Chr$(10)
    .s2 = Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(116) & Chr$(121) & Chr$(112) & Chr$(101) & Chr$(61) & Chr$(34) & Chr$(119) & Chr$(105) & Chr$(110) & Chr$(51) & Chr$(50) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(47) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(60) & Chr$(100) & Chr$(101) & Chr$(115) & Chr$(99) & Chr$(114) & Chr$(105) & Chr$(112) & Chr$(116) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(62) & Chr$(87) & Chr$(105) & Chr$(110) & Chr$(100) & Chr$(111) & Chr$(119) & Chr$(115) & Chr$(69) & Chr$(120) & Chr$(101) & Chr$(99) & Chr$(117) & Chr$(116) & Chr$(97) & Chr$(98) & Chr$(108) & Chr$(101) & Chr$(60) & Chr$(47) & Chr$(100) & Chr$(101) & Chr$(115) & Chr$(99) & Chr$(114) & Chr$(105) & Chr$(112) & Chr$(116) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(60) & Chr$(100) & Chr$(101) & Chr$(112) & Chr$(101) & Chr$(110) & Chr$(100) & Chr$(101) & Chr$(110) & Chr$(99) & Chr$(121) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(60)
    .s2 = .s2 & Chr$(100)
    .s2 = .s2 & Chr$(101) & Chr$(112) & Chr$(101) & Chr$(110) & Chr$(100) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(65) & Chr$(115) & Chr$(115) & Chr$(101) & Chr$(109) & Chr$(98) & Chr$(108) & Chr$(121) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(60) & Chr$(97) & Chr$(115) & Chr$(115) & Chr$(101) & Chr$(109) & Chr$(98) & Chr$(108) & Chr$(121) & Chr$(73) & Chr$(100) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(105) & Chr$(116) & Chr$(121) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(116) & Chr$(121) & Chr$(112) & Chr$(101) & Chr$(61) & Chr$(34) & Chr$(119) & Chr$(105) & Chr$(110) & Chr$(51) & Chr$(50) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(110) & Chr$(97) & Chr$(109) & Chr$(101)
    .s2 = .s2 & Chr$(61) & Chr$(34) & Chr$(77) & Chr$(105) & Chr$(99) & Chr$(114) & Chr$(111) & Chr$(115) & Chr$(111) & Chr$(102) & Chr$(116) & Chr$(46) & Chr$(87) & Chr$(105) & Chr$(110) & Chr$(100) & Chr$(111) & Chr$(119) & Chr$(115) & Chr$(46) & Chr$(67) & Chr$(111) & Chr$(109) & Chr$(109) & Chr$(111) & Chr$(110) & Chr$(45) & Chr$(67) & Chr$(111) & Chr$(110) & Chr$(116) & Chr$(114) & Chr$(111) & Chr$(108) & Chr$(115) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(118) & Chr$(101) & Chr$(114) & Chr$(115) & Chr$(105) & Chr$(111) & Chr$(110) & Chr$(61) & Chr$(34) & Chr$(54) & Chr$(46) & Chr$(48) & Chr$(46) & Chr$(48) & Chr$(46) & Chr$(48) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(112) & Chr$(114) & Chr$(111) & Chr$(99) & Chr$(101) & Chr$(115) & Chr$(115)
    .s2 = .s2 & Chr$(111) & Chr$(114) & Chr$(65) & Chr$(114) & Chr$(99) & Chr$(104) & Chr$(105) & Chr$(116) & Chr$(101) & Chr$(99) & Chr$(116) & Chr$(117) & Chr$(114) & Chr$(101) & Chr$(61) & Chr$(34) & Chr$(88) & Chr$(56) & Chr$(54) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(112) & Chr$(117) & Chr$(98) & Chr$(108) & Chr$(105) & Chr$(99) & Chr$(75) & Chr$(101) & Chr$(121) & Chr$(84) & Chr$(111) & Chr$(107) & Chr$(101) & Chr$(110) & Chr$(61) & Chr$(34) & Chr$(54) & Chr$(53) & Chr$(57) & Chr$(53) & Chr$(98) & Chr$(54) & Chr$(52) & Chr$(49) & Chr$(52) & Chr$(52) & Chr$(99) & Chr$(99) & Chr$(102) & Chr$(49) & Chr$(100) & Chr$(102) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(108) & Chr$(97) & Chr$(110) & Chr$(103) & Chr$(117) & Chr$(97) & Chr$(103) & Chr$(101)
    .s2 = .s2 & Chr$(61) & Chr$(34) & Chr$(42) & Chr$(34) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(47) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(60) & Chr$(47) & Chr$(100) & Chr$(101) & Chr$(112) & Chr$(101) & Chr$(110) & Chr$(100) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(65) & Chr$(115) & Chr$(115) & Chr$(101) & Chr$(109) & Chr$(98) & Chr$(108) & Chr$(121) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(60) & Chr$(47) & Chr$(100) & Chr$(101) & Chr$(112) & Chr$(101) & Chr$(110) & Chr$(100) & Chr$(101) & Chr$(110) & Chr$(99) & Chr$(121) & Chr$(62) & Chr$(13) & Chr$(10) & Chr$(60) & Chr$(47) & Chr$(97) & Chr$(115) & Chr$(115) & Chr$(101) & Chr$(109) & Chr$(98) & Chr$(108) & Chr$(121) & Chr$(62) ' & Chr$(13) & Chr$(10)
End With


    'First, we set up the type strings for use with the registry
With RegInf
    .BaseAdr = "Software\NYxZ Software Developement\WinXP Manifest Creator"
    .ManifestAdr = "Software\NYxZ Software Developement\WinXP Manifest Creator\Manifests\"
    .NTLayer = "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"

If LenB(Command$) <> 0 Then
    Dim cmd As String, cmFileName As String, cmCommand As String

    ' ''this code is here for the context menus '
    cmd = Trim$(Command$)
    cmd = Right$(cmd, Len(cmd) - 1)
    cmd = Left$(cmd, Len(cmd) - 1)

    cmCommand = Left$(cmd, 1)
    cmFileName = Right$(cmd, Len(cmd) - 1)
    cmFileName = Left$(cmFileName, Len(cmFileName) - 1)

    If LenB(cmd) <> 0 Then
        Dim sManifestFile As String
        sManifestFile = cmFileName & ".manifest"
        Select Case cmCommand
            Case "E"
                Dim ssFileName As String
                ssFileName = ShortFileName(cmFileName)
With fso
                If (.FileExists(sManifestFile)) Then .DeleteFile (sManifestFile)
End With
                Dim ff As Integer
                ff = FreeFile
                Open sManifestFile For Append As #ff
With MF
                Print #ff, .s1 & ssFileName & Chr$(34)
                Print #ff, .s2
End With
                Close #ff

                ' Set HIDDEN attribute to .manifest file
                SetAttr sManifestFile, vbHidden

                ' save to manifest settings '
                SaveRegString HKEY_LOCAL_MACHINE, .ManifestAdr & ssFileName, "FileName", ssFileName
                SaveRegString HKEY_LOCAL_MACHINE, .ManifestAdr & ssFileName, "FilePath", cmFileName

                ' save to win reg '
                SaveRegString HKEY_CURRENT_USER, .NTLayer, cmFileName, "WIN2000"

            Case "D"
                ' Delete  from registry '
                DeleteKey HKEY_LOCAL_MACHINE, .ManifestAdr & ssFileName
                DeleteValue HKEY_CURRENT_USER, .NTLayer, cmFileName

                ' delete manifest '
                fso.DeleteFile (sManifestFile)

        End Select

        ' Clean up tmp vars
        sManifestFile = vbNullString
        cmd = vbNullString
        ssFileName = vbNullString
        cmFileName = vbNullString
        Set frmManifester = Nothing
        Set frmAbout = Nothing
        Unload Me
        End
    End If
End If

    Dim fApp As String
    Dim fName As String
    Dim fPath As String

    Dim i As Long
    
    For i = 1 To 3
        picTab(i).Left = 240
    Next

    Check3.Value = GetRegLong(HKEY_LOCAL_MACHINE, .BaseAdr, "ContxMenu")

    Dim CountIt As Long
    
    ' load saved alerts to listview '
    CountIt = CountRegKeys(HKEY_LOCAL_MACHINE, .ManifestAdr)

    For i = 0 To CountIt - 1
        fApp = GetRegKey(HKEY_LOCAL_MACHINE, .ManifestAdr, i)
        fName = GetRegString(HKEY_LOCAL_MACHINE, .ManifestAdr & fApp, "FileName")
        fPath = GetRegString(HKEY_LOCAL_MACHINE, .ManifestAdr & fApp, "FilePath")
        lstvRestore.ListItems.Add(, , Left$(fName, Len(fName) - 4)).SubItems(1) = fPath
    Next i


End With

    If lstvRestore.ListItems.Count = 0 Then cmdRestore.Enabled = False

End Sub

Private Sub TabStrip1_Click()
    Dim i As Integer
    For i = 1 To 3
With picTab(i)
        .Visible = False
End With
    Next
With picTab(TabStrip1.SelectedItem.Index)
    .Visible = True
End With
End Sub


Private Sub tbxApp_Change()
With cmdManifest
    If LenB(tbxApp.Text) <> 0 Then .Enabled = True Else .Enabled = False
End With
End Sub

