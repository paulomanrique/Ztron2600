VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ZTron 2600"
   ClientHeight    =   2670
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPri.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   2040
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   100
      ImageHeight     =   100
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPri.frx":08CA
            Key             =   "nada"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboDificultyP2 
      Height          =   315
      Left            =   5745
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox cboDificultyP1 
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   80
      Picture         =   "frmPri.frx":1B0D
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   120
      Width           =   1530
   End
   Begin VB.ComboBox cboLang 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1850
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlay 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7200
      Picture         =   "frmPri.frx":2D40
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.ListBox lstGames 
      Height          =   1620
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblDificultyP2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dificuldade P2"
      Height          =   195
      Left            =   4680
      TabIndex        =   9
      Top             =   2235
      Width           =   1005
   End
   Begin VB.Label lblDificultyP1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Dificuldade P1"
      Height          =   195
      Left            =   4695
      TabIndex        =   8
      Top             =   1875
      Width           =   1005
   End
   Begin VB.Label lblSite 
      AutoSize        =   -1  'True
      Caption         =   "http://www.ztron2600.tk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MousePointer    =   15  'Size All
      TabIndex        =   2
      Top             =   2280
      Width           =   2130
   End
   Begin VB.Label lblLanguage 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Idioma:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   540
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOpcoes 
      Caption         =   "&Opções"
      Begin VB.Menu mnuDir 
         Caption         =   "&Diretórios"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuControles 
         Caption         =   "&Controles"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuVideo 
         Caption         =   "&Vídeo"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnuSite 
         Caption         =   "http://www.ztron2600.tk"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&Sobre..."
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmPri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare General ini variables
Public iniRomDir, iniZ26Dir, iniLanguage, iniDificultyP1, iniDificultyP2 As String
Dim cacete As String

'Declare Controler ini variables
Public iniReverseJoystick, iniAllowAllDirections, iniEnableMindlink, iniMindlinkSide, iniMousePaddleEnable, iniMousePaddleId, iniMouseTwoPaddleEnable, iniMouseXPaddle, iniMouseYPaddle, iniKeyboardPaddleEnable, iniKeyboardPaddleId, iniKeyboardPaddleSensitivity, iniLightgunEnable, iniLightgunCycles, iniLightgunAdjustByScanlines, iniLightgunAdjustByScanlinesMuch As String

'Declare Video ini variables
Public iniInterlacedGames, iniLowResolutions, iniPhosphorescentEffect, iniPhosphorescentEffectCount, iniShowFPS, iniScanlines, iniSimulateColorLossPal, iniVsync, iniFullScreen, iniColorPalete, iniFPS, iniMode As String
Dim CommandLine As String

'Declare Language variables
Public String01, String02, String03, String04, String05, String06, String07, String08, String09, String10, String11, String12, String13, String14, String15, String16, String17, String18, String19, String20, String21, String22, String23, String24, String25 As String
Public String26, String27, String28, String29, String30, String31, String32, String33, String34, String35, String36, String37, String38, String39, String40, String41, String42, String43, String44, String45, String46, String47, String48, String49, String50 As String
Public String51, String52, String53, String54, String55, String56, String57, String58, String59 As String

Dim fso As New FileSystemObject
Dim fld As Folder
Dim ShortAppPath, li As String
Const m_def_CompressedSize = 0
Const m_def_OriginalSize = 0
Dim m_CompressedSize As Long
Dim m_OriginalSize As Long
Private Const SW_SHOW = 5
Private Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
   String, ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function FindExecutable Lib "shell32.dll" Alias _
   "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
   String, ByVal lpResult As String) As Long


Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
        (ByVal lpszLongPath As String, _
         ByVal lpszShortPath As String, _
         ByVal cchBuffer As Long) As Long
Private Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    GetShortName = Left(sShortPathName, lRetVal)
End Function
Private Function FindFile(ByVal sFol As String, sFile As String, _
        nDirs As Integer, nFiles As Integer) As Long
On Error GoTo hell
        Dim tFld As Folder, tFil As File, FileName As String

        Set fld = fso.GetFolder(sFol)
        FileName = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
          vbHidden Or vbSystem Or vbReadOnly)
          While Len(FileName) <> 0
            FindFile = FindFile + FileLen(fso.BuildPath(fld.Path, _
                       FileName))
            nFiles = nFiles + 1
                lstGames.AddItem FileName
            FileName = Dir()
               DoEvents
            Wend
       nDirs = nDirs + 1
               
hell:
    Exit Function
End Function

Private Sub cboLang_Click()
    Call CarregaIdioma
End Sub

Private Sub cmdPlay_Click()
    If lstGames.ListIndex <> -1 Then
        Play
    Else
        MsgBox String04, vbCritical, String53
    End If
End Sub
Private Sub Form_Load()
    CarregaPadroes
    Dim nDirs As Integer, nFiles As Integer, lSize As Long
    Dim sDir As String, sSrchString As String
    sDir = App.Path & "\language"
    sSrchString = "*.ini"
    lSize = FindIni(sDir, sSrchString, nDirs, nFiles)
    CarregaIdioma
    AtualizaLista
   
    Me.Caption = "ZTron 2600 " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If fso.FolderExists(App.Path & "\_tmp") Then
        fso.DeleteFolder App.Path & "\_tmp", True
    End If
    Call fWriteValue(App.Path & "\ztron2600.ini", "General", "DificultyP1", "S", cboDificultyP1.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "General", "DificultyP2", "S", cboDificultyP2.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "General", "Language", "S", cboLang.Text)
    End
End Sub

Private Sub lblSite_Click()
      Dim FileName, Dummy As String
      Dim BrowserExec As String * 255
      Dim RetVal As Long
      Dim FileNumber As Integer

      BrowserExec = Space(255)
      FileName = "C:\temphtm.HTM"
      FileNumber = FreeFile
      Open FileName For Output As #FileNumber
          Write #FileNumber, "<HTML> <\HTML>"
      Close #FileNumber
      RetVal = FindExecutable(FileName, Dummy, BrowserExec)
      BrowserExec = Trim(BrowserExec)
      If RetVal <= 32 Or IsEmpty(BrowserExec) Then
          MsgBox "Could not find associated Browser", vbExclamation, _
            "Browser Not Found"
      Else
          RetVal = ShellExecute(Me.hwnd, "open", BrowserExec, _
            "http://www.ztron2600.tk", Dummy, SW_SHOWNORMAL)
          If RetVal <= 32 Then
              MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
          End If
      End If
      Kill FileName
End Sub

Private Sub lstGames_DblClick()
    Play
End Sub

Private Sub mnuAbout_Click()
    MsgBox "ZTron 2600 " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & String47 & vbCrLf & String48 & vbCrLf & String49 & vbCrLf & String50, vbInformation, String12
End Sub

Private Sub mnuControles_Click()
    frmControles.Show vbModal, Me
End Sub

Private Sub mnuDir_Click()
    frmDir.Show vbModal, Me
    AtualizaLista
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub
Function AtualizaLista()
    Screen.MousePointer = 11
    Me.Enabled = False
    lstGames.Clear
    Dim nDirs As Integer, nFiles As Integer, lSize As Long
    Dim sDir As String, sSrchString As String
    sDir = iniRomDir
    sSrchString = "*.zip"
    lSize = FindFile(sDir, sSrchString, nDirs, nFiles)
    sSrchString = "*.7z"
    lSize = FindFile(sDir, sSrchString, nDirs, nFiles)
    sSrchString = "*.bin"
    lSize = FindFile(sDir, sSrchString, nDirs, nFiles)
    sSrchString = "*.a26"
    lSize = FindFile(sDir, sSrchString, nDirs, nFiles)
    Screen.MousePointer = 0
    Me.Enabled = True
End Function
Function Play()
    Call MontaLinhaComando
    Dim nDirs As Integer, nFiles As Integer, lSize As Long
    Dim sDir As String, sSrchString As String
    If Right(lstGames.Text, 4) = ".zip" Then
        extARchive iniRomDir & "\" & lstGames.Text, App.Path & "\_tmp"
        sDir = App.Path & "\_tmp"
        sSrchString = "*.bin"
        lSize = FindRom(sDir, sSrchString, nDirs, nFiles)
        sSrchString = "*.a26"
        lSize = FindRom(sDir, sSrchString, nDirs, nFiles)
        Shell iniZ26Dir & "\z26.exe """ & sDir & "\" & cacete & """" & CommandLine
    ElseIf Right(lstGames.Text, 3) = ".7z" Then
        Shell """" & App.Path & "\7z.exe""" & " e -o""" & App.Path & "\_tmp"" " & """" & iniRomDir & "\" & lstGames.Text & """" & " *.*", vbHide
        Timer1.Enabled = True
    Else
        Shell iniZ26Dir & "\z26.exe """ & iniRomDir & "\" & lstGames.Text & """" & CommandLine
    End If
End Function
Function CarregaPadroes()
'General
    Call fReadValue(App.Path & "\ztron2600.ini", "General", "RomDir", "S", "", iniRomDir)
    Call fReadValue(App.Path & "\ztron2600.ini", "General", "Z26Dir", "S", "", iniZ26Dir)
    Call fReadValue(App.Path & "\ztron2600.ini", "General", "Language", "S", "", iniLanguage)
    Call fReadValue(App.Path & "\ztron2600.ini", "General", "DificultyP1", "S", "", iniDificultyP1)
    Call fReadValue(App.Path & "\ztron2600.ini", "General", "DificultyP2", "S", "", iniDificultyP2)

'Controlers
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "ReverseJoystick", "S", "", iniReverseJoystick)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "AllowAllDirections", "S", "", iniAllowAllDirections)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "EnableMindlink", "S", "", iniEnableMindlink)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "MindlinkSide", "S", "", iniMindlinkSide)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "MousePaddleEnable", "S", "", iniMousePaddleEnable)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "MousePaddleId", "S", "", iniMousePaddleId)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "MouseTwoPaddleEnable", "S", "", iniMouseTwoPaddleEnable)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "MouseXPaddle", "S", "", iniMouseXPaddle)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "MouseYPaddle", "S", "", iniMouseYPaddle)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "KeyboardPaddleEnable", "S", "", iniKeyboardPaddleEnable)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "KeyboardPaddleId", "S", "", iniKeyboardPaddleId)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "KeyboardPaddleSensitivity", "S", "", iniKeyboardPaddleSensitivity)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunEnable", "S", "", iniLightgunEnable)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunCycles", "S", "", iniLightgunCycles)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunAdjustByScanlines", "S", "", iniLightgunAdjustByScanlines)
    Call fReadValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunAdjustByScanlinesMuch", "S", "", iniLightgunAdjustByScanlinesMuch)

'Video
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "InterlacedGames", "S", "", iniInterlacedGames)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "LowResolutions", "S", "", iniLowResolutions)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "PhosphorescentEffect", "S", "", iniPhosphorescentEffect)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "PhosphorescentEffectCount", "S", "", iniPhosphorescentEffectCount)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "ShowFPS", "S", "", iniShowFPS)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "Scanlines", "S", "", iniScanlines)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "SimulateColorLossPal", "S", "", iniSimulateColorLossPal)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "Vsync", "S", "", iniVsync)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "FullScreen", "S", "", iniFullScreen)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "ColorPalete", "S", "", iniColorPalete)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "FPS", "S", "", iniFPS)
    Call fReadValue(App.Path & "\ztron2600.ini", "Video", "Mode", "S", "", iniMode)

End Function

Private Sub mnuSite_Click()
      Dim FileName, Dummy As String
      Dim BrowserExec As String * 255
      Dim RetVal As Long
      Dim FileNumber As Integer

      BrowserExec = Space(255)
      FileName = "C:\temphtm.HTM"
      FileNumber = FreeFile
      Open FileName For Output As #FileNumber
          Write #FileNumber, "<HTML> <\HTML>"
      Close #FileNumber
      RetVal = FindExecutable(FileName, Dummy, BrowserExec)
      BrowserExec = Trim(BrowserExec)
      If RetVal <= 32 Or IsEmpty(BrowserExec) Then
          MsgBox "Could not find associated Browser", vbExclamation, _
            "Browser Not Found"
      Else
          RetVal = ShellExecute(Me.hwnd, "open", BrowserExec, _
            "http://www.ztron2600.tk", Dummy, SW_SHOWNORMAL)
          If RetVal <= 32 Then
              MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
          End If
      End If
      Kill FileName
End Sub

Private Sub mnuVideo_Click()
    frmVideo.Show vbModal, Me
End Sub
Public Function extARchive(aPath As String, extPath As String)
Dim bzip As CGUnzipFiles
Set bzip = New CGUnzipFiles

With bzip
    .Unzip aPath, extPath
End With
End Function

Private Function FindRom(ByVal sFol As String, sFile As String, _
        nDirs As Integer, nFiles As Integer) As Long
On Error GoTo hell
        Dim tFld As Folder, tFil As File, FileName As String

        Set fld = fso.GetFolder(sFol)
        FileName = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
          vbHidden Or vbSystem Or vbReadOnly)
          While Len(FileName) <> 0
            FindRom = FindRom + FileLen(fso.BuildPath(fld.Path, _
                       FileName))
            nFiles = nFiles + 1
                cacete = FileName
            FileName = Dir()
               DoEvents
            Wend
       nDirs = nDirs + 1
               
hell:
    Exit Function
End Function

Private Function FindIni(ByVal sFol As String, sFile As String, _
        nDirs As Integer, nFiles As Integer) As Long
        Dim tFld As Folder, tFil As File, FileName As String

        Set fld = fso.GetFolder(sFol)
        FileName = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
          vbHidden Or vbSystem Or vbReadOnly)
          While Len(FileName) <> 0
            FindIni = FindIni + FileLen(fso.BuildPath(fld.Path, _
                       FileName))
            nFiles = nFiles + 1
                cboLang.AddItem Replace(FileName, ".ini", "")
                            FileName = Replace(FileName, ".ini", "")
            If iniLanguage = FileName Then
                    langid = (cboLang.ListCount)
            End If

            FileName = Dir()
               DoEvents
            Wend
            cboLang.ListIndex = (langid - 1)
       nDirs = nDirs + 1
End Function

Function MontaLinhaComando()
'Misc
    Dim clDificultyp1, clDificultyp2 As String
    If cboDificultyP1.ListIndex = 1 Then clDificultyp1 = " -0"
    If cboDificultyP2.ListIndex = 1 Then clDificultyp2 = " -1"
    CommandLine = clDificultyp1 & clDificultyp2
    
 'Controlers
    Dim clReverseJoystick, clAllowAllDirections, clEnableMindlink, clMousePaddleEnable, clMouseTwoPaddleEnable, clKeyboardPaddleEnable, clLightgunEnable, clLightgunAdjustByScanlines As String
    If iniReverseJoystick = 1 Then clReverseJoystick = " -J1"
    If iniAllowAllDirections = 1 Then clAllowAllDirections = " -4"
    If iniEnableMindlink = 1 Then clEnableMindlink = " j" & (iniMindlinkSide + 1)
    If iniMousePaddleEnable = 1 Then clMousePaddleEnable = " m" & iniMousePaddleId
    If iniMouseTwoPaddleEnable = 1 Then clMouseTwoPaddleEnable = " m1" & iniMouseXPaddle & iniMouseYPaddle
    If iniKeyboardPaddleEnable = 1 Then clKeyboardPaddleEnable = " k" & iniKeyboardPaddleId & " -p" & iniKeyboardPaddleSensitivity
    If iniLightgunEnable = 1 Then clLightgunEnable = " l" & iniLightgunCycles
    If iniLightgunAdjustByScanlines = 1 Then clLightgunAdjustByScanlines = " -a" & iniLightgunAdjustByScanlinesMuch
    CommandLine = CommandLine & clReverseJoystick & clAllowAllDirections & clEnableMindlink & clMousePaddleEnable & clMouseTwoPaddleEnable & clKeyboardPaddleEnable & clLightgunEnable & clLightgunAdjustByScanlines
'Video
    Dim clInterlacedGames, clLowResolutions, clPhosphorescentEffect, clShowFPS, clScanlines, clSimulateColorLossPal, clVsync, clFullScreen, clColorPalete, clFPS As String
    If iniInterlacedGames = 1 Then clInterlacedGames = " -!"
    If iniLowResolutions = 1 Then clLowResolutions = " -e1"
    If iniPhosphorescentEffect = 1 Then clPhosphorescentEffect = " -f" & iniPhosphorescentEffectCount
    If iniShowFPS = 1 Then clShowFPS = " -n"
    
    If iniScanlines <> 0 Then clScanlines = " -h" & iniScanlines
    If iniSimulateColorLossPal = 1 Then clSimulateColorLossPal = " -o"
    If iniVsync = 0 Then clVsync = " -r"
    If iniFullScreen = 0 Then
        clFullScreen = " -v1" & iniMode
    Else
        clFullScreen = " -v" & iniMode
    End If
    clColorPalete = " -c" & iniColorPalete
    clFPS = " -r" & iniFPS
    CommandLine = CommandLine & clInterlacedGames & clLowResolutions & clPhosphorescentEffect & clShowFPS & clScanlines & clSimulateColorLossPal & clVsync & clFullScreen & clColorPalete & clFPS
End Function
Function CarregaIdioma()
    iniLanguage = cboLang.Text
    Dim tmp As Variant
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String01", "S", "", String01)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String02", "S", "", String02)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String03", "S", "", String03)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String04", "S", "", String04)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String05", "S", "", String05)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String06", "S", "", String06)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String07", "S", "", String07)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String08", "S", "", String08)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String09", "S", "", String09)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String10", "S", "", String10)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String11", "S", "", String11)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String12", "S", "", String12)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String13", "S", "", String13)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String14", "S", "", String14)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String15", "S", "", String15)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String16", "S", "", String16)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String17", "S", "", String17)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String18", "S", "", String18)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String19", "S", "", String19)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String20", "S", "", String20)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String21", "S", "", String21)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String22", "S", "", String22)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String23", "S", "", String23)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String24", "S", "", String24)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String25", "S", "", String25)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String26", "S", "", String26)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String27", "S", "", String27)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String28", "S", "", String28)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String29", "S", "", String29)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String30", "S", "", String30)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String31", "S", "", String31)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String32", "S", "", String32)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String33", "S", "", String33)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String34", "S", "", String34)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String35", "S", "", String35)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String36", "S", "", String36)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String37", "S", "", String37)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String38", "S", "", String38)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String39", "S", "", String39)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String40", "S", "", String40)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String41", "S", "", String41)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String42", "S", "", String42)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String43", "S", "", String43)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String44", "S", "", String44)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String45", "S", "", String45)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String46", "S", "", String46)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String47", "S", "", String47)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String48", "S", "", String48)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String49", "S", "", String49)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String50", "S", "", String50)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String51", "S", "", String51)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String52", "S", "", String52)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String53", "S", "", String53)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String54", "S", "", String54)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String55", "S", "", String55)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String56", "S", "", String56)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String57", "S", "", String57)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String58", "S", "", String58)
    Call fReadValue(App.Path & "\language\" & iniLanguage & ".ini", "Language", "String59", "S", "", String59)
    
    lblLanguage = String01
    lblDificultyP1 = String02
    lblDificultyP2 = String03
    tmp = cboDificultyP1.ListIndex
    cboDificultyP1.Clear
    cboDificultyP1.AddItem String51
    cboDificultyP1.AddItem String52
    cboDificultyP1.ListIndex = tmp
    tmp = cboDificultyP2.ListIndex
    cboDificultyP2.Clear
    cboDificultyP2.AddItem String51
    cboDificultyP2.AddItem String52
    cboDificultyP2.ListIndex = tmp
    cboDificultyP1.ListIndex = iniDificultyP1
    cboDificultyP2.ListIndex = iniDificultyP2
    mnuArquivo.Caption = String05
    mnuSair.Caption = String06
    mnuOpcoes.Caption = String07
    mnuDir.Caption = String08
    mnuControles.Caption = String09
    mnuVideo.Caption = String10
    mnuAjuda.Caption = String11
    mnuAbout.Caption = String12
End Function

Private Sub Timer1_Timer()
    Dim nDirs As Integer, nFiles As Integer, lSize As Long
    Dim sDir As String, sSrchString As String
        sDir = App.Path & "\_tmp"
        sSrchString = "*.bin"
        lSize = FindRom(sDir, sSrchString, nDirs, nFiles)
        sSrchString = "*.a26"
        lSize = FindRom(sDir, sSrchString, nDirs, nFiles)
        Shell iniZ26Dir & "\z26.exe """ & sDir & "\" & cacete & """" & CommandLine
        Timer1.Enabled = False
End Sub
