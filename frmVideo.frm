VERSION 5.00
Begin VB.Form frmVideo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vídeo"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkInterlacedGames 
      Caption         =   "Rodar jogos entrelaçados"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Value           =   2  'Grayed
      Width           =   2295
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame frameVideoCFG 
      Caption         =   "Configurações de vídeo"
      Height          =   5175
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   4560
         Width           =   4095
         Begin VB.ComboBox cboScreenMode 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   120
            Width           =   2550
         End
         Begin VB.Label lblScreenMode 
            AutoSize        =   -1  'True
            Caption         =   "Modo de tela:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   150
            Width           =   990
         End
      End
      Begin VB.Frame Frame5 
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   4095
         Begin VB.TextBox txtFps 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   10
            Text            =   "30"
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblFPS 
            AutoSize        =   -1  'True
            Caption         =   "Frames por segundo:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   150
            Width           =   1530
         End
      End
      Begin VB.Frame Frame4 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   4095
         Begin VB.ComboBox cboColorPalete 
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   120
            Width           =   1230
         End
         Begin VB.Label lblColorPalete 
            AutoSize        =   -1  'True
            Caption         =   "Paleta de cores"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   150
            Width           =   1110
         End
      End
      Begin VB.CheckBox chkFullScreen 
         Caption         =   "Tela inteira"
         Height          =   195
         Left            =   200
         TabIndex        =   8
         Top             =   3120
         Value           =   2  'Grayed
         Width           =   1095
      End
      Begin VB.CheckBox chkRunMonitorSpeed 
         Caption         =   "Rodar na velocidade do monitor"
         Height          =   195
         Left            =   200
         TabIndex        =   7
         Top             =   2760
         Value           =   2  'Grayed
         Width           =   2655
      End
      Begin VB.CheckBox chkSimulateLostPal 
         Caption         =   "Simular perda de cores no modo PAL"
         Height          =   195
         Left            =   190
         TabIndex        =   6
         Top             =   2400
         Value           =   2  'Grayed
         Width           =   3255
      End
      Begin VB.CheckBox chkShowFPS 
         Caption         =   "Mostrar contador de scanlines e de FPS"
         Height          =   195
         Left            =   185
         TabIndex        =   4
         Top             =   1560
         Value           =   2  'Grayed
         Width           =   3255
      End
      Begin VB.TextBox txtPhosphorescent 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3600
         TabIndex        =   17
         Text            =   "0"
         Top             =   1130
         Width           =   375
      End
      Begin VB.VScrollBar scrollPhosphorescent 
         Height          =   270
         Left            =   3960
         Max             =   0
         Min             =   100
         TabIndex        =   3
         Top             =   1130
         Width           =   150
      End
      Begin VB.CheckBox chkPhosphorescent 
         Caption         =   "Habilitar efeito fosflorecente"
         Height          =   195
         Left            =   185
         TabIndex        =   2
         Top             =   1130
         Value           =   2  'Grayed
         Width           =   3255
      End
      Begin VB.CheckBox chkLowResolutions 
         Caption         =   "Habilitar resoluções baixas"
         Height          =   195
         Left            =   185
         TabIndex        =   1
         Top             =   720
         Value           =   2  'Grayed
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   4095
         Begin VB.TextBox txtMaxScanlines 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3330
            TabIndex        =   5
            Text            =   "1"
            Top             =   140
            Width           =   375
         End
         Begin VB.Label lblMaxScanlines 
            AutoSize        =   -1  'True
            Caption         =   "Número máximo de linhas para renderizar:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   160
            Width           =   3030
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quanto:"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   1080
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboColorPalete_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cboScreenMode_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkFullScreen_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkInterlacedGames_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkLowResolutions_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkPhosphorescent_Click()
    If chkPhosphorescent = 0 Then
        txtPhosphorescent.Enabled = False
        scrollPhosphorescent.Enabled = False
    Else
        txtPhosphorescent.Enabled = True
        scrollPhosphorescent.Enabled = True
    End If
    cmdApply.Enabled = True
End Sub

Private Sub chkRunMonitorSpeed_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkShowFPS_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkSimulateLostPal_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    GravaDados
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    GravaDados
    Unload Me
End Sub

Private Sub Form_Load()
    frameVideoCFG.Caption = frmPri.String19
    chkInterlacedGames.Caption = frmPri.String20
    chkLowResolutions.Caption = frmPri.String21
    chkPhosphorescent.Caption = frmPri.String22
    chkShowFPS.Caption = frmPri.String23
    lblMaxScanlines.Caption = frmPri.String24
    chkSimulateLostPal.Caption = frmPri.String25
    chkRunMonitorSpeed.Caption = frmPri.String26
    chkFullScreen.Caption = frmPri.String27
    lblColorPalete.Caption = frmPri.String28
    lblFPS.Caption = frmPri.String29
    lblScreenMode.Caption = frmPri.String30
    cmdApply.Caption = frmPri.String16
    cmdCancel.Caption = frmPri.String17
    cmdOK.Caption = frmPri.String18
    cboColorPalete.AddItem "PAL"
    cboColorPalete.AddItem "NTSC"
    cboColorPalete.AddItem "SECAM"
    cboColorPalete.ListIndex = 0
    
    cboScreenMode.AddItem "400x300"
    cboScreenMode.AddItem "320x240"
    cboScreenMode.AddItem "320x200"
    cboScreenMode.AddItem "800x600 scanline/interlaced"
    cboScreenMode.AddItem "640x480 scanline/interlaced"
    cboScreenMode.AddItem "640x400 scanline/interlaced"
    cboScreenMode.AddItem "800x600 double scanline"
    cboScreenMode.AddItem "640x480 double scanline"
    cboScreenMode.AddItem "640x400 double scanline"
    cboScreenMode.ListIndex = 0
    
    chkInterlacedGames = frmPri.iniInterlacedGames
    chkLowResolutions = frmPri.iniLowResolutions
    chkPhosphorescent = frmPri.iniPhosphorescentEffect
    scrollPhosphorescent.Value = frmPri.iniPhosphorescentEffectCount
    chkShowFPS = frmPri.iniShowFPS
    txtMaxScanlines = frmPri.iniScanlines
    chkSimulateLostPal = frmPri.iniSimulateColorLossPal
    chkRunMonitorSpeed = frmPri.iniVsync
    chkFullScreen = frmPri.iniFullScreen
    cboColorPalete.ListIndex = frmPri.iniColorPalete
    txtFps = frmPri.iniFPS
    cboScreenMode.ListIndex = frmPri.iniMode
    cmdApply.Enabled = False
End Sub

Private Sub scrollPhosphorescent_Change()
    txtPhosphorescent.Text = scrollPhosphorescent.Value
End Sub

Private Sub txtFps_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtMaxScanlines_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPhosphorescent_Change()
    If txtPhosphorescent < 0 And txtPhosphorescent > 100 Then
        MsgBox String59, vbCritical, String53
        txtPhosphorescent.SetFocus
    End If
    cmdApply.Enabled = True
End Sub

Function GravaDados()
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "InterlacedGames", "S", chkInterlacedGames)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "LowResolutions", "S", chkLowResolutions)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "PhosphorescentEffect", "S", chkPhosphorescent)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "PhosphorescentEffectCount", "S", txtPhosphorescent)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "ShowFPS", "S", chkShowFPS)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "Scanlines", "S", txtMaxScanlines)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "SimulateColorLossPal", "S", chkSimulateLostPal)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "Vsync", "S", chkRunMonitorSpeed)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "FullScreen", "S", chkFullScreen)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "ColorPalete", "S", cboColorPalete.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "FPS", "S", txtFps)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Video", "Mode", "S", cboScreenMode.ListIndex)
    frmPri.iniInterlacedGames = chkInterlacedGames
    frmPri.iniLowResolutions = chkLowResolutions
    frmPri.iniPhosphorescentEffect = chkPhosphorescent
    frmPri.iniPhosphorescentEffectCount = scrollPhosphorescent.Value
    frmPri.iniShowFPS = chkShowFPS
    frmPri.iniScanlines = txtMaxScanlines
    frmPri.iniSimulateColorLossPal = chkSimulateLostPal
    frmPri.iniVsync = chkRunMonitorSpeed
    frmPri.iniFullScreen = chkFullScreen
    frmPri.iniColorPalete = cboColorPalete.ListIndex
    frmPri.iniFPS = txtFps
    frmPri.iniMode = cboScreenMode.ListIndex
    cmdApply.Enabled = False
End Function
