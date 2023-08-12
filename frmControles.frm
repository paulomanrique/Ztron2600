VERSION 5.00
Begin VB.Form frmControles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controles"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4440
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
   ScaleHeight     =   5880
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   18
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame frameControlerCFG 
      Caption         =   "Configuração de controles"
      Height          =   5295
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   4335
      Begin VB.CheckBox chkAllowAllDirections 
         Caption         =   "Permitir que as 4 direções do joystick sejam precionadas simultaneamente"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Value           =   2  'Grayed
         Width           =   3855
      End
      Begin VB.CheckBox chkEnableMindlink 
         Caption         =   "Habilitar controlador Mindlink"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Value           =   2  'Grayed
         Width           =   2415
      End
      Begin VB.ComboBox cboMindLink 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1155
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   4095
         Begin VB.TextBox txtAjustLightgunScanlines 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            TabIndex        =   15
            Text            =   "1"
            Top             =   840
            Width           =   375
         End
         Begin VB.CheckBox chkAjustLightgunScanlines 
            Caption         =   "Ajustar por scanlines"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Value           =   2  'Grayed
            Width           =   2535
         End
         Begin VB.TextBox txtEnableLightgunCycles 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            TabIndex        =   13
            Text            =   "1"
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox chkEnableLightgun 
            Caption         =   "Emular a pistola com o mouse"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   2  'Grayed
            Width           =   2415
         End
         Begin VB.Label lblAjustLightgunScanlines 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Quantas:"
            Height          =   195
            Left            =   2700
            TabIndex        =   29
            Top             =   840
            Width           =   675
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000D&
            BorderStyle     =   6  'Inside Solid
            X1              =   0
            X2              =   3870
            Y1              =   670
            Y2              =   670
         End
         Begin VB.Label lblEnableLightgunCycles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ciclos:"
            Height          =   195
            Left            =   2880
            TabIndex        =   28
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   4095
         Begin VB.CheckBox chkTwoMousePaddle 
            Caption         =   "Emular dois paddles com o mouse"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Value           =   2  'Grayed
            Width           =   2775
         End
         Begin VB.CheckBox chkMousePaddleEnable 
            Caption         =   "Emular o mouse como paddle"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   2  'Grayed
            Width           =   2415
         End
         Begin VB.ComboBox cboYMousePaddle 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1080
            Width           =   510
         End
         Begin VB.ComboBox cboXMousePaddle 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1080
            Width           =   510
         End
         Begin VB.ComboBox cboMousePaddleEnable 
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   510
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000D&
            BorderStyle     =   6  'Inside Solid
            X1              =   120
            X2              =   3990
            Y1              =   630
            Y2              =   630
         End
         Begin VB.Label lblTwoMousePaddle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Quais:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label lblMousePaddleEnable 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qual:"
            Height          =   195
            Left            =   3000
            TabIndex        =   25
            Top             =   285
            Width           =   390
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   4095
         Begin VB.ComboBox cboKeyboardPaddle 
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   510
         End
         Begin VB.VScrollBar scrollKeyboardPaddleSensitivity 
            Height          =   255
            Left            =   1560
            Max             =   0
            Min             =   15
            TabIndex        =   10
            Top             =   495
            Width           =   150
         End
         Begin VB.TextBox txtKeyboardPaddleSensibilty 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   23
            Text            =   "1"
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox chkKeyboardPaddleEnable 
            Caption         =   "Emular o teclado como paddle"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Alguns jogos como o Raiders of the Lost Ark usam o joystick de maneira errada. Use essa opção para inverter o joystick"
            Top             =   150
            Value           =   2  'Grayed
            Width           =   2535
         End
         Begin VB.Label lblKeyboardPaddleSensibility 
            AutoSize        =   -1  'True
            Caption         =   "Sensibilidade:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   525
            Width           =   975
         End
         Begin VB.Label lblKeyboardPaddle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qual:"
            Height          =   195
            Left            =   2880
            TabIndex        =   21
            Top             =   280
            Width           =   390
         End
      End
      Begin VB.CheckBox chkReverseJoystick 
         Caption         =   "Inverter joystick"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Alguns jogos como o Raiders of the Lost Ark usam o joystick de maneira errada. Use essa opção para inverter o joystick"
         Top             =   360
         Value           =   2  'Grayed
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAjustLightgunScanlines_Click()
    If chkAjustLightgunScanlines = 0 Then
        txtAjustLightgunScanlines.Enabled = False
    Else
        txtAjustLightgunScanlines.Enabled = True
    End If
End Sub

Private Sub chkEnableLightgun_Click()
    If chkEnableLightgun = 0 Then
        txtEnableLightgunCycles.Enabled = False
        chkAjustLightgunScanlines.Enabled = False
        txtAjustLightgunScanlines.Enabled = False
    Else
        txtEnableLightgunCycles.Enabled = True
        chkAjustLightgunScanlines.Enabled = True
        txtAjustLightgunScanlines.Enabled = True
    End If
End Sub

Private Sub chkEnableMindlink_Click()
    If chkEnableMindlink = 0 Then
        cboMindLink.Enabled = False
    Else
        cboMindLink.Enabled = True
    End If
End Sub

Private Sub chkKeyboardPaddleEnable_Click()
    If chkKeyboardPaddleEnable = 0 Then
        txtKeyboardPaddleSensibilty.Enabled = False
        scrollKeyboardPaddleSensitivity.Enabled = False
        cboKeyboardPaddle.Enabled = False
    Else
        txtKeyboardPaddleSensibilty.Enabled = True
        scrollKeyboardPaddleSensitivity.Enabled = True
        cboKeyboardPaddle.Enabled = True
    End If
    cmdApply.Enabled = True
End Sub

Private Sub chkMousePaddleEnable_Click()
    If chkMousePaddleEnable = 0 Then
        cboMousePaddleEnable.Enabled = False
    Else
        cboMousePaddleEnable.Enabled = True
        chkTwoMousePaddle = 0
    End If
    cmdApply.Enabled = True
End Sub

Private Sub chkReverseJoystick_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkTwoMousePaddle_Click()
    If chkTwoMousePaddle = 0 Then
        cboXMousePaddle.Enabled = False
        cboYMousePaddle.Enabled = False
    Else
        cboXMousePaddle.Enabled = True
        cboYMousePaddle.Enabled = True
        chkMousePaddleEnable = 0
    End If
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
    frameControlerCFG.Caption = frmPri.String31
    chkReverseJoystick.Caption = frmPri.String32
    chkAllowAllDirections.Caption = frmPri.String33
    chkEnableMindlink.Caption = frmPri.String34
    chkMousePaddleEnable.Caption = frmPri.String37
    lblMousePaddleEnable.Caption = frmPri.String38
    chkTwoMousePaddle.Caption = frmPri.String39
    lblTwoMousePaddle.Caption = frmPri.String40
    chkKeyboardPaddleEnable.Caption = frmPri.String41
    lblKeyboardPaddleSensibility.Caption = frmPri.String42
    lblKeyboardPaddle.Caption = frmPri.String38
    chkEnableLightgun.Caption = frmPri.String43
    lblEnableLightgunCycles.Caption = frmPri.String44
    chkAjustLightgunScanlines.Caption = frmPri.String45
    lblAjustLightgunScanlines.Caption = frmPri.String46
    cmdApply.Caption = frmPri.String16
    cmdCancel.Caption = frmPri.String17
    cmdOK.Caption = frmPri.String18
    cboMindLink.Clear
    cboMindLink.AddItem frmPri.String35
    cboMindLink.AddItem frmPri.String36
    cboMindLink.ListIndex = frmPri.iniMindlinkSide

    
    cboMousePaddleEnable.AddItem "0"
    cboMousePaddleEnable.AddItem "1"
    cboMousePaddleEnable.AddItem "2"
    cboMousePaddleEnable.AddItem "3"
    cboMousePaddleEnable.ListIndex = 0
    
    cboXMousePaddle.AddItem "0"
    cboXMousePaddle.AddItem "1"
    cboXMousePaddle.AddItem "2"
    cboXMousePaddle.AddItem "3"
    cboXMousePaddle.ListIndex = 0
    
    cboYMousePaddle.AddItem "0"
    cboYMousePaddle.AddItem "1"
    cboYMousePaddle.AddItem "2"
    cboYMousePaddle.AddItem "3"
    cboYMousePaddle.ListIndex = 0
    
    cboKeyboardPaddle.AddItem "0"
    cboKeyboardPaddle.AddItem "1"
    cboKeyboardPaddle.AddItem "2"
    cboKeyboardPaddle.AddItem "3"
    cboKeyboardPaddle.ListIndex = 0

    chkReverseJoystick = frmPri.iniReverseJoystick
    chkAllowAllDirections = frmPri.iniAllowAllDirections
    chkEnableMindlink = frmPri.iniEnableMindlink
    chkMousePaddleEnable = frmPri.iniMousePaddleEnable
    cboMousePaddleEnable.ListIndex = frmPri.iniMousePaddleId
    chkTwoMousePaddle = frmPri.iniMouseTwoPaddleEnable
    cboXMousePaddle.ListIndex = frmPri.iniMouseXPaddle
    cboYMousePaddle.ListIndex = frmPri.iniMouseYPaddle
    chkKeyboardPaddleEnable = frmPri.iniKeyboardPaddleEnable
    cboKeyboardPaddle.ListIndex = frmPri.iniKeyboardPaddleId
    scrollKeyboardPaddleSensitivity.Value = frmPri.iniKeyboardPaddleSensitivity
    chkEnableLightgun = frmPri.iniLightgunEnable
    txtEnableLightgunCycles = frmPri.iniLightgunCycles
    chkAjustLightgunScanlines = frmPri.iniLightgunAdjustByScanlines
    txtAjustLightgunScanlines = frmPri.iniLightgunAdjustByScanlinesMuch
End Sub

Private Sub txtMousePaddle_Change()
    cmdApply.Enabled = True
End Sub

Function GravaDados()
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "ReverseJoystick", "S", chkReverseJoystick)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "AllowAllDirections", "S", chkAllowAllDirections)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "EnableMindlink", "S", chkEnableMindlink)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "MindlinkSide", "S", cboMindLink.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "MousePaddleEnable", "S", chkMousePaddleEnable)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "MousePaddleId", "S", cboMousePaddleEnable.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "MouseTwoPaddleEnable", "S", chkTwoMousePaddle)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "MouseXPaddle", "S", cboXMousePaddle.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "MouseYPaddle", "S", cboYMousePaddle.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "KeyboardPaddleEnable", "S", chkKeyboardPaddleEnable)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "KeyboardPaddleId", "S", cboKeyboardPaddle.ListIndex)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "KeyboardPaddleSensitivity", "S", txtKeyboardPaddleSensibilty)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunEnable", "S", chkEnableLightgun)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunCycles", "S", txtEnableLightgunCycles)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunAdjustByScanlines", "S", chkAjustLightgunScanlines)
    Call fWriteValue(App.Path & "\ztron2600.ini", "Controlers", "LightgunAdjustByScanlinesMuch", "S", txtAjustLightgunScanlines)
    frmPri.iniReverseJoystick = chkReverseJoystick
    frmPri.iniAllowAllDirections = chkAllowAllDirections
    frmPri.iniEnableMindlink = chkEnableMindlink
    frmPri.iniMindlinkSide = cboMindLink.ListIndex
    frmPri.iniMousePaddleEnable = chkMousePaddleEnable
    frmPri.iniMousePaddleId = cboMousePaddleEnable.ListIndex
    frmPri.iniMouseTwoPaddleEnable = chkTwoMousePaddle
    frmPri.iniMouseXPaddle = cboXMousePaddle.ListIndex
    frmPri.iniMouseYPaddle = cboYMousePaddle.ListIndex
    frmPri.iniKeyboardPaddleEnable = chkKeyboardPaddleEnable
    frmPri.iniKeyboardPaddleId = cboKeyboardPaddle.ListIndex
    frmPri.iniKeyboardPaddleSensitivity = scrollKeyboardPaddleSensitivity.Value
    frmPri.iniLightgunEnable = chkEnableLightgun
    frmPri.iniLightgunCycles = txtEnableLightgunCycles
    frmPri.iniLightgunAdjustByScanlines = chkAjustLightgunScanlines
    frmPri.iniLightgunAdjustByScanlinesMuch = txtAjustLightgunScanlines
    cmdApply.Enabled = False
End Function

Private Sub scrollKeyboardPaddleSensitivity_Change()
    txtKeyboardPaddleSensibilty = scrollKeyboardPaddleSensitivity
End Sub

Private Sub txtKeyboardPaddleSensibilty_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKey0 And KeyAscii <> vbKey1 And KeyAscii <> vbKey2 And KeyAscii <> vbKey3 And KeyAscii <> vbKey4 And KeyAscii <> vbKey5 And KeyAscii <> vbKey6 And KeyAscii <> vbKey7 And KeyAscii <> vbKey8 And KeyAscii <> vbKey9 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKeyboardPaddleSensibilty_LostFocus()
    If txtKeyboardPaddleSensibilty < 1 Or txtKeyboardPaddleSensibilty > 15 Then
        MsgBox frmPri.String58, vbCritical, String53
        txtKeyboardPaddleSensibilty.SetFocus
    End If
End Sub
