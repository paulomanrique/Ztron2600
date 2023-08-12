VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diretórios"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4995
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
   ScaleHeight     =   2160
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Aplicar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame frameRomDir 
      Caption         =   "Diretório das Roms"
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   4935
      Begin VB.TextBox txtRomDir 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   3735
      End
      Begin VB.CommandButton cmdProcuraRoms 
         Caption         =   "Procurar..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame frameZ26Dir 
      Caption         =   "Diretório do Z26"
      Height          =   735
      Left            =   15
      TabIndex        =   7
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdProcuraZ26 
         Caption         =   "Procurar..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   1
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox txtZ26Dir 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim fso As New FileSystemObject
Dim fld As Folder
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

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type


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

Private Sub cmdProcuraRoms_Click()
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo

   szTitle = frmPri.String54
   With tBrowseInfo
      .hWndOwner = Me.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With

   lpIDList = SHBrowseForFolder(tBrowseInfo)

   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   End If
        txtRomDir = sBuffer
End Sub

Private Sub cmdProcuraZ26_Click()
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo

   szTitle = frmPri.String55
   With tBrowseInfo
      .hWndOwner = Me.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With

   lpIDList = SHBrowseForFolder(tBrowseInfo)

   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   End If
    If fso.FileExists(sBuffer & "\Z26.EXE") = False Then
        MsgBox frmPri.String57, vbCritical, frmPri.String56
        txtZ26Dir = sBuffer
    Else
        txtZ26Dir = sBuffer
    End If
End Sub

Private Sub Form_Load()
    frameRomDir = frmPri.String13
    frameZ26Dir = frmPri.String14
    cmdProcuraRoms.Caption = frmPri.String15
    cmdProcuraZ26.Caption = frmPri.String15
    cmdApply.Caption = frmPri.String16
    cmdCancel.Caption = frmPri.String17
    cmdOK.Caption = frmPri.String18
    
    txtRomDir = frmPri.iniRomDir
    txtZ26Dir = frmPri.iniZ26Dir
    cmdApply.Enabled = False
End Sub

Private Sub txtRomDir_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtZ26Dir_Change()
    cmdApply.Enabled = True
End Sub
Function GravaDados()
    Call fWriteValue(App.Path & "\ztron2600.ini", "General", "Z26Dir", "S", txtZ26Dir)
    Call fWriteValue(App.Path & "\ztron2600.ini", "General", "RomDir", "S", txtRomDir)
    frmPri.iniRomDir = txtRomDir
    frmPri.iniZ26Dir = txtZ26Dir
    cmdApply.Enabled = False
End Function
