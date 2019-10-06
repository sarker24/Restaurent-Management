VERSION 5.00
Begin VB.Form Frm_Splash 
   BackColor       =   &H001FD383&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3210
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Frm_Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -600
      Top             =   2760
   End
   Begin VB.Frame Fra_Splash 
      BackColor       =   &H00E0E0E0&
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   7560
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   650
         Left            =   -960
         Top             =   2400
      End
      Begin VB.Label LblTradeMark 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TradeMark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6615
         TabIndex        =   10
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label LblBuild 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Build:001"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label LblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2025
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   6720
         Picture         =   "Frm_Splash.frx":09EA
         Top             =   600
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -120
         Picture         =   "Frm_Splash.frx":18B4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7695
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   -120
         Picture         =   "Frm_Splash.frx":1E536
         Stretch         =   -1  'True
         Top             =   2805
         Width           =   7755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1920
         TabIndex        =   8
         Top             =   3000
         Width           =   1380
      End
      Begin VB.Label LblProductCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code:"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   3030
         Width           =   1755
      End
      Begin VB.Label LblCompanyName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2835
      End
      Begin VB.Label LblCopyright 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6615
         TabIndex        =   3
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8220
         TabIndex        =   2
         Top             =   4560
         Width           =   45
      End
      Begin VB.Label LblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6330
         TabIndex        =   4
         Top             =   720
         Width           =   330
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licensed To: LOTUS ETANG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   5670
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000013&
         BorderColor     =   &H00808080&
         FillColor       =   &H001FD383&
         FillStyle       =   0  'Solid
         Height          =   2280
         Left            =   30
         Top             =   600
         Width           =   7470
      End
   End
End
Attribute VB_Name = "Frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim FadeIn As Boolean
    
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Initialize()
    
    Transparency Me, 0
    FadeIn = True
        
    LblCompanyName.Caption = App.CompanyName
    LblProductName.Caption = App.FileDescription
    LblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    LblBuild.Caption = "Build " & VBA.Format$(App.Revision, "0000")
    LblCopyright.Caption = App.LegalCopyright
    LblTradeMark.Caption = App.LegalTrademarks
    
'    LblLicenseTo.Caption = "Licensed To: " & App.CompanyName '& Licence.Company.Name
    LblProductCode.Caption = VBA.Replace(VBA.GetSetting("Microsoft", App.Title, "SC", "0"), "|", "")
    
End Sub

Private Sub Form_Load()
    VBA.Randomize
End Sub

Private Sub Fra_Splash_Click()
On Error Resume Next
    Load frmLogin: frmLogin.Show
    Timer1.Enabled = False: Timer2.Enabled = False
    Unload Me
End Sub

Private Sub LblCompanyName_Click()
    Call Fra_Splash_Click
End Sub

Private Sub lblProductName_Click()
    Call Fra_Splash_Click
End Sub

Private Sub Timer1_Timer()
On Error GoTo Err
    
    Static Index%
    
    Index = VBA.IIf(FadeIn = True, Index + 5, Index - 5)
    Transparency Me, Index
    Exit Sub
    
Err:
    
    If FadeIn = False Then Call Fra_Splash_Click: Exit Sub
    FadeIn = Not FadeIn: Timer1.Enabled = False: Timer2.Enabled = True
    
End Sub

Private Sub Timer2_Timer()
    
    Static vCounter&
    
    iArrayList = VBA.Split(" ", " ")
    
    If vCounter >= UBound(iArrayList) + 1 Then Timer2.Enabled = False: Timer1.Enabled = True: Exit Sub
    
    Sleep VBA.Int((1200 * VBA.Rnd) + 400)
    VBA.DoEvents
    vCounter = vCounter + 1
    
End Sub
