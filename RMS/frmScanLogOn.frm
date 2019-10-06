VERSION 5.00
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#3.1#0"; "pvprgbar.ocx"
Begin VB.Form frmScanLogOn 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Scanning"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   ControlBox      =   0   'False
   Icon            =   "frmScanLogOn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin PVProgressBarLib.PVProgressBar prbScan 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   4695
      _Version        =   196609
      _ExtentX        =   8281
      _ExtentY        =   873
      _StockProps     =   237
      FillColor       =   32768
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   4200
      Top             =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "       Scanning logged on users.                              Please wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmScanLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 1.94
End Sub

Private Sub Timer1_Timer()
'Text1.text = Time$
prbScan.Value = prbScan.Value + 1
If prbScan.Value = 100 Then Unload Me
End Sub
