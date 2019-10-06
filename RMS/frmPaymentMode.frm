VERSION 5.00
Begin VB.Form frmPaymentMode 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Payment Mode"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   Icon            =   "frmPaymentMode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &New"
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmPaymentMode.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete "
      Height          =   615
      Left            =   2040
      Picture         =   "frmPaymentMode.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Find"
      Height          =   615
      Left            =   3000
      Picture         =   "frmPaymentMode.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Edit"
      Height          =   615
      Left            =   1080
      Picture         =   "frmPaymentMode.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Quit"
      Height          =   615
      Index           =   2
      Left            =   3960
      Picture         =   "frmPaymentMode.frx":1EF2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtPaymentMode 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   " "
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtModeID 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblPaymentMode 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Payment Mode Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lbleModeIDI 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Payment Mode ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmPaymentMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFind_Click()
frmFind.Show vbModal
End Sub


Private Sub cmdQuit_Click(Index As Integer)
Unload Me
End Sub
