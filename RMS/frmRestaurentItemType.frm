VERSION 5.00
Begin VB.Form frmRestaurentItemCatagory 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Restaurent Item Catagory"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   Icon            =   "frmRestaurentItemType.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Quit"
      Height          =   615
      Index           =   2
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Edit"
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Pre&veiw"
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Print"
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Find"
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete "
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &New"
      Height          =   615
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtDescriptions 
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Text            =   " "
      Top             =   720
      Width           =   8535
   End
   Begin VB.TextBox txtItemCatagory 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Descriptions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   1455
   End
   Begin VB.Label lblType 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Item Catagory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRestaurentItemCatagory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private rs As ADODB.Recordset
'
'
'
'
'Private Sub cmdFind_Click()
'frmFind.Show vbModal
'End Sub
'
'Private Sub cmdnew_Click(Index As Integer)
'Set rs = New ADODB.Recordset
'If cmdNew.Caption = "New" Then
'cmdNew.Caption = "Save"
'cmd
'
'End Sub
'
'Private Sub cmdQuit_Click(Index As Integer)
'Unload Me
'End Sub

