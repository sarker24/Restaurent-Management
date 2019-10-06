VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H00C0B4A9&
   Caption         =   " Find"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Text            =   " "
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Cl&ose"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindPrevious 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Find P&revious"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindLast 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Find &Last"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Find Ne&xt"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindFirst 
      BackColor       =   &H00C0B4A9&
      Caption         =   " F&ind First"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Text            =   " "
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   " "
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Available Value"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Field Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub
