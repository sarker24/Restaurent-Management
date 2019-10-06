VERSION 5.00
Object = "{B283E209-2CB3-11D0-ADA6-00400520799C}#3.1#0"; "pvprgbar.ocx"
Begin VB.Form frmBackUp 
   BackColor       =   &H00C0B4A9&
   Caption         =   "BackUp Information"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmBackUP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin PVProgressBarLib.PVProgressBar ProgressBar 
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
      _Version        =   196609
      _ExtentX        =   5318
      _ExtentY        =   1296
      _StockProps     =   237
      BackColor       =   16777215
      BackColor       =   16777215
      FillColor       =   32768
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   735
      Left            =   3720
      Picture         =   "frmBackUP.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdBackup 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Backup"
      Height          =   735
      Left            =   3000
      Picture         =   "frmBackUP.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.Frame sdf 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Backup Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4455
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   3120
         Top             =   0
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0B4A9&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   1155
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            BackColor       =   &H00C0B4A9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   120
            MaxLength       =   20
            TabIndex        =   5
            Text            =   "LOTUS"
            Top             =   240
            Width           =   4035
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0B4A9&
            Caption         =   "Backup File Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   4035
         End
      End
   End
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BACKUP INFORMATION  "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con          As New ADODB.Connection
Dim cmd          As New ADODB.Command
Dim rst          As New ADODB.Recordset
Dim strSQL       As String

Private Sub cmdBackup_Click()
If Trim(txtname) = "" Then
    MsgBox "Enter backup file name.", vbInformation, "Confirmation"
    txtname.SetFocus
    Exit Sub
End If

Screen.MousePointer = vbHourglass
Timer1.Enabled = True

On Error GoTo err_h
con.Open
Set cmd.ActiveConnection = con
Dim strName As String
strName = Trim(txtname) & ".bak"

strSQL = "Declare @BackFolder varchar(50)" _
& " SELECT @BackFolder='E:\" & Trim(strName) & "' From sysfiles where name='master'" _
& " Declare @DBname varchar(50)" _
& " set @DBname='" & SDatabaseName & "'" _
& " BACKUP DATABASE @DBname TO DISK =@BackFolder"


cmd.CommandText = strSQL

cmd.Execute
con.Close
If ProgressBar.Value < 4 Then
    ProgressBar.Value = ProgressBar.Value + 100
    Timer1.Enabled = False
End If
MsgBox "BackUp Succesfully Completed.", vbInformation, "Confirmation"
ProgressBar.Value = 0
Unload Me
Screen.MousePointer = vbDefault
Exit Sub
err_h:
      Screen.MousePointer = vbDefault
      MsgBox Err.Description & Chr(13) & "BackUp Failed.", vbInformation, "Confirmation"
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call Connect
'    con.ConnectionString = "Provider=SQLOLEDB.1;Database=Master;User ID=sa;Data Source=" & sServerName & ";Persist Security Info=False;"
     con.ConnectionString = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=Master;Server=" & sServerName
'     con.ConnectionString = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=" & SDatabaseName & ";Server=" & sServerName
     ProgressBar.Value = 0
     Timer1.Enabled = False
 End Sub
Private Sub txtValid(KeyAscii As Integer)
    If KeyAscii < 27 Then
        Exit Sub
    Else
        If Strings.InStr("`~!@#$%^&*()+=\][}{|:';?></,. ", Strings.Chr(KeyAscii)) <> 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Call txtValid(KeyAscii)
End Sub






