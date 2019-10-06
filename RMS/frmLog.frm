VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   0  'None
   Caption         =   "Login To application ..."
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLog.frx":058A
   ScaleHeight     =   2820
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTime 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3840
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   120
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   2880
      Picture         =   "frmLog.frx":3320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton CmdEnter 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   1920
      Picture         =   "frmLog.frx":3BEA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "1991"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtUID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "debdas"
      Top             =   960
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker ExpiryDate 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   50528257
      CurrentDate     =   41037
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Tag             =   "&User Name:"
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Tag             =   "&Password:"
      Top             =   1440
      Width           =   960
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub CmdEnter_Click()
    Call Connect
    Dim str As String
    Dim cm As New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set cm = New ADODB.Connection '
    Set cn = New ADODB.Connection
'   str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & SDatabaseName & ";Data Source=" & sServerName
    str = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=" & SDatabaseName & ";Server=" & sServerName
    cn.Open str
    txtDate.text = Date        'Format((cm.Execute("Select GetDate()")), "dd-MM-yyyy")
    txtTime.text = Time        'Format((cm.Execute("Select GetDate()")), "hh:mm:ss")
    str = "select UID,UPassword,Upper(UID)as Name  from RMSUser where UID ='" & txtUID.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
    If rs!UPassword = Trim(CStr(txtPassword.text)) Then
    frmScanLogOn.Show vbModal
        Call frmMain.Show
'        If rs!Name = "ADMIN" And rs!UPassword = "VB5DELPHI2" Then
            If rs!Name = "ADMIN" Then
         frmLogin.Hide
             frmMain.Enabled = True
             frmMain.mnuCashMemo.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuCommunication.Enabled = True
             frmMain.mnuExit.Enabled = True
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuReport.Enabled = True
             frmMain.mnuCalculator.Enabled = True
             frmMain.mnuCashMemoModify.Enabled = True
             frmCashMemo.cmdEdit.Enabled = True
             frmCashMemo.cmdChange.Enabled = True
             frmCashMemo.cboMode.Enabled = True
             frmMain.mnuRGuest.Enabled = True
             
         cn.Execute " insert into RMSLogin Values ('" & txtUID.text & "','" & Format(txtDate.text, "yyyy-mm-dd") & "','" & Format(txtTime.text, "HH:MM:SS") & "')"
         
'         cm.Execute str
'         cm.CommitTrans
         
         ElseIf rs!Name = "BORHAN" Then
         frmLogin.Hide
             
             frmMain.Enabled = True
             frmMain.mnuCashMemo.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuCommunication.Enabled = False
             frmMain.mnuExit.Enabled = True
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuReport.Enabled = True
             frmMain.mnuSDStatement.Visible = True
             frmMain.mnuWTStatement.Visible = True
             frmMain.mnuWSStatement.Visible = True
             frmMain.mnuCalculator.Enabled = True
             frmMain.mnuUser.Visible = False
             frmCashMemo.cmdEdit.Enabled = True
             frmCashMemo.cmdChange.Visible = True
             frmMain.mnuRGuest.Visible = True
             
         cn.Execute " insert into RMSLogin Values ('" & txtUID.text & "','" & Format(txtDate.text, "yyyy-mm-dd") & "','" & Format(txtTime.text, "HH:MM:SS") & "')"
             
'            cm.Execute str
'            cm.CommitTrans
'            Unload Me
        Else
        
           frmMain.Enabled = True
             frmMain.mnuCashMemo.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuCommunication.Enabled = False
             frmMain.mnuExit.Enabled = True
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuReport.Visible = False
             frmMain.mnuCalculator.Enabled = True
             frmMain.mnuUser.Visible = False
             frmMain.mnuCashMemoModify.Enabled = False
             frmMain.mnuReport.Visible = True
             frmMain.mnuReport.Visible = True
             frmMain.mnuSDStatement.Visible = True
             frmMain.mnuWTStatement.Visible = True
             frmMain.mnuWSStatement.Visible = True
             frmCashMemo.cmdEdit.Enabled = True
             frmCashMemo.cmdChange.Visible = False
             frmMain.mnuDailySales.Visible = True
             frmMain.mnuSalesSummery.Visible = True
             frmMain.mnuRGuest.Visible = True
             frmMain.mnuWSStatement.Visible = False
'             frmCashMemo.ChkNBR.Visible = False
            
          frmLogin.Hide
          cn.Execute " insert into RMSLogin Values ('" & txtUID.text & "','" & Format(txtDate.text, "yyyy-mm-dd") & "','" & Format(txtTime.text, "HH:MM:SS") & "')"

        End If
    Else
            MsgBox "Invalid Password or User Name. Please try again.", vbInformation, "Confarmation"
            txtPassword.text = ""
            txtPassword.SetFocus
    End If

End Sub

Private Sub CmdCancel_GotFocus()
    cmdCancel.FontBold = True
End Sub

Private Sub CmdCancel_LostFocus()
    cmdCancel.FontBold = False
End Sub



Private Sub TxtUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtPassword.SetFocus
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call CmdEnter_Click
'    CmdEnter.SetFocus
    End If
End Sub

Private Sub Form_Load()
Call TimeExpired

If App.PrevInstance = True Then
MsgBox "Software already running......"
Unload Me
End If
End Sub

Private Sub TimeExpired()
Dim CurrentDate As Date
Dim ExpiryDate
CurrentDate = Now
ExpiryDate = "12-12-2019"
If CurrentDate > ExpiryDate Then
'MsgBox ("Database Connection Fail.")
'MsgBox ("Your system has Expired")
'frmLogIn.Hide
Unload Me
End If
End Sub

Private Sub Timer1_Timer()
    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
    txtCTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub

Private Sub TxtUID_GotFocus()
    txtUID.BackColor = &HFFC0C0
    txtUID.SelStart = 0
    txtUID.SelLength = Len(txtUID)
End Sub

Private Sub TxtUID_LostFocus()
    txtUID.BackColor = &HFFFFFF
    txtUID.text = StrConv(txtUID.text, vbProperCase)
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.BackColor = &HFFC0C0
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword.BackColor = &HFFFFFF
End Sub

