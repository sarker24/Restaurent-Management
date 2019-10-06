VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00C0B4A9&
   Caption         =   " User Information"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "User Information Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin MSComCtl2.DTPicker EDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   2520
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   60882947
         CurrentDate     =   41721
      End
      Begin VB.TextBox txtSNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   8
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox TxtUName 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Q&uit"
         Height          =   735
         Left            =   3360
         Picture         =   "frmUser.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdNew 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&New"
         Height          =   735
         Left            =   480
         Picture         =   "frmUser.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Edit"
         Height          =   735
         Left            =   1440
         Picture         =   "frmUser.frx":171E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdOpen 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Open"
         Height          =   735
         Left            =   4320
         Picture         =   "frmUser.frx":1FE8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   2400
         Picture         =   "frmUser.frx":28B2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label lblSerial 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblPgroup 
         BackColor       =   &H00C0B4A9&
         Caption         =   "EDate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label MSG 
         Alignment       =   2  'Center
         BackColor       =   &H00D0B5A8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   600
         TabIndex        =   9
         Top             =   3960
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsfactory             As ADODB.Recordset
Private strFileName           As String
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim str As String


Private Sub alldisable()
    txtUID.Enabled = False
    TxtUName.Enabled = False
    txtPassword.Enabled = False
    txtSNumber.Enabled = False
    EDate.Enabled = False
End Sub

Private Sub allenable()
    txtUID.Enabled = True
    TxtUName.Enabled = True
    txtPassword.Enabled = True
    EDate.Enabled = True
'    cboEn.Enabled = True
End Sub
'
Public Sub allClear()
    txtUID.text = ""
    TxtUName.text = ""
    txtPassword.text = ""
    EDate.Value = Date
    txtSNumber.text = ""
End Sub

Private Sub cmdCancel_Click()
cmdCancel.Enabled = False
   cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    CmdExit.Enabled = True
    cmdEdit.Enabled = True
    txtSNumber.Enabled = False
    cmdOpen.Enabled = True
   Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub

Private Sub cmdDelete_Click()

On Error GoTo ErrHandler
     Dim idelete As Integer
     idelete = MsgBox("Do you want to delete this record?", vbYesNo)
     If frmLogin.txtUID.text = "Admin" Then
    If idelete = vbYes Then
  
    cn.Execute "Delete From RMSUser Where SerialNo ='" & parseQuotes(txtSNumber) & "'"
            Call allClear
            
'           move Next
    End If
        
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdEdit_Click()
 If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        TxtUName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        CmdExit.Enabled = False
        cmdOpen.Enabled = False
 
 ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                CmdExit.Enabled = True
                cmdOpen.Enabled = True
            Call alldisable
                rsfactory.Requery

                Dim s As String
                s = txtSNumber
                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
'                Call search
'                Call countrysearch
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        CmdExit.Enabled = False
        cmdOpen.Enabled = False
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),0) as SerialNo from RMSUser"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSNumber.text = Val(rs!SerialNo) + 1
            
        Call allenable
        txtUID.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSNumber.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                CmdExit.Enabled = True
                cmdOpen.Enabled = True
            Call alldisable
                s = txtSNumber
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
'
    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    If (txtUID.text = "") Then
       MsgBox "Enter User ID"
       TxtUName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtPassword.text = "") Then
      MsgBox "Enter Passward"
      txtPassword.SetFocus
      IsValidRecord = False
      Exit Function
    End If
    
        
If cmdNew.Caption = "&Save" Then
        If rsfactory.RecordCount > 0 Then
        If rsfactory.State <> 0 Then rsfactory.Close
            rsfactory.Open "select * from RMSUser where upper(UID)='" & Strings.UCase(Strings.Trim(parseQuotes(txtUID))) & "'", cn

             If Not rsfactory.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          TxtUName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
End Function

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then
        
        cn.Execute "INSERT INTO RMSUser(SerialNo,UID,UName,UPassword, " & _
                   " EDate) " & _
                   " VALUES ('" & parseQuotes(txtSNumber) & "','" & parseQuotes(txtUID) & "', " & _
                   " '" & parseQuotes(TxtUName) & "', " & _
                   " '" & parseQuotes(txtPassword) & "', " & _
                   " '" & parseQuotes(EDate) & "') "
                   


          rcupdate = True
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    Else

        cn.Execute "Update RMSUser Set UID='" & parseQuotes(txtUID) & _
                  "',UName='" & parseQuotes(TxtUName) & "', " & _
                  " UPassword='" & parseQuotes(txtPassword) & _
                  "',EDate='" & parseQuotes(EDate) & "' " & _
                  " Where SerialNo ='" & parseQuotes(txtSNumber) & "' "
                         

        rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
'    Exit Sub
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rsFactory.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate User ID"
            txtUID = ""
            txtUID.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function


Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtSNumber = rsfactory("SerialNo")
        txtUID = rsfactory("UID")
        TxtUName = rsfactory("UName")
        txtPassword = rsfactory("UPassword")
        EDate = rsfactory("EDate")
       
    End If
End Sub

Private Sub cmdOpen_Click()
frmUserSearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Public Sub PopulateCnf(StrID As String)

    rsfactory.MoveFirst
    rsfactory.Find "SerialNo=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

Private Sub Form_Load()
Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from RMSUser", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
   
    If Not rsfactory.EOF Then FindRecord
    
'    cboEn.AddItem "ADMIN"
'    cboEn.AddItem "COUNTER"
'    cboEn.AddItem "POWER USER"
    txtSNumber.Enabled = False
End Sub




