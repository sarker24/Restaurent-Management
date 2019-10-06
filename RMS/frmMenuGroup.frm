VERSION 5.00
Begin VB.Form frmMenuGroup 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Menugroup Setup"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmMenuGroup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Menu Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   5385
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   465
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtMGID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0B4A9&
         Caption         =   "Menu Group Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label lbiCompanyID 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Group ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   4560
      Picture         =   "frmMenuGroup.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "C&lose"
      Height          =   795
      Left            =   3480
      Picture         =   "frmMenuGroup.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1065
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1320
      Picture         =   "frmMenuGroup.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1065
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   240
      Picture         =   "frmMenuGroup.frx":1FE8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   795
      Left            =   2400
      Picture         =   "frmMenuGroup.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MENU GROUP  SETUP  "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   5865
   End
End
Attribute VB_Name = "frmMenuGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private rs                     As ADODB.Recordset
Private rsMenuGroup              As ADODB.Recordset
'Private strStream             As ADODB.Stream
Private strFileName            As String
Private bRecordExists          As Boolean
Dim str                        As String
'Private rm                    As New ADODB.Recordset
'Private rc                    As New ADODB.Recordset

Private Sub cmdCancel_Click()

   CmdCancel.Enabled = False
   cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    txtMGID.Enabled = False
    Call allClear
'    txtCompanyID.Enabled = False
    Call alldisable
    If Not rsMenuGroup.EOF Then FindRecord
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
       Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        CmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
'        txtMGID.Enabled = True
        Call allClear
'       ModFunction.TextEnable Me, True
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(MG_ID),0) as SerialNo from MenuGroup"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtMGID.text = Val(rs!SerialNo) + 1
           Call allenable
'           Call alldisable
        txtname.SetFocus

    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then

                txtMGID.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                CmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
'                ModFunction.TextEnable Me, False
                Call alldisable
                s = txtname
                rsMenuGroup.Requery
                rsMenuGroup.MoveFirst
                rsMenuGroup.Find "MGName='" & parseQuotes(s) & "'"
               
                FindRecord
            End If
        End If
    End If

    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtname.SetFocus
        cmdEdit.Caption = "&Update"
        CmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False

    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                CmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                rsMenuGroup.Requery

                Dim s As String
                s = txtname
                rsMenuGroup.Find "MGName='" & parseQuotes(s) & "'"
                
                FindRecord
            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
   strCallingForm = LCase("frmMenuGroup")
    frmMenuGroupSearch.Show vbModal
    cmdOpen.Enabled = True
    CmdCancel.Enabled = True
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
    ModFunction.StartUpPosition Me
    Set rsMenuGroup = New ADODB.Recordset
'    Set rsImage = New ADODB.Recordset
    rsMenuGroup.Open "select  DISTINCT * from MenuGroup", cn, adOpenStatic, adLockReadOnly
    
ModFunction.TextEnable Me, False
    
    Call alldisable

   If rsMenuGroup.RecordCount > 0 Then

        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
   If Not rsMenuGroup.EOF Then FindRecord
    
    txtMGID.Enabled = False
    
End Sub

Private Sub allClear()
'    ModFunction.TextClear Me
txtname.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO MenuGroup(MG_ID,MGName) " & _
                   " VALUES ('" & parseQuotes(txtMGID) & "','" & parseQuotes(txtname) & "')"
                   

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else

        cn.Execute "Update MenuGroup Set MGName='" & parseQuotes(txtname) & _
                  "'WHERE  MG_ID ='" & parseQuotes(txtMGID) & "' "
            
                  
                 
'        If (UCase(txtMGID.text) = UCase(rsMenuGroup!MG_ID)) And UCase(txtname.text) = UCase(rsMenuGroup!MenuGroupName) Then
'    MsgBox "Trying Duplicate MenuGroup Name"
'        Exit Function
'    End If
'
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rsMenuGroup.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate MenuGroup Name"
            txtname = ""
            txtname.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsMenuGroup.EOF Then
        txtMGID = rsMenuGroup("MG_ID")
        txtname = rsMenuGroup("MGName")
        
   End If
End Sub


Private Sub allenable()
    txtname.Enabled = True
    
End Sub

Private Sub alldisable()
    txtname.Enabled = False
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (txtname.text = "") Then
       MsgBox "Enter MenuGroup Name"
       txtname.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtMGID.text = "") Then
     MsgBox "Enter MG_ID"
     txtMGID.SetFocus
     IsValidRecord = False
     Exit Function
     
    End If
    
    If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
        If rsMenuGroup.RecordCount > 0 Then
        If rsMenuGroup.State <> 0 Then rsMenuGroup.Close
            rsMenuGroup.Open "select * from MenuGroup where upper(MGName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtname))) & "'", cn

             If Not rsMenuGroup.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtname.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
    
'    If cmdEdit.Caption <> "&Update" Then
'        If rsMenuGroup.RecordCount > 0 Then
'            If (UCase(txtname.text) = UCase(rsMenuGroup!MenuGroupName)) Then
'                  MsgBox "Trying Duplicate MenuGroup Name"
'                  IsValidRecord = False
'                 Exit Function
'            End If
'         End If
'    End If


End Function

Public Sub PopulateMenuGroup(StrID As String)


    rsMenuGroup.MoveFirst
    rsMenuGroup.Find "MG_ID=" & parseQuotes(StrID)
    If rsMenuGroup.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub









