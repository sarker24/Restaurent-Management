VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Informations"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   5775
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   7095
      Begin VB.TextBox txtDiscountAmt 
         Height          =   480
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   17
         Text            =   " "
         Top             =   4575
         Width           =   5175
      End
      Begin VB.TextBox txtDiscountCard 
         Height          =   480
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   16
         Top             =   720
         Width           =   5175
      End
      Begin VB.TextBox txtCName 
         Height          =   480
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1280
         Width           =   5175
      End
      Begin VB.TextBox txtCAddress 
         Height          =   1155
         Left            =   1680
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox txtCPhone 
         Height          =   480
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   13
         Text            =   " "
         Top             =   3000
         Width           =   5175
      End
      Begin VB.TextBox txtCEmail 
         Height          =   480
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   4065
         Width           =   5175
      End
      Begin VB.TextBox txtCID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   465
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Left            =   5760
         Top             =   240
      End
      Begin VB.TextBox txtRPoint 
         Height          =   480
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   9
         Text            =   " "
         Top             =   5130
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker BDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   3600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   66191363
         CurrentDate     =   41350
      End
      Begin VB.Label lblDiscountAmt 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Discount Amount"
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
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   4575
         Width           =   1575
      End
      Begin VB.Label lblDiscountCard 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Discount Card"
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
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblCustomerName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Customer Name"
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
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblCustomerAddress 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Customer Address"
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
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblCustomerID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Customer ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCustomerPhone 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Customer Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblBDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Birth Date"
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
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblCEmail 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Customer E-mail"
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
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   4065
         Width           =   1575
      End
      Begin VB.Label lblRPoint 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Reward Point/Amt"
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
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   5130
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6120
      Picture         =   "frmCustomer.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4200
      Picture         =   "frmCustomer.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   945
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2280
      Picture         =   "frmCustomer.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   945
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   360
      Picture         =   "frmCustomer.frx":1CA8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   945
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1320
      Picture         =   "frmCustomer.frx":2572
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3240
      Picture         =   "frmCustomer.frx":2E3C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   945
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5160
      Picture         =   "frmCustomer.frx":3706
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      Caption         =   " Customer Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmCustomer"
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
'--------------------------------------------------------------
Private oReportApp                        As CRPEAuto.Application
Private oReport                           As CRPEAuto.Report
Private oReportDatabase                   As CRPEAuto.Database
Private oReportDatabaseTables             As CRPEAuto.DatabaseTables
Private oReportDatabaseTable              As CRPEAuto.DatabaseTable
'Private oReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private oReportFF                         As CRPEAuto.FormulaFieldDefinition
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions

Private Sub cmdPreview_Click()
    Call printReport
End Sub

'Private Sub cmdPreview_Click()
'    Call printReport
'End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdNew.Caption = "&New"
    cmdEdit.Caption = "&Edit"
    cmdPreview.Enabled = True
    CmdDelete.Enabled = True
    cmdOpen.Enabled = True
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    txtCID.Enabled = False
    Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub

'
Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdDelete.Enabled = False
        cmdOpen.Enabled = False
        cmdPreview.Enabled = False
'        chkActive.Enabled = False
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(CID),0) as SerialNo from RMSCustomer"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtCID.text = Val(rs!SerialNo) + 1
            
        Call allenable
        txtDiscountCard.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtCID.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
'                chkActive.Enabled = True
                Call alldisable
                s = txtCName
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "CName='" & parseQuotes(s) & "'"
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

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtCName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdDelete.Enabled = False
        cmdPreview.Enabled = False
        cmdOpen.Enabled = False
'        chkActive.Enabled = False
ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
'                chkActive.Enabled = True
        cmdPreview.Enabled = True
        cmdOpen.Enabled = True
                Call alldisable
                rsfactory.Requery

                Dim s As String
                s = txtCName
                rsfactory.Find "CName='" & parseQuotes(s) & "'"
'                Call search
'                Call countrysearch
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    frmCustomerSearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from RMSCustomer", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
   
    If Not rsfactory.EOF Then FindRecord
    
    txtCID.Enabled = False
    
End Sub

Private Sub allenable()
    txtDiscountCard.Enabled = True
    txtCName.Enabled = True
    txtCAddress.Enabled = True
    txtCPhone.Enabled = True
    BDate.Enabled = True
    txtCEmail.Enabled = True
    txtDiscountAmt.Enabled = True
    txtRPoint.Enabled = True
    chkActive.Enabled = True
End Sub

Private Sub alldisable()
    txtDiscountCard.Enabled = False
    txtCName.Enabled = False
    txtCAddress.Enabled = False
    txtCPhone.Enabled = False
    BDate.Enabled = False
    txtCEmail.Enabled = False
    txtDiscountAmt.Enabled = False
    txtRPoint.Enabled = False
    chkActive.Enabled = False
End Sub

Private Sub allClear()
    txtDiscountCard.text = ""
    txtCName.text = ""
    txtCAddress.text = ""
    txtCPhone.text = ""
    BDate.Value = Date
    txtCEmail.text = ""
    chkActive.Value = 0
    txtDiscountAmt.text = "0"
    txtRPoint.text = "0"
End Sub

Private Function rcupdate() As Boolean
'    On Error GoTo ErrHandler
    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then
    
    cn.Execute "INSERT INTO RMSCustomer(CID,DCard,CName,CAddress, " & _
                   " CPhone,BDate,CEmail,DiscountAmt,RPoint,Active) " & _
                   " VALUES ('" & parseQuotes(txtCID) & "','" & parseQuotes(txtDiscountCard) & "','" & parseQuotes(txtCName) & "', " & _
                   " '" & parseQuotes(txtCAddress) & "','" & parseQuotes(txtCPhone) & "', " & _
                   " '" & parseQuotes(BDate) & "', " & _
                   " '" & parseQuotes(txtCEmail) & "', " & _
                   " " & Val(txtDiscountAmt) & "," & Val(txtRPoint.text) & "," & _
                   " '" & parseQuotes(chkActive) & "') "


          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else
    
    
cn.Execute "Update RMSCustomer Set BDate='" & Format(BDate, "dd-mmm-yyyy") & "',DCard='" & parseQuotes(txtDiscountCard) & "', " & _
                  "CName='" & parseQuotes(txtCName) & "',CAddress='" & parseQuotes(txtCAddress) & "', " & _
                  "CEmail='" & parseQuotes(txtCEmail) & "',CPhone='" & parseQuotes(txtCPhone) & "'," & _
                  "DiscountAmt=" & Val(txtDiscountAmt.text) & ",RPoint=" & Val(txtRPoint.text) & ", " & _
                  "Active= '" & chkActive & "' Where CID ='" & txtCID & "'"

        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
'    Exit Sub
    Exit Function

End Function

Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtCID = rsfactory("CID")
        txtDiscountCard = rsfactory("DCard")
        txtCName = rsfactory("CName")
        txtCAddress = rsfactory("CAddress")
        txtCPhone = rsfactory("CPhone") & ""
        BDate = rsfactory("Bdate")
        txtCEmail = IIf(IsNull(rsfactory("CEmail")), "", rsfactory("CEmail"))
        txtDiscountAmt = rsfactory("DiscountAmt")
        txtRPoint = rsfactory("RPoint")
'        chkActive.Value = rsfactory("Active")
    End If
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If (txtDiscountCard.text = "") Then
       MsgBox "Enter Guest Discount Card No"
       txtDiscountCard.SetFocus
       IsValidRecord = False
       Exit Function
    End If
    
    If (txtCName.text = "") Then
      MsgBox "Enter Guest Name"
      txtCName.SetFocus
      IsValidRecord = False
      Exit Function
    
    End If
    
    If (txtCAddress.text = "") Then
      MsgBox "Enter Guest Address"
      txtCAddress.SetFocus
      IsValidRecord = False
      Exit Function
    End If
    
    If (txtDiscountAmt.text = "") Then
      MsgBox "Enter Guest Discount Info."
      txtDiscountAmt.SetFocus
      IsValidRecord = False
      Exit Function
    End If
    
'If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
'        If rsfactory.RecordCount > 0 Then
        If rsfactory.State <> 0 Then rsfactory.Close
            rsfactory.Open "select * from RMSCustomer where upper(DCard)='" & Strings.UCase(Strings.Trim(parseQuotes(txtDiscountCard))) & "'", cn

             If Not rsfactory.EOF Then
        MsgBox "This Card No already exists Please Enter Another.", vbInformation, Me.Caption & " - " & App.Title
          txtDiscountCard.SetFocus
          IsValidRecord = False
         Exit Function
            End If

'         End If
'        End If
    End Function
'.............................................................................

Public Sub printReport()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf      As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\CustomerInformationPreview.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select RMSCustomer.CID,RMSCustomer.CName,RMSCustomer.CAddress, " & _
             "  " & _
             "RMSCustomer.CPhone,RMSCustomer.CFax,RMSCustomer.CEmail " & _
             "from RMSCustomer where " & _
             "RMSCustomer.CID='" & Me.txtCID & "'"

    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True

'      Set oReportFormulaFieldDefinations = oReport.FormulaFields
'      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
'      oReportFF.text = "'Factory Information'"

oReport.DiscardSavedData
oReport.Preview "Customer Infromation of '" & txtCName.text & "'", , , , , 16777216 Or 524288 Or 65536

End Sub

Public Sub PopulateCnf(StrID As String)
    rsfactory.MoveFirst
    rsfactory.Find "CID=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub



