VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAccountsHead 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Head of Accounts"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   Icon            =   "frmAccountsHead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   5520
      Picture         =   "frmAccountsHead.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete"
      Height          =   795
      Left            =   4440
      Picture         =   "frmAccountsHead.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "C&lose"
      Height          =   795
      Left            =   3360
      Picture         =   "frmAccountsHead.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   795
      Left            =   2280
      Picture         =   "frmAccountsHead.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1200
      Picture         =   "frmAccountsHead.frx":2734
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   120
      Picture         =   "frmAccountsHead.frx":2FFE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Find Next"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Find Last"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Find Previous"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Find First"
      Top             =   3120
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Accounts Head Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6465
      Begin VB.TextBox txtAHName 
         Appearance      =   0  'Flat
         Height          =   400
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtAID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   400
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   3495
      End
      Begin MSForms.ComboBox cmbAHType 
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   1920
         Width           =   3495
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "6165;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDepartment 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   1440
         Width           =   3495
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "6165;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblDName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Department"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "A. Head Type"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label lblAYName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "A. Head Name"
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
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblAYearID 
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts ID"
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
         Top             =   480
         Width           =   1305
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3000
      Top             =   3240
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DCSearch"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Accounts Head Setup   "
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
      Left            =   -120
      TabIndex        =   11
      Top             =   0
      Width           =   6945
   End
End
Attribute VB_Name = "frmAccountsHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAccountsHead        As ADODB.Recordset
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
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions

Private Sub chameleonButton1_Click()
'    Call printReport
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'
Private Sub cmdCancel_Click()

    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
'    chameleonButton1.Enabled = True
    txtAID.Enabled = False
    Call allClear
    Call alldisable
    If Not rsAccountsHead.EOF Then FindRecord
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
     idelete = MsgBox("Do you want to delete this record?", vbYesNo)
     If frmLogin.txtUID.text = "md" Then
    If idelete = vbYes Then
  
    cn.Execute "Delete From AccountsHead Where AID ='" & parseQuotes(txtAID) & "'"
            Call allClear
    End If
        
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.EOF = True Then
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

        txtAID = Adodc1.Recordset!AID
       txtAHName = Adodc1.Recordset!AHName
       cmbAHType = Adodc1.Recordset!AHType
       
End If

End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtAID = Adodc1.Recordset!AID
       txtAHName = Adodc1.Recordset!AHName
       cmbAHType = Adodc1.Recordset!AHType
       
End If
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
'        chameleonButton1.Enabled = False
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(AID),0) as AID from AccountsHead"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtAID.text = Val(rs!AID) + 1
            
        Call allenable
        txtAHName.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtAID.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                s = txtAID
                rsAccountsHead.Requery
                rsAccountsHead.MoveFirst
                rsAccountsHead.Find "AID='" & parseQuotes(s) & "'"
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
        txtAHName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                rsAccountsHead.Requery

                Dim s As String
                s = txtAID
                rsAccountsHead.Find "AID='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtAID = Adodc1.Recordset!AID
       txtAHName = Adodc1.Recordset!AHName
       cmbAHType = Adodc1.Recordset!AHType
       

End If
End Sub

Private Sub cmdOpen_Click()
    frmAHSearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub Find_Click()
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
       cmdPrevious.Enabled = False
 Else
      cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtAID = Adodc1.Recordset!AID
       txtAHName = Adodc1.Recordset!AHName
       cmbAHType = Adodc1.Recordset!AHType
       
End If
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsAccountsHead = New ADODB.Recordset
    rsAccountsHead.Open "select * from AccountsHead order by AHName", cn, adOpenStatic, adLockReadOnly
    
    Call alldisable
    Call Department
    
   If rsAccountsHead.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If

    If Not rsAccountsHead.EOF Then FindRecord

    txtAID.Enabled = False

    cmbAHType.AddItem "Expenses"
    cmbAHType.AddItem "Income"
    cmbAHType.AddItem "Assets"
    cmbAHType.AddItem "Liabilites"

    Adodc1.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "AccountsHead"

  Adodc1.Refresh

    
End Sub

Private Sub Department()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT DName FROM Department ORDER BY DName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbDepartment.AddItem rsTemp2("DName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub allenable()
    txtAHName.Enabled = True
    cmbDepartment.Enabled = True
    cmbAHType.Enabled = True
    End Sub

Private Sub alldisable()
    txtAID.Enabled = False
    txtAHName.Enabled = False
    cmbDepartment.Enabled = False
    cmbAHType.Enabled = False
    
End Sub


Private Sub allClear()
    txtAHName.text = ""
    cmbDepartment.text = ""
    cmbAHType.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO AccountsHead(AID,AHName,Department,AHType) " & _
                   " VALUES ('" & parseQuotes(txtAID) & "','" & parseQuotes(txtAHName) & "', " & _
                   " '" & parseQuotes(cmbDepartment) & "','" & parseQuotes(cmbAHType) & "')"
                   
                   
         rcupdate = True
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    Else

        cn.Execute "Update AccountsHead Set AHName='" & parseQuotes(txtAHName) & "',Department='" & parseQuotes(cmbDepartment) & _
                  "',AHType='" & parseQuotes(cmbAHType) & "' WHERE AID = '" & parseQuotes(txtAID) & "'"

                  
                 
     rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans

    Exit Function



ErrHandler:
    cn.RollbackTrans
    rsAccountsHead.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate AccountsHead Name"
            txtAHName = ""
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsAccountsHead.EOF Then
        txtAID = rsAccountsHead("AID")
        txtAHName = rsAccountsHead("AHName")
        cmbDepartment = rsAccountsHead("Department")
        cmbAHType = rsAccountsHead("AHType")
End If
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (txtAHName.text = "") Then
       MsgBox "Enter AccountsHead Name"
       txtAHName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    
If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
        If rsAccountsHead.RecordCount > 0 Then
        If rsAccountsHead.State <> 0 Then rsAccountsHead.Close
            rsAccountsHead.Open "select * from AccountsHead where upper(AHName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtAHName))) & "'", cn

             If Not rsAccountsHead.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtAHName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
End Function
'.............................................................................

Public Sub PopulateForm(StrID As String)

 Set rsAccountsHead = New ADODB.Recordset

    If rsAccountsHead.State <> 0 Then rsAccountsHead.Close
        rsAccountsHead.Open "select AID,AHName,Department,AHType from AccountsHead", cn, adOpenStatic, adLockReadOnly
                          

        rsAccountsHead.MoveFirst
    
    rsAccountsHead.Find "AID=" & parseQuotes(StrID)
    If rsAccountsHead.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub



Public Sub printReport()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf         As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\PartyInformationPreview.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select AccountsHead.AID,AccountsHead.WaiterName,AccountsHead.WaiterRemaks"
             
    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True
oReport.DiscardSavedData
oReport.Preview "AccountsHead Infromation of '" & txtAHName.text & "'", , , , , 16777216 Or 524288 Or 65536


End Sub


Public Sub PopulateCnf(StrID As String)


    rsAccountsHead.MoveFirst
    rsAccountsHead.Find "AID=" & parseQuotes(StrID)
    If rsAccountsHead.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

Private Sub txtAHName_GotFocus()
txtAHName.SelStart = 0
txtAHName.SelLength = Len(txtAHName)

End Sub


