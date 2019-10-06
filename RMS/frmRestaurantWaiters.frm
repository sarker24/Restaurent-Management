VERSION 5.00
Begin VB.Form frmRestaurantWaiters 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restaurant Waiter "
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "frmRestaurantWaiters.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtWaiterName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   3975
      End
      Begin VB.CommandButton CmdOpen 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4920
         Picture         =   "frmRestaurantWaiters.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton chameleonButton1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3945
         Picture         =   "frmRestaurantWaiters.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   1065
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0B4A9&
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3000
         Picture         =   "frmRestaurantWaiters.frx":171E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1080
         Picture         =   "frmRestaurantWaiters.frx":1FE8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0B4A9&
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         Picture         =   "frmRestaurantWaiters.frx":28B2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2040
         Picture         =   "frmRestaurantWaiters.frx":317C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   945
      End
      Begin VB.TextBox txtSerialNo 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblWaiterId 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Waiter ID"
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblWaiterName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Waiter Name"
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
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmRestaurantWaiters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsWaitersGroup        As ADODB.Recordset
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
    chameleonButton1.Enabled = True
    txtSerialNo.Enabled = False
    Call allClear
    Call alldisable
    If Not rsWaitersGroup.EOF Then FindRecord
End Sub

Private Sub cmdnew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
        chameleonButton1.Enabled = False
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),0) as SerialNo from tblWaiterName"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo.text = Val(rs!SerialNo) + 1
            
        Call allenable
        txtWaiterName.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSerialNo.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                chameleonButton1.Enabled = True
                Call alldisable
                s = txtSerialNo
                rsWaitersGroup.Requery
                rsWaitersGroup.MoveFirst
                rsWaitersGroup.Find "SerialNo='" & parseQuotes(s) & "'"
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
        txtWaiterName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
        chameleonButton1.Enabled = False
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                chameleonButton1.Enabled = True
                Call alldisable
                rsWaitersGroup.Requery

                Dim s As String
                s = txtSerialNo
                rsWaitersGroup.Find "SerialNo='" & parseQuotes(s) & "'"
'                Call search
'                Call countrysearch
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    frmRestrudentWaiterSearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub Find_Click()
    frmItemGroup.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsWaitersGroup = New ADODB.Recordset
    rsWaitersGroup.Open "select * from tblWaiterName order by SerialNo", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsWaitersGroup.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
   
    If Not rsWaitersGroup.EOF Then FindRecord
    
    txtSerialNo.Enabled = False
    
End Sub

Private Sub allenable()
   txtSerialNo.Enabled = True
    txtWaiterName.Enabled = True
    txtRemarks.Enabled = True
End Sub

Private Sub alldisable()
    txtSerialNo.Enabled = False
    txtWaiterName.Enabled = False
    txtRemarks.Enabled = False
    
End Sub


Private Sub allClear()
'    ModFunction.TextClear Me
    txtWaiterName.text = ""
    txtRemarks.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO tblWaiterName(SerialNo,WaiterName,Remarks) " & _
                   " VALUES ('" & parseQuotes(txtSerialNo) & "','" & parseQuotes(txtWaiterName) & "', " & _
                   " '" & parseQuotes(txtRemarks) & "')"
                   
                   
         rcupdate = True
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    Else

        cn.Execute "Update tblWaiterName Set WaiterName='" & parseQuotes(txtWaiterName) & _
                  "',Remarks='" & parseQuotes(txtRemarks) & "' WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"

                  
                 
     rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans

    Exit Function



ErrHandler:
    cn.RollbackTrans
    rsWaitersGroup.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Item Group Name"
            txtWaiterName = ""
            txtWaiterName.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsWaitersGroup.EOF Then
        txtSerialNo = rsWaitersGroup("SerialNo")
        txtWaiterName = rsWaitersGroup("WaiterName")
        txtRemarks = rsWaitersGroup("Remarks")
End If
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (txtWaiterName.text = "") Then
       MsgBox "Enter Waiter Name"
       txtWaiterName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    
'If CmdEdit.Caption <> "&Update" Or CmdEdit.Caption = "&Update" Then
'        If rsWaitersGroup.RecordCount > 0 Then
'        If rsWaitersGroup.State <> 0 Then rsWaitersGroup.Close
'            rsWaitersGroup.Open "select * from tblWaiterName where upper(WaiterName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtWaiterName))) & "'", cn
'
'             If Not rsWaitersGroup.EOF Then
'        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
'          txtWaiterName.SetFocus
'          IsValidRecord = False
'         Exit Function
'            End If
'
'         End If
'    End If
End Function
'.............................................................................

Public Sub PrintReport()
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

    strSQL = "select tblWaiterName.SerialNo,tblWaiterName.WaiterName,tblWaiterName.Remarks"
             
    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True
oReport.DiscardSavedData
oReport.Preview "Item Group Infromation of '" & txtWaiterName.text & "'", , , , , 16777216 Or 524288 Or 65536


End Sub


Public Sub PopulateIteam(StrID As String)


    rsWaitersGroup.MoveFirst
    rsWaitersGroup.Find "SerialNo=" & parseQuotes(StrID)
    If rsWaitersGroup.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

