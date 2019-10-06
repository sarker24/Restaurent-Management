VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRRoom 
   BackColor       =   &H00C0B4A9&
   Caption         =   " Restaurant Room"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "frmRRoom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   120
      Picture         =   "frmRRoom.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   870
   End
   Begin VB.CommandButton cmdLAdd 
      BackColor       =   &H00C0B4A9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      Picture         =   "frmRRoom.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Add"
      Top             =   840
      Width           =   420
   End
   Begin VB.CommandButton cmdLDelete 
      BackColor       =   &H00C0B4A9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      Picture         =   "frmRRoom.frx":13DE
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Remove"
      Top             =   1200
      Width           =   420
   End
   Begin VB.TextBox txtTableID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtTableName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Text            =   " "
      Top             =   240
      Width           =   2895
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
      Height          =   735
      Left            =   960
      Picture         =   "frmRRoom.frx":1968
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Picture         =   "frmRRoom.frx":2232
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Picture         =   "frmRRoom.frx":2AFC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdPreview 
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
      Height          =   735
      Left            =   4560
      Picture         =   "frmRRoom.frx":33C6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Picture         =   "frmRRoom.frx":3C90
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Find "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      Picture         =   "frmRRoom.frx":421A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   " &Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      Picture         =   "frmRRoom.frx":47A4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   975
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTableNo 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7605
      _cx             =   13414
      _cy             =   9128
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   12629161
      BackColorAlternate=   12629161
      GridColor       =   12629161
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRRoom.frx":506E
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblTableID 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Table ID"
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
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Name"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmRRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 
  Private rsTableMaster                   As ADODB.Recordset
  Private rsTableDetail                   As ADODB.Recordset
  Private rs                              As ADODB.Recordset
  Private bRecordExists                  As Boolean
  Dim intStatus                               As Integer
  Dim str As String
  Dim flagSlNo                            As Integer
' Dim s As String
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'----Add For Reporting Perpose----------------------------------------------
Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition


Private objReportSub                        As CRPEAuto.Report 'sub
Private objReportDatabaseSub                As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition


Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private rsDailyRpt                          As ADODB.Recordset
Private Tracer                              As Integer
Private strGroupName                        As String
Private temp As Double
Private temp1 As Double
'--------------------------------------------------------------------------------


Private Sub chameleonButton1_Click()
    Call PrintReport
End Sub


Private Sub cmdCancel_Click()
    
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    CmdDelete.Enabled = True
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
'    chameleonButton1.Enabled = True
    Call alldisable
    If Not rsTableMaster.EOF Then FindRecord
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDelete_Click()
Call MakeSound
End Sub

Private Sub cmdEdit_Click()
     If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtTableName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
        CmdDelete.Enabled = False
        cmdPreview.Enabled = False
        cmdPrint.Enabled = False
'        chameleonButton1.Enabled = False
        cmdLAdd.Enabled = True
        cmdLDelete.Enabled = True
        fgTableNo.Editable = flexEDKbdMouse
        txtTableID.Enabled = False
        
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = True
                CmdDelete.Enabled = True
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
                fgTableNo.Editable = flexEDNone
                Call alldisable
                rsTableMaster.Requery
                Dim s As String
                s = txtTableID
                rsTableMaster.MoveFirst
                rsTableMaster.Find "TableID='" & parseQuotes(s) & "'"
                FindRecord
            End If
        End If
    End If
End Sub

Private Sub cmdLAdd_Click()

If fgTableNo.Row = -1 Then
        fgTableNo.AddItem ""
        Exit Sub
    End If
    
    If fgTableNo.Col = fgTableNo.Cols - 1 Then
        fgTableNo.AddItem ""
        fgTableNo.Row = fgTableNo.Row + 1
    Else
        fgTableNo.AddItem ""
    End If
End Sub

Private Sub cmdLDelete_Click()
            With fgTableNo
        If .Row = 0 Or .Row = -1 Then Exit Sub

        If .Rows > 1 Then .RemoveItem .Row
    End With
End Sub

Private Sub cmdnew_Click()
    
    Set rs = New ADODB.Recordset

If cmdNew.Caption = "&New" Then
        
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdOpen.Enabled = False
        CmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdClose.Enabled = False
        cmdLAdd.Enabled = True
        cmdLDelete.Enabled = True
        cmdPreview.Enabled = False
        
'        chameleonButton1.Enabled = False
'        TextClear Me
        Call Clear
         
        fgTableNo.Rows = 1
        fgTableNo.Editable = flexEDKbdMouse
        Call allenable
        txtTableName.SetFocus
           If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(TableID),0) as TblID from tblRMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtTableID = Val(rs!TblID) + 1
'
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdLAdd.Enabled = False
                cmdLDelete.Enabled = False
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
                CmdDelete.Enabled = True
                Call alldisable
                s = txtTableID
                rsTableMaster.Requery
                rsTableMaster.MoveFirst
                rsTableMaster.Find "TableID='" & parseQuotes(s) & "'"
                FindRecord
            End If
        End If
    End If
End Sub

Private Sub Clear()
    txtTableID.text = ""
    txtTableName.text = ""
End Sub

Private Sub allenable()
    txtTableID.Enabled = True
    txtTableName.Enabled = True
    fgTableNo.Enabled = True
    cmdLAdd.Enabled = True
    cmdLDelete.Enabled = True
End Sub

Private Sub alldisable()
    txtTableID.Enabled = False
    txtTableName.Enabled = False
    fgTableNo.Enabled = False
    cmdLAdd.Enabled = False
    cmdLDelete.Enabled = False
End Sub

Private Sub cmdOpen_Click()
    frmRRoomSearch.Show vbModal
        
End Sub
    
Private Sub cmdPreview_Click()
intStatus = 0
Call PrintReport
End Sub


Private Sub cmdPrint_Click()
intStatus = 1
Screen.MousePointer = vbHourglass
Call PrintReport
Screen.MousePointer = vbDefault
End Sub

    Private Sub Form_Load()
         Call Connect
      ModFunction.StartUpPosition Me
         Call alldisable
         Set rsTableMaster = New ADODB.Recordset
         
 
  If rsTableMaster.State <> 0 Then rsTableMaster.Close
         rsTableMaster.Open "select TableID,TableName " & _
                         "FROM tblRMaster order by TableID", cn, adOpenStatic, adLockReadOnly


  If rsTableMaster.RecordCount > 0 Then
      rsTableMaster.MoveFirst
        bRecordExists = True
        
    Else
        bRecordExists = False
    End If

If Not rsTableMaster.EOF Then FindRecord
    txtTableID.Enabled = False

End Sub
 
 
 Private Function rcupdate() As Boolean

On Error GoTo ErrHandler
Dim strSQL As String
    Dim iRow As Integer
    Dim j As Integer


   cn.BeginTrans
'   flagSlNo = 0
 If cmdNew.Caption = "&Save" Then

'        cn.BeginTrans

        'General Information for Payment Master
     strSQL = "INSERT INTO tblRMaster (TableID, TableName) " & _
                "VALUES ('" & Me.txtTableID & "','" & Me.txtTableName & "')"
     cn.Execute strSQL



    'payment Detail Information Enter This table
            j = 0
            For j = 1 To fgTableNo.Rows - 1
                If fgTableNo.TextMatrix(j, 2) <> "" Then
                cn.Execute "INSERT INTO tblRDetail (TableID,TableNo,Remarks) " & _
                           "Values ('" & parseQuotes(Me.txtTableID) & "', " & _
                           "'" & parseQuotes(fgTableNo.TextMatrix(j, 2)) & "', " & _
                           "'" & parseQuotes(fgTableNo.TextMatrix(j, 3)) & "')"

          End If

            Next
        rcupdate = True
'        cn.CommitTrans
        MsgBox "Record added Successfully", vbInformation, "Confirmation"
    Else




    ' Update Information

'    cn.BeginTrans
       cn.Execute "UPDATE tblRMaster SET TableName='" & _
                 txtTableName & "' WHERE TableID = '" & parseQuotes(txtTableID) & "'"

 

        cn.Execute "DELETE FROM tblRDetail WHERE TableID='" & parseQuotes(txtTableID) & "'"
'          cn.Execute strSQL

           j = 0
            For j = 1 To fgTableNo.Rows - 1
                If fgTableNo.TextMatrix(j, 2) <> "" Then
                cn.Execute "INSERT INTO tblRDetail (TableID,TableNo,Remarks) " & _
                           "Values ('" & parseQuotes(Me.txtTableID) & "', " & _
                           "'" & parseQuotes(fgTableNo.TextMatrix(j, 2)) & "', " & _
                           "'" & parseQuotes(fgTableNo.TextMatrix(j, 3)) & "')"

          End If

            Next

        rcupdate = True
'        cn.CommitTrans
        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
    End If
'    flagSlNo = 1
    cn.CommitTrans

    Exit Function

ErrHandler:
    cn.RollbackTrans
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Order Number"
            txtTableID.SetFocus
        Case Else


   MsgBox cn.Errors(0).NativeError & " : " & cn.Errors(0).Description
    End Select

   
End Function

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If Trim(txtTableID) = "" Then
        MsgBox "Your are missing ReviceivedBy", vbInformation
        txtTableID.SetFocus
        IsValidRecord = False
        Exit Function
'    ElseIf Trim(cmbPaymentBy) = "" Then
'        MsgBox "Your are missing PaymentBy.", vbInformation
'        cmbPaymentBy.SetFocus
'        IsValidRecord = False
'        Exit Function
    End If
    End Function


Private Sub FindRecord()

    Dim i As Integer
    Dim strTableDetail As String
    Set rsTableDetail = New ADODB.Recordset
    txtTableID.text = rsTableMaster!TableId
    txtTableName.text = rsTableMaster!TableName
    
    

    fgTableNo.Rows = 1
    strTableDetail = "SELECT  TableID,TableNo, Remarks" & _
                " FROM tblRDetail " & _
                "WHERE TableID='" & parseQuotes(txtTableID.text) & "' order by TableID "
    
    
    
    rsTableDetail.CursorLocation = adUseClient
    rsTableDetail.Open strTableDetail, cn, adOpenStatic, adLockReadOnly


    If rsTableDetail.RecordCount <> 0 Then

        fgTableNo.Rows = rsTableDetail.RecordCount + 1
                i = 0
        For i = 1 To rsTableDetail.RecordCount
            fgTableNo.TextMatrix(i, 1) = rsTableDetail("TableID")
            fgTableNo.TextMatrix(i, 2) = rsTableDetail("TableNo")
            fgTableNo.TextMatrix(i, 3) = rsTableDetail("Remarks")
            rsTableDetail.MoveNext
        Next
    End If
        rsTableDetail.Close
End Sub


Public Sub PrintReport()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    If rsTableMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\Reports\Restaurant Table Information.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        


    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close

           
        strSQL = "SELECT tblRMaster.TableID,tblRMaster.TableName," & _
                 "tblRDetail.TableNo , tblRDetail.Remarks " & _
                 "FROM tblRMaster,tblRDetail where tblRMaster.TableID = tblRDetail.TableID " & _
                 "and tblRMaster.TableID ='" & Me.txtTableID & "'"

                
        rsDailyRpt.Open strSQL, cn, adOpenStatic
        
'        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"


        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        
        If intStatus = 0 Then
        objReport.Preview "Payment Report", , , , , 16777216 Or 524288 Or 65536
        
        Else
        
        objReport.PrintOut
        
        End If
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, " Restrudent Table  Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Restrudent Table Information Report"
    End Select
End Sub


Public Sub PopulateForm(StrID As String)
    rsTableMaster.MoveFirst
    rsTableMaster.Find "TableID=" & parseQuotes(StrID)
    If rsTableMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


