VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RptWaiterSales 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Waiter Wise Sales Statement"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   Icon            =   "RptWaiterSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDateSelect 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   3015
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   5775
      Begin VB.OptionButton OpCustomDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&stom Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton OpCurrentDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&rrent Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbWaiter 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox chkAllWaiter 
         BackColor       =   &H00C0B4A9&
         Caption         =   "All Waiter Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2280
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptWaiterSales.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptWaiterSales.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptWaiterSales.frx":11C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   2880
         TabIndex        =   6
         Top             =   2160
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   1058
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Preview"
               Object.ToolTipText     =   "Preview"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Close"
               Object.ToolTipText     =   "Close"
               ImageIndex      =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComCtl2.DTPicker CurrentDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65142787
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker FDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   1680
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65142787
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker TDate 
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65142787
         CurrentDate     =   37114
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "To Date"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H00C0B4A9&
         Caption         =   "From Date"
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
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label lblWaiter 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Waiter Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label lblIWSSales 
      BackColor       =   &H00808000&
      Caption         =   "Waiter Wise Sales Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "RptWaiterSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsMaster                            As ADODB.Recordset
Private rsSelect                            As ADODB.Recordset 'sub
Private rscashmaster                        As New ADODB.Recordset

Private rsTemp2                             As ADODB.Recordset
Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition
Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private Tracer                              As Integer
Private strGroupName                        As String

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub chkAllWaiter_Click()
If chkAllWaiter.Value = 1 Then
cmbWaiter.Enabled = False
Else
cmbWaiter.Enabled = True
End If
End Sub

Private Sub cmbWaiter_Change()
If cmbWaiter <> rscashmaster!WaiterName Then
MsgBox "Invalid Waiter Name.", vbInformation, "Confirmation"
cmbWaiter.text = ""
cmbWaiter.SetFocus
End If
End Sub

Private Sub Form_Load()
    Call Connect
    Call WaiterName
    ModFunction.StartUpPosition Me
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    CurrentDate.Value = Date
    FDate.Value = Date
    TDate.Value = Date
            
End Sub

Private Sub cmbWaiter_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbWaiter, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbWaiter_GotFocus()
cmbWaiter.BackColor = &HFFFFC0
End Sub

Private Sub cmbWaiter_LostFocus()
cmbWaiter.BackColor = vbWhite
End Sub

Private Sub WaiterName()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT WaiterName FROM tblWaiterName ORDER BY WaiterName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbWaiter.AddItem rsTemp2("WaiterName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub OpCurrentDate_Click()
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    CurrentDate.Visible = True
    lblFrom.Visible = False
    FDate.Visible = False
    lblTo.Visible = False
    TDate.Visible = False
End Sub
Private Sub opCustomDate_Click()
    OpCustomDate.Visible = True
    OpCurrentDate.Visible = True
    lblFrom.Visible = True
    lblTo.Visible = True
    CurrentDate.Visible = False
    FDate.Visible = True
    TDate.Visible = True
End Sub
Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "Preview"
            If Validate Then
            If chkAllWaiter.Value = 1 Then
                Tracer = 0
                Call FetchData1
                Call previewReport1
                Else
                Call FetchData
            Call previewReport
                    End If
               End If
     Case "Print"
            If Validate Then
            If chkAllWaiter.Value = 1 Then
                Tracer = 1
                Call FetchData1
                Call previewReport1
                Else
                Call FetchData
            Call previewReport
                    End If
               End If
     Case "Close"
               Unload Me
    End Select

End Sub
Private Function Validate() As Boolean
           Validate = True
        If FDate.Value > TDate.Value Then
            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
            FDate.SetFocus
            Validate = False
            Exit Function
        End If
    End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Public Function FetchData()

    Set rsMaster = New ADODB.Recordset
    
    If OpCurrentDate.Value = True Then
    
    
    rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate, " & _
                  "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
                  "tblCashDetail.ItemCode,tblCashDetail.ItemName,tblCashDetail.Qty,tblCashDetail.Rate,tblCashDetail.Tips, " & _
                  "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount,(tblCashDetail.Qty*tblCashDetail.Tips) as TTips " & _
                  "FROM tblCashMaster INNER JOIN " & _
                  "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate='" & CurrentDate.Value & "' " & _
                  "And tblCashMaster.CActive = 'Active' And tblCashMaster.CPost = 'Posted' and tblCashMaster.WaiterName='" & parseQuotes(cmbWaiter) & "'", cn, adOpenStatic, adLockReadOnly
    
                             
     End If
             
      If OpCustomDate.Value = True Then
      
       rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate, " & _
                         "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
                         "tblCashDetail.ItemCode,tblCashDetail.ItemName,(tblCashDetail.Qty),tblCashDetail.Rate,tblCashDetail.Tips, " & _
                         "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount,(tblCashDetail.Qty*tblCashDetail.Tips) as TTips " & _
                         "FROM tblCashMaster INNER JOIN " & _
                         "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate  " & _
                         "BETWEEN '" & FDate.Value & "' AND '" & TDate.Value & "' And tblCashMaster.CActive = 'Active' " & _
                         "And tblCashMaster.CPost = 'Posted' and tblCashMaster.WaiterName='" & parseQuotes(cmbWaiter) & "'", cn, adOpenStatic, adLockReadOnly
                                            
      End If
                  
End Function


Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Waiter Wise Sales.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
   
   If OpCurrentDate.Value = True Then
   
   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
              objReportFF.text = "'" + Format(CurrentDate, "dd-MMM-yyyy") + "'"
   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbWaiter + "'"

              
  End If
  
  
  If OpCustomDate.Value = True Then
  
   Set objReportFF = objReportFormulaFieldDefinations.Item(2)

              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"
             
   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbWaiter + "'"

             
   End If


      
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Waiter Insformations", , , , , 16777216 Or 524288 Or 65536
    
      
     If Tracer = 1 Then
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
            MsgBox "Request cancelled by the user", vbInformation, "Waiter Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Waiter Information Report"
    End Select
End Sub


Public Function FetchData1()

Set rsMaster = New ADODB.Recordset
    
    If OpCurrentDate.Value = True Then
    
    
    rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate,tblCashMaster.WaiterName, " & _
                  "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
                  "tblCashDetail.ItemCode,tblCashDetail.ItemName,tblCashDetail.Qty,tblCashDetail.Rate,tblCashDetail.Tips, " & _
                  "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount,(tblCashDetail.Qty*tblCashDetail.Tips) as TTips " & _
                  "FROM tblCashMaster INNER JOIN " & _
                  "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate='" & CurrentDate.Value & "' " & _
                  "And tblCashMaster.CActive = 'Active' And tblCashMaster.CPost = 'Posted' ", cn, adOpenStatic, adLockReadOnly
    
                             
     End If
             
      If OpCustomDate.Value = True Then
      
       rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate,tblCashMaster.WaiterName, " & _
                         "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
                         "tblCashDetail.ItemCode,tblCashDetail.ItemName,(tblCashDetail.Qty),tblCashDetail.Rate,tblCashDetail.Tips, " & _
                         "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount,(tblCashDetail.Qty*tblCashDetail.Tips) as TTips " & _
                         "FROM tblCashMaster INNER JOIN " & _
                         "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate  " & _
                         "BETWEEN '" & FDate.Value & "' AND '" & TDate.Value & "' And tblCashMaster.CActive = 'Active' " & _
                         "And tblCashMaster.CPost = 'Posted' ", cn, adOpenStatic, adLockReadOnly
                                            
      End If
'    Set rsMaster = New ADODB.Recordset
'
'    If OpCurrentDate.Value = True Then
'
'
'    rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate, " & _
'                  "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
'                  "tblCashDetail.ItemCode,tblCashDetail.ItemName,tblCashDetail.Qty,tblCashDetail.Rate,tblCashDetail.Tips, " & _
'                  "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount,(tblCashDetail.Qty*tblCashDetail.Tips) as TTips " & _
'                  "FROM tblCashMaster INNER JOIN " & _
'                  "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate='" & CurrentDate.Value & "' " & _
'                  "And tblCashMaster.CActive = 'Active' And tblCashMaster.CPost = 'Posted' ", cn, adOpenStatic, adLockReadOnly
'
'
'     End If
'
'      If OpCustomDate.Value = True Then
'
'       rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate, " & _
'                         "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
'                         "tblCashDetail.ItemCode,tblCashDetail.ItemName,(tblCashDetail.Qty),tblCashDetail.Rate,tblCashDetail.Tips, " & _
'                         "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount,(tblCashDetail.Qty*tblCashDetail.Tips) as TTips " & _
'                         "FROM tblCashMaster INNER JOIN " & _
'                         "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate  " & _
'                         "BETWEEN '" & FDate.Value & "' AND '" & TDate.Value & "' And tblCashMaster.CActive = 'Active' " & _
'                         "And tblCashMaster.CPost = 'Posted'", cn, adOpenStatic, adLockReadOnly
'
'      End If
                  
End Function


Public Sub previewReport1()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Waiter Wise Sales Summary.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
   
   If OpCurrentDate.Value = True Then
   
   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
              objReportFF.text = "'" + Format(CurrentDate, "dd-MMM-yyyy") + "'"
              
  End If
  
  
  If OpCustomDate.Value = True Then
  
   Set objReportFF = objReportFormulaFieldDefinations.Item(2)

              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"
             
   End If


      
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Waiter wise Sales Summary Statement.", , , , , 16777216 Or 524288 Or 65536
    
      
     If Tracer = 1 Then
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
            MsgBox "Request cancelled by the user", vbInformation, "Waiter wise Sales Summery Statement."
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Waiter wise Sales Summery Statement."
    End Select
End Sub

'On Error GoTo ErrH
'    Dim strPath As String
'
'    If rsMaster.RecordCount = 0 Then
'        MsgBox "Data not available", vbInformation
'        Exit Sub
'    End If
'
'
'        strPath = App.Path + "\reports\Waiter Wise Sales Summary.rpt"
'        Set objReportApp = CreateObject("Crystal.CRPE.Application")
'        Set objReport = objReportApp.OpenReport(strPath)
'        Set objReportDatabase = objReport.Database
'        Set objReportDatabaseTables = objReportDatabase.Tables
'        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
'        Set ObjPrinterSetting = objReport.PrintWindowOptions
'        Set objReportFormulaFieldDefinations = objReport.FormulaFields
'
'   If OpCurrentDate.Value = True Then
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'              objReportFF.text = "'" + Format(CurrentDate, "dd-MMM-yyyy") + "'"
'   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
'              objReportFF.text = "'" + cmbWaiter + "'"
'
'
'  End If
'
'
'  If OpCustomDate.Value = True Then
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
'
'              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
'             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
'              objReportFF.text = "'" + cmbWaiter + "'"
'
'
'   End If
'
'
'
'        objReportDatabaseTable.SetPrivateData 3, rsMaster
'
'        ObjPrinterSetting.HasPrintSetupButton = True
'        ObjPrinterSetting.HasRefreshButton = True
'        ObjPrinterSetting.HasSearchButton = True
'        ObjPrinterSetting.HasZoomControl = True
'
'        objReport.DiscardSavedData
'        objReport.Preview "Waiter Insformations", , , , , 16777216 Or 524288 Or 65536
'
'
'     If Tracer = 1 Then
'    objReport.PrintOut
'    End If
'
'        Set objReport = Nothing
'        Set objReportDatabase = Nothing
'        Set objReportDatabaseTables = Nothing
'        Set objReportDatabaseTable = Nothing
'    Exit Sub
'
'ErrH:
'
'    Select Case Err.Number
'        Case 20545
'            MsgBox "Request cancelled by the user", vbInformation, "Waiter Information Report"
'        Case Else
'            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Waiter Information Report"
'    End Select
'End Sub

