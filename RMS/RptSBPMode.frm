VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RptDSales 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Daily Sales Statement"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   Icon            =   "RptSBPMode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5655
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
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
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
         TabIndex        =   2
         Top             =   960
         Width           =   1815
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
               Picture         =   "RptSBPMode.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptSBPMode.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptSBPMode.frx":11C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   2880
         TabIndex        =   5
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
      Begin MSComCtl2.DTPicker SSCurrentDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65339395
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker SSFDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65339395
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker SSTDate 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65339395
         CurrentDate     =   37114
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
         TabIndex        =   8
         Top             =   960
         Width           =   1620
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
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Sales Statement Details"
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
      TabIndex        =   9
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "RptDSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsMaster                            As ADODB.Recordset
Private rsSelect                            As ADODB.Recordset 'sub

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

Private Sub Form_Load()
    Call Connect
    ModFunction.StartUpPosition Me
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    SSCurrentDate.Value = Date
    SSFDate.Value = Date
    SSTDate.Value = Date
            
End Sub

Private Sub OpCurrentDate_Click()
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    SSCurrentDate.Visible = True
    lblFrom.Visible = False
    SSFDate.Visible = False
    lblTo.Visible = False
    SSTDate.Visible = False
End Sub
Private Sub opCustomDate_Click()
    OpCustomDate.Visible = True
    OpCurrentDate.Visible = True
    lblFrom.Visible = True
    lblTo.Visible = True
    SSCurrentDate.Visible = False
    SSFDate.Visible = True
    SSTDate.Visible = True
End Sub
Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "Preview"
            If Validate Then
                Tracer = 0
                Call FetchData
                Call previewReport
               End If
     Case "Print"
            If Validate Then
                Tracer = 1
                Call FetchData
                Call previewReport
               End If
     Case "Close"
               Unload Me
    End Select

End Sub
Private Function Validate() As Boolean
           Validate = True
        If SSFDate.Value > SSTDate.Value Then
            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
            SSFDate.SetFocus
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
    
    
rsMaster.Open "SELECT SerialNo, strDate, TableName,(CONVERT(money,Guest)), ServiceCharge, Vat, Discount,DiscountAmt,strTime, TotalBill, TotalVat, " & _
              "TSCharge, TotalDiscount, ((CONVERT(money,NetPayable))-Due) as NetPayable, PaymentMode, " & _
              "Paid , Due, CActive, CPost " & _
              "From tblCashMaster WHERE " & _
               "tblCashMaster.strDate='" & SSCurrentDate.Value & "'", cn, adOpenStatic, adLockReadOnly

     End If
             
      If OpCustomDate.Value = True Then
      
rsMaster.Open "SELECT SerialNo, strDate, TableName, Guest, ServiceCharge, Vat, Discount, DiscountAmt,strTime,TotalBill, TotalVat, " & _
                "TSCharge, TotalDiscount, ( (CONVERT(money,NetPayable))-Due) as NetPayable, PaymentMode, " & _
                "Paid , Due, CActive, CPost " & _
                "From tblCashMaster WHERE " & _
                 "strDate BETWEEN '" & SSFDate.Value & "' AND '" & SSTDate.Value & "'", cn, adOpenStatic, adLockReadOnly
                                                  
      End If
                  
End Function


Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Daily Sales Statement.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
   
   If OpCurrentDate.Value = True Then
   
   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
              objReportFF.text = "'" + Format(SSCurrentDate, "dd-MMM-yyyy") + "'"
              
  End If
  
  
  If OpCustomDate.Value = True Then
  
   Set objReportFF = objReportFormulaFieldDefinations.Item(2)

              objReportFF.text = "'" + Format(SSFDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
             objReportFF.text = "'" + Format(SSTDate, "dd-MMM-yyyy") + "'"
             
   End If


      
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Sales Insformations", , , , , 16777216 Or 524288 Or 65536
    
      
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
            MsgBox "Request cancelled by the user", vbInformation, "Sales Summery Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Sales Summery Information Report"
    End Select
End Sub





