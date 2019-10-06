VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptSalesSummery 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Sales Summery"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "RptSalesSummery.frx":0000
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
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   5775
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
               Picture         =   "RptSalesSummery.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptSalesSummery.frx":0E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptSalesSummery.frx":173E
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
      Begin MSComCtl2.DTPicker CurrentDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   64946179
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker FDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   64946179
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker TDate 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   64946179
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
      BackColor       =   &H00808000&
      Caption         =   " Sales Summary Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "RptSalesSummery"
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
    CurrentDate.Value = Date
    TDate.Value = Date
    FDate.Value = Date
            
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
    
     rsMaster.Open "SELECT SerialNo, strDate, TableName, Guest, strTime, TotalBill, TotalVat, TSCharge, TotalDiscount, NetPayable,Paid, Due, " & _
           "PaymentMode , CActive, CPost, UName " & _
           "From tblCashMaster where strDate='" & CurrentDate.Value & "' " & _
           "And CActive = 'Active' And CPost = 'Posted'", cn, adOpenStatic, adLockReadOnly
    
             
     End If
             
      If OpCustomDate.Value = True Then
      
    rsMaster.Open "SELECT SerialNo, strDate, TableName, Guest, strTime, TotalBill, TotalVat, TSCharge, TotalDiscount, NetPayable,Paid, Due, " & _
           "PaymentMode , CActive, CPost, UName " & _
           "From tblCashMaster where strDate BETWEEN '" & FDate.Value & "' AND '" & TDate.Value & "' " & _
           "And CActive = 'Active' And CPost = 'Posted'", cn, adOpenStatic, adLockReadOnly
                                                  
      End If
                  
End Function


Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Sales Summary Statement.rpt"
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
        objReport.Preview "Sales Summary Insformations", , , , , 16777216 Or 524288 Or 65536
    
      
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

