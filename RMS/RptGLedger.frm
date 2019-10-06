VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RptGLedger 
   BackColor       =   &H00C0B4A9&
   Caption         =   "General Ledger Informations"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   Icon            =   "RptGLedger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDateSelect 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Date Range"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
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
         Left            =   240
         TabIndex        =   3
         Top             =   1320
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
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox CmbAHead 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   3375
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2760
         Top             =   2880
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
               Picture         =   "RptGLedger.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptGLedger.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptGLedger.frx":11C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   3360
         TabIndex        =   4
         Top             =   2880
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
         Left            =   2160
         TabIndex        =   5
         Top             =   1320
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   67239939
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker FDate 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   2160
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   67239939
         CurrentDate     =   37114
      End
      Begin MSComCtl2.DTPicker TDate 
         Height          =   285
         Left            =   3810
         TabIndex        =   7
         Top             =   2160
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   67239939
         CurrentDate     =   37114
      End
      Begin MSAdodcLib.Adodc dcAHName 
         Height          =   330
         Left            =   120
         Top             =   2640
         Visible         =   0   'False
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AHName"
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
         Left            =   2160
         TabIndex        =   10
         Top             =   1920
         Width           =   2535
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
         Left            =   3840
         TabIndex        =   9
         Top             =   1920
         Width           =   1470
      End
      Begin VB.Label lblWaiter 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Accounts Head Name"
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
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "General Ledger Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "RptGLedger"
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

Private Sub CmbAHead_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(CmbAHead, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub


Private Sub AHead()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT AHName FROM AccountsHead ORDER BY AHName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        CmbAHead.AddItem rsTemp2("AHName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub Form_Load()
    Call Connect
    Call AHead
    ModFunction.StartUpPosition Me
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    CurrentDate.Value = Date
    FDate.Value = Date
    TDate.Value = Date

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
            MsgBox "Invalid Date and select accuDescription date range", vbInformation, "Genelral Ledger Report"
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


    rsMaster.Open "SELECT     VID,VDate,AHead, Debit, Credit,  " & _
                    "Description " & _
                    "Posted " & _
                    "From Voucher " & _
                    "WHERE    VDate='" & CurrentDate.Value & "' AND Posted = 'Posted' order by AHead", cn, adOpenStatic, adLockReadOnly



     End If

      If OpCustomDate.Value = True Then


     rsMaster.Open "SELECT VID,VDate,AHead, Debit, Credit, " & _
                    "Description " & _
                    "Posted " & _
                    "From Voucher " & _
                    "WHERE  VDate BETWEEN '" & FDate.Value & "' AND " & _
                    "'" & TDate.Value & "' AND Posted = 'Posted' and AHead='" & parseQuotes(CmbAHead) & "'", cn, adOpenStatic, adLockReadOnly


      End If

End Function


Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If


        strPath = App.Path + "\reports\General Ledger.rpt"
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
              objReportFF.text = "'" + CmbAHead + "'"


  End If


  If OpCustomDate.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)

              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + CmbAHead + "'"


   End If



        objReportDatabaseTable.SetPrivateData 3, rsMaster

        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True

        objReport.DiscardSavedData
        objReport.Preview "General Ledger of + CmbAHead", , , , , 16777216 Or 524288 Or 65536


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
            MsgBox "Request cancelled by the user", vbInformation, "General Ledger Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "General Ledger Report"
    End Select
End Sub




