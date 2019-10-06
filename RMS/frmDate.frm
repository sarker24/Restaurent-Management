VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Date"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2550
   Icon            =   "frmDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   16711680
      BackColor       =   12632256
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   20381697
      CurrentDate     =   38015
   End
   Begin VB.Label lblRow 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblCol 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Dim dt As Variant
'
'Sub UpdateDate()
'    If Not Enabled Then Exit Sub
'    'If Not IsDate(txtMonth & "/" & txtDay & "/" & txtYear) Then Beep: Exit Sub
'    dt = Format(MonthView1.Value, "dd-mmm-yyyy")
'    'lblDate = Format(dt, "Long Date")
'    Tag = dt
'End Sub
'
'Private Sub Form_Activate()
'    dt = Tag
'    Enabled = False
'    'lblDate = Format(dt, "Long Date")
'    'txtYear = Year(dt)
'    'txtDay = Day(dt)
'    'txtMonth.ListIndex = Month(dt) - 1
'    MonthView1.Value = CDate(dt)
'    Enabled = True
'End Sub
'
'Private Sub Form_Deactivate()
'    On Error Resume Next
'
'    ' update grid value
'    If IsDate(Tag) Then
'        Select Case LCase(strCallingForm)
'            Case LCase("frmLedger")
'                 frmLedger.fgLedger.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
'                 frmLedger.fgLedger.SetFocus
''            Case LCase("frmLabdip")
''                 frmLabdip.vfgLabdipDet.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmLabdip.vfgLabdipDet.SetFocus
''            Case LCase("frmJob")
''                 frmJob.fgTask.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmJob.fgTask.SetFocus
''            Case LCase("frmSampleDetail")
''               If frmSampleDetail.grdItem Then
''                  frmSampleDetail.grdItem.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                  frmSampleDetail.grdItem.SetFocus
''               End If
''               If frmSampleDetail.grdSample Then
''                  frmSampleDetail.grdSample.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                  frmSampleDetail.grdSample.SetFocus
''               End If
''            Case LCase("frmRoughCosting")
''                 frmRoughCosting.fgStyle.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmRoughCosting.fgStyle.SetFocus
''            Case LCase("frmInsurance")
''                 frmInsurance.grdInsurDet.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmInsurance.grdInsurDet.SetFocus
''            Case LCase("frmPliminaryCosting")
''                 frmPliminaryCosting.fgStyle.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmPliminaryCosting.fgStyle.SetFocus
''            Case LCase("frmExportOrder")
''                 frmExportOrder.fgDelvShcedule.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmExportOrder.fgDelvShcedule.SetFocus
''            Case LCase("frmPostshipment")
''                 frmPostshipment.fgMasterLC.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmPostshipment.fgMasterLC.SetFocus
''            Case LCase("frmBTBLC")
''                 frmBTBLC.grdResult.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmBTBLC.grdResult.SetFocus
''            Case LCase("frmPrchLocalPurchaseOrder")
''                 frmPrchLocalPurchaseOrder.fgPODetails.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmPrchLocalPurchaseOrder.fgPODetails.SetFocus
''            Case LCase("frmReceiveMasterLC")
''                 frmReceiveMasterLC.vfgrecmasLCDet.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
''                 frmReceiveMasterLC.vfgrecmasLCDet.SetFocus
''            Case LCase("frmMasterLC")
''                frmMasterLC.fgDelvShcedule.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-MMM-yyyy")
''                frmMasterLC.fgDelvShcedule.SetFocus
''            Case LCase("frmOrderPORequisition")
''                frmOrderPORequisition.fgPODetails.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-MMM-yyyy")
''                frmOrderPORequisition.fgPODetails.SetFocus
''            Case LCase("frmPI")
''                frmPI.fgPODetails.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-MMM-yyyy")
''                frmPI.fgPODetails.SetFocus
'        End Select
'    End If
'
'    ' go away
'    Hide
'
'    strCallingForm = ""
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    ' esc clears, then quits
'    If KeyAscii = 27 Then Tag = "": Hide
'
'    ' enter quits
'    If KeyAscii = 13 Then
'        UpdateDate
'        Hide
'    End If
'End Sub
'
'
'Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)
'    UpdateDate
'    Hide
'End Sub
'
'
'
