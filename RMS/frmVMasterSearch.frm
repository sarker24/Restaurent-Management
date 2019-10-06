VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmVMasterSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Voucher Search"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   7575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10455
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   7125
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   10200
         _cx             =   17992
         _cy             =   12568
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   12629161
         ForeColor       =   -2147483640
         BackColorFixed  =   12632064
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   12629161
         BackColorAlternate=   14737632
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
         GridLines       =   3
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmVMasterSearch.frx":0000
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
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Text            =   " "
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7440
      Picture         =   "frmVMasterSearch.frx":00C4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8520
      Picture         =   "frmVMasterSearch.frx":098E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9480
      Picture         =   "frmVMasterSearch.frx":1258
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   7920
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   7920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyyy"
      Format          =   16646147
      CurrentDate     =   41840
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RMS;Data Source=NOTEBOOK"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RMS;Data Source=NOTEBOOK"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.Label lblIteamGroup 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Field Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Available Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   7680
      Width           =   1815
   End
End
Attribute VB_Name = "frmVMasterSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsExport                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
If cboMode.text = "Voucher Number" Then

      rsTemp.Open "SELECT TOP 50 VID,VDate,AHead,Department,Amode " & _
                 "FROM Voucher WHERE Voucher.VID LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
ElseIf cboMode.text = "Date Search" Then
      
      rsTemp.Open "SELECT TOP 50 VID,VDate,AHead,Department,Amode " & _
                 "FROM Voucher WHERE Voucher.VDate = '" & dtDate.Value & "'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Table Number" Then

      rsTemp.Open "SELECT TOP 50 VID,VDate,AHead,Department,Amode " & _
                 "FROM Voucher WHERE Voucher.AHead LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Waiter Name" Then

      rsTemp.Open "SELECT TOP 50 VID,VDate,AHead,Department,Amode " & _
                 "FROM Voucher WHERE Voucher.Department LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

Else
 
 rsTemp.Open "SELECT TOP 50 VID,VDate,AHead,Department,Amode " & _
                 "FROM Voucher WHERE Voucher.Amode LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

End If

'   rsTemp.Open
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
       fgExport.AddItem "" & vbTab & rsTemp("VID") & vbTab & Format(rsTemp("VDate"), "dd-mmm-yyyy") & _
         vbTab & rsTemp("AHead") & vbTab & rsTemp("Department") & vbTab & rsTemp("Amode")
         
        rsTemp.MoveNext
        Wend

End Sub

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select Voucher From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me

Set frmVMasterSearch = Nothing
End Sub

Private Sub dtDate_click()
cmdFind_Click
End Sub

Private Sub fgExport_Click()
cmdOk_Click
End Sub

Private Sub fgExport_KeyPress(KeyAscii As Integer)
cmdOk_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     fgExport.Editable = flexEDKbdMouse
     
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
     rsTemp.Open "SELECT TOP 50 VID,VDate,AHead,Department,Amode FROM Voucher", cn, adOpenStatic, adLockReadOnly
         
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("VID") & vbTab & Format(rsTemp("VDate"), "dd-mmm-yyyy") & _
         vbTab & rsTemp("AHead") & vbTab & rsTemp("Department") & vbTab & rsTemp("Amode")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
'     If fgExport.Rows = 1 Then fgExport.AddItem ""

       cboMode.AddItem "Voucher Number"
       cboMode.AddItem "Date Search"
       cboMode.AddItem "Accounts Head"
       cboMode.AddItem "Department"
       cboMode.AddItem "Accounts Mode"
       cboMode.text = "Voucher Number"

       dtDate.Value = Date
       

End Sub
  
    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmMoneyReceipt.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub
    
Private Sub txtSearch_Change()
cmdFind_Click
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys Chr(9)
'Call cmdOk_Click

End If
End Sub






