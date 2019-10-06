VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmCurrencySearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Currency Search"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   5295
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   750
      Left            =   6375
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame frmImportarSearch 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Currency Detail"
      Height          =   3015
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   7215
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   2610
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6945
         _cx             =   12250
         _cy             =   4604
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
         BackColor       =   12629161
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CURRENCY  SEARCH  "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7305
   End
End
Attribute VB_Name = "frmCurrencySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsTemp                      As ADODB.Recordset
Private rsExport                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select  Currency From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmCompanySearch = Nothing
End Sub



Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub Form_Load()
     modCommon.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
'            rsTemp.CursorLocation = adUseClient
     rsTemp.Open "SELECT  CurrencyName,ShortForm,CurrencyID FROM Currency", cn, adOpenStatic, adLockReadOnly
        
'   End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("CurrencyID") & vbTab & rsTemp("CurrencyName") & vbTab & rsTemp("ShortForm")
         
         rsTemp.MoveNext
    Wend
        GridCount fgExport
'     If fgExport.Rows = 1 Then fgExport.AddItem ""


End Sub



    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmCurrency.PopulateCurrency fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub







