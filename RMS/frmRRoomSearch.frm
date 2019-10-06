VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRRoomSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Table Search"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "frmRRoomSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3240
      MaskColor       =   &H00808080&
      Picture         =   "frmRRoomSearch.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5400
      MaskColor       =   &H00808080&
      Picture         =   "frmRRoomSearch.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4335
      MaskColor       =   &H00808080&
      Picture         =   "frmRRoomSearch.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1100
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5640
      Width           =   1935
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTableSearch 
      Height          =   4290
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   6075
      _cx             =   10716
      _cy             =   7567
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
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
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   0
      ForeColorFixed  =   32768
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   12629161
      BackColorAlternate=   8421504
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRRoomSearch.frx":1FE8
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
   Begin VB.Frame frmImportarSearch 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Table Master Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6375
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
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Table Search "
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
      Left            =   -2400
      TabIndex        =   6
      Top             =   0
      Width           =   8985
   End
End
Attribute VB_Name = "frmRRoomSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsTemp                      As ADODB.Recordset
Private rsTableSearch                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If fgTableSearch.RowSel < 0 Then
        MsgBox "Please Select a Country From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
    
     Unload Me
     Set frmRRoomSearch = Nothing
End Sub

Private Sub fgTableSearch_DblClick()
    cmdOK_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
  
             rsTemp.Open "SELECT TableID,TableName FROM tblRMaster", cn, adOpenStatic, adLockReadOnly
 
  
         fgTableSearch.Rows = 1
             While Not rsTemp.EOF
             fgTableSearch.AddItem "" & vbTab & rsTemp("TableID") & vbTab & rsTemp("TableName")
             rsTemp.MoveNext
        Wend
        
        GridCount fgTableSearch
        
End Sub



    Private Sub PopulateCompanySearch()
        
          If fgTableSearch.Row > 0 Then
             frmRRoom.PopulateForm fgTableSearch.TextMatrix(fgTableSearch.Row, 1)
          End If
   End Sub


