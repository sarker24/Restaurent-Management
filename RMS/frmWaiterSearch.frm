VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRestrudentWaiterSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Waiter Search"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "frmWaiterSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   240
      TabIndex        =   3
      Top             =   6600
      Width           =   3015
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   5880
      Picture         =   "frmWaiterSearch.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   750
      Left            =   6975
      Picture         =   "frmWaiterSearch.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   4800
      Picture         =   "frmWaiterSearch.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1100
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   5610
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   7785
      _cx             =   13732
      _cy             =   9895
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
      FormatString    =   $"frmWaiterSearch.frx":1FE8
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Waiters  Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label lblIteamGroup 
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
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   3015
   End
End
Attribute VB_Name = "frmRestrudentWaiterSearch"
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
'frmLedgerParty.Show vbModal
  
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
'            rsTemp.CursorLocation = adUseClient
     rsTemp.Open "SELECT SerialNo,WaiterName,Remarks " & _
                 "FROM tblWaiterName WHERE tblWaiterName.WaiterName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
'   End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("WaiterName") & _
         vbTab & rsTemp("Remarks")
         
        rsTemp.MoveNext
        Wend

End Sub

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Menu Group From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmItemGroupSearch = Nothing
End Sub



Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
'            rsTemp.CursorLocation = adUseClient
     rsTemp.Open "SELECT SerialNo,WaiterName,Remarks FROM tblWaiterName", cn, adOpenStatic, adLockReadOnly
        
'   End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("WaiterName") & _
         vbTab & rsTemp("Remarks")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
'     If fgExport.Rows = 1 Then fgExport.AddItem ""


End Sub



    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmRestaurantWaiters.PopulateIteam fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub txtSearch_Change()
cmdFind_Click
End Sub




