VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmUserSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "User Search"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   2760
      Picture         =   "frmUserSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   750
      Left            =   4935
      Picture         =   "frmUserSearch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   3840
      Picture         =   "frmUserSearch.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1100
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   3450
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7305
      _cx             =   12885
      _cy             =   6085
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmUserSearch.frx":1A5E
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Find User Name"
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
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "frmUserSearch"
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
        
       
'            rsTemp.CursorLocation = adUseClient
     rsTemp.Open "SELECT SerialNo,UID,UName,UPassword,EDate " & _
                 "FROM RMSUser WHERE RMSUser.UID LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
'   End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("UID") & vbTab & rsTemp("UName") & _
         vbTab & rsTemp("UPassword")
         
        rsTemp.MoveNext
        Wend
End Sub

Private Sub cmdOk_Click()

If fgExport.RowSel < 0 Then
        MsgBox "Please Select a User Name From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmUserSearch = Nothing
End Sub

Private Sub Form_Load()

ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
'            rsTemp.CursorLocation = adUseClient
     rsTemp.Open "SELECT SerialNo,UID,UName,UPassword,EDate FROM RMSUser", cn, adOpenStatic, adLockReadOnly
        
'   End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("UID") & vbTab & rsTemp("UName") & _
         vbTab & rsTemp("UPassword") & vbTab & rsTemp("EDate")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport

End Sub

Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        frmUser.PopulateCnf fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub
    Private Sub txtSearch_Change()
        cmdFind_Click
    End Sub
 

