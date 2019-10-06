VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmItemSearch 
   BackColor       =   &H00C0B4A9&
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   Icon            =   "frmItemSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
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
      Left            =   8040
      Picture         =   "frmItemSearch.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
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
      Left            =   6960
      Picture         =   "frmItemSearch.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1100
   End
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
      Left            =   5880
      Picture         =   "frmItemSearch.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   1100
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   " "
      Top             =   8280
      Width           =   2535
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   8280
      Width           =   1935
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   7410
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   9105
      _cx             =   16060
      _cy             =   13070
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
      BackColor       =   12629161
      ForeColor       =   -2147483640
      BackColorFixed  =   32768
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12629161
      ForeColorSel    =   0
      BackColorBkg    =   12629161
      BackColorAlternate=   12632256
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmItemSearch.frx":1FE8
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
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Search "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9225
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
      Left            =   120
      TabIndex        =   7
      Top             =   8040
      Width           =   2535
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
      Left            =   2760
      TabIndex        =   6
      Top             =   8040
      Width           =   1815
   End
End
Attribute VB_Name = "frmItemSearch"
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

If cboMode.text = "Item Name" Then

     
rsTemp.Open "SELECT SerialNo,ItemCode, ItemName,ItemPrice,ItemGroup, ItemCatagory " & _
            "FROM tblItemDetail WHERE tblItemDetail.ItemName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Item Code" Then

rsTemp.Open "SELECT SerialNo,ItemCode, ItemName,ItemPrice,ItemGroup, ItemCatagory " & _
            "FROM tblItemDetail WHERE tblItemDetail.ItemCode LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Menu Group" Then

rsTemp.Open "SELECT SerialNo,ItemCode, ItemName,ItemPrice,ItemGroup, ItemCatagory " & _
            "FROM tblItemDetail WHERE tblItemDetail.ItemGroup LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
ElseIf cboMode.text = "Menu Catagory" Then

rsTemp.Open "SELECT SerialNo,ItemCode, ItemName,ItemPrice,ItemGroup, ItemCatagory " & _
            "FROM tblItemDetail WHERE tblItemDetail.ItemCatagory LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
Else
 
 rsTemp.Open "SELECT SerialNo,ItemCode, ItemName,ItemPrice,ItemGroup, ItemCatagory " & _
             "FROM tblItemDetail WHERE tblItemDetail.SerialNo LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ItemCode") & _
         vbTab & rsTemp("ItemName") & vbTab & rsTemp("ItemPrice") & vbTab & rsTemp("ItemGroup") & vbTab & rsTemp("ItemCatagory")
         
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
     rsTemp.Open "SELECT  SerialNo, ItemCode, ItemName, ItemQty, ItemPrice, Tips, NoDiscount, ItemGroup, ItemCatagory FROM tblItemDetail order by ItemCode ASC ", cn, adOpenStatic, adLockReadOnly
        
'   End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ItemCode") & _
         vbTab & rsTemp("ItemName") & vbTab & rsTemp("ItemPrice") & vbTab & rsTemp("ItemGroup") & vbTab & rsTemp("ItemCatagory")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
       
       cboMode.AddItem "Item Name"
       cboMode.AddItem "Item Code"
       cboMode.AddItem "Menu Group"
       cboMode.AddItem "Menu Catagory"
       cboMode.AddItem "Serial Number"
       
       cboMode.text = "Item Name"

End Sub



    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmItemEntry.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub txtSearch_Change()
cmdFind_Click
End Sub




