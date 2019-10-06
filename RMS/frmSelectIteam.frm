VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmCashSelectIteam 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Select Item Information"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   Icon            =   "frmSelectIteam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5280
      Width           =   1815
   End
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
      Left            =   8400
      Picture         =   "frmSelectIteam.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   7320
      Picture         =   "frmSelectIteam.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
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
      Left            =   6240
      Picture         =   "frmSelectIteam.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1100
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   4365
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   9840
      _cx             =   17357
      _cy             =   7699
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12632319
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
      SelectionMode   =   2
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSelectIteam.frx":1FE8
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
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblIteamGroup 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Item Code"
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
      Left            =   720
      TabIndex        =   6
      Top             =   5040
      Width           =   975
   End
End
Attribute VB_Name = "frmCashSelectIteam"
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
              
If txtSearch.text <> "" Then

rsTemp.Open "SELECT TOP 50 SerialNo,ItemCode,ItemName,ItemQty,ItemPrice,Tips,NoDiscount,NoVAT,ItemGroup,ItemCatagory " & _
                "FROM tblItemDetail WHERE tblItemDetail.ItemCode= '" & parseQuotes(txtSearch.text) & "'", cn, adOpenStatic, adLockReadOnly
        
Else
'If Text1.text <> "" Then
rsTemp.Open "SELECT TOP 50 SerialNo,ItemCode,ItemName,ItemQty,ItemPrice,Tips,NoDiscount,NoVAT,ItemGroup,ItemCatagory " & _
                 "FROM tblItemDetail WHERE tblItemDetail.ItemName LIKE '" & Text1.text & "%'", cn, adOpenStatic, adLockReadOnly
End If
'   rsTemp.Open
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
'
 fgExport.AddItem "" & vbTab & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ItemCode") & _
         vbTab & rsTemp("ItemName") & vbTab & rsTemp("ItemQty") & vbTab & rsTemp("ItemPrice") & vbTab & rsTemp("Tips") & vbTab & rsTemp("NoDiscount") & vbTab & rsTemp("NoVAT") & vbTab & rsTemp("ItemGroup") & vbTab & rsTemp("ItemCatagory")
        rsTemp.MoveNext
        Wend
        
        If fgExport.Rows = 0 Then fgExport.AddItem ""
        On Error Resume Next
        If fgExport.Rows <= 1 Then Exit Sub
Dim i As Integer
Dim j As Integer
    For i = 0 To frmCashMemo.fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If frmCashMemo.fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next
        
GridCount fgExport
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &HFFC0C0
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyBack Then
txtSearch.text = ""
End If
End Sub

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Menu Group From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
     
     
'     ----------------
If fgExport.Rows = 1 Then fgExport.AddItem ""
        On Error Resume Next
        If fgExport.Rows <= 1 Then Exit Sub
Dim i As Integer
Dim j As Integer
    For i = 0 To frmCashMemo.fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If frmCashMemo.fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next
    
'    ---------------------
     
     
   
'Unload Me

Set frmItemGroupSearch = Nothing

End Sub

Private Sub fgExport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case Col
   Case 3, 4, 6

      Cancel = True
End Select
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     fgExport.Editable = flexEDKbdMouse
     fgExport.ColDataType(1) = flexDTBoolean
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
'            rsTemp.CursorLocation = adUseClient
     rsTemp.Open "SELECT TOP 50 SerialNo,ItemCode,ItemName,ItemQty,ItemPrice,Tips,NoDiscount,NoVAT,ItemGroup,ItemCatagory FROM tblItemDetail", cn, adOpenStatic, adLockReadOnly
        
'   End If
   
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ItemCode") & _
         vbTab & rsTemp("ItemName") & vbTab & rsTemp("ItemQty") & vbTab & rsTemp("ItemPrice") & _
         vbTab & rsTemp("Tips") & vbTab & rsTemp("NoDiscount") & _
         vbTab & rsTemp("NoVAT") & vbTab & rsTemp("ItemGroup") & vbTab & rsTemp("ItemCatagory")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
'     If fgExport.Rows = 1 Then fgExport.AddItem ""
'--------------------------------------------------------
If fgExport.Rows = 1 Then fgExport.AddItem ""
        On Error Resume Next
        If fgExport.Rows <= 1 Then Exit Sub
Dim i As Integer
Dim j As Integer
    For i = 0 To frmCashMemo.fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If frmCashMemo.fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next


'---------------------------------------------------------


End Sub

Private Sub PopulateCompanySearch()

Dim iRows As Integer
Dim i As Integer
Dim temp As Double

temp = 0
    iRows = fgExport.Rows
    If fgExport.Rows <= 1 Then Exit Sub
    
    For i = 0 To iRows - 1
If fgExport.Cell(flexcpChecked, i, 1) = flexChecked Then
frmCashMemo.fgCashMemo.AddItem "" & vbTab & vbTab & fgExport.TextMatrix(i, 2) & vbTab & fgExport.TextMatrix(i, 3) & _
        vbTab & fgExport.TextMatrix(i, 4) & vbTab & fgExport.TextMatrix(i, 5) & vbTab & fgExport.TextMatrix(i, 6) & _
        vbTab & fgExport.TextMatrix(i, 7) & vbTab & fgExport.TextMatrix(i, 10) & vbTab & fgExport.TextMatrix(i, 11) & _
        vbTab & vbTab & vbTab & vbTab & fgExport.TextMatrix(i, 8) & vbTab & fgExport.TextMatrix(i, 9)
                     
        End If
        

    Next

    End Sub

Private Sub Text1_Change()
cmdFind_Click
End Sub

Private Sub txtSearch_Change()
cmdFind_Click
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If KeyAscii = 13 And txtSearch = "" Then
SendKeys Chr(9)
Call cmdOk_Click
End If
   
   If txtSearch <> "" Then
   
   Dim iRows As Integer
   Dim i As Integer
   Dim temp As Double

    temp = 0
    iRows = fgExport.Rows
    If fgExport.Rows <= 1 Then Exit Sub
     
    For i = 1 To iRows - 1
fgExport.Cell(flexcpChecked, i, 1) = flexChecked
frmCashMemo.fgCashMemo.AddItem "" & vbTab & vbTab & fgExport.TextMatrix(i, 2) & vbTab & fgExport.TextMatrix(i, 3) & _
        vbTab & fgExport.TextMatrix(i, 4) & vbTab & fgExport.TextMatrix(i, 5) & vbTab & fgExport.TextMatrix(i, 6) & vbTab & fgExport.TextMatrix(i, 7) & _
        vbTab & fgExport.TextMatrix(i, 10) & vbTab & fgExport.TextMatrix(i, 11) & _
        vbTab & vbTab & vbTab & vbTab & fgExport.TextMatrix(i, 8) & vbTab & fgExport.TextMatrix(i, 9)
'                   temp = temp + Val(fgExport.TextMatrix(i, 4))
txtSearch.text = ""

Next
End If
End If
Call deleteRow
'txtSearch.text = ""
End Sub

Private Sub deleteRow()
If fgExport.Rows = 1 Then fgExport.AddItem ""
        On Error Resume Next
        If fgExport.Rows <= 1 Then Exit Sub
Dim i As Integer
Dim j As Integer
    For i = 0 To frmCashMemo.fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If frmCashMemo.fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next
End Sub






