VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmMenuGroupSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Menugroup Search"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   Icon            =   "frmMenuGroupSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   4080
      Picture         =   "frmMenuGroupSearch.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   750
      Left            =   6255
      Picture         =   "frmMenuGroupSearch.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   5175
      Picture         =   "frmMenuGroupSearch.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1100
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   2610
      Left            =   240
      TabIndex        =   3
      Top             =   960
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMenuGroupSearch.frx":1FE8
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
      Caption         =   "Group  Details"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7215
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Menu Group Search"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7545
   End
End
Attribute VB_Name = "frmMenuGroupSearch"
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

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a MenuGroup From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
    
     Unload Me
     Set frmMenuGroupSearch = Nothing
End Sub



Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
       
       If strCallingForm = LCase("frmMenuGroup") Then
 '     rsTemp.CursorLocation = adUseClient
             rsTemp.Open "SELECT DISTINCT MGName,MG_ID FROM MenuGroup", cn, adOpenStatic, adLockReadOnly
     End If
   
         fgExport.Rows = 1

        If strCallingForm = LCase("frmMenuGroup") Then
             Label58.Caption = "MENU SEARCH"
              While Not rsTemp.EOF
            fgExport.AddItem "" & vbTab & rsTemp("MG_ID") & vbTab & rsTemp("MGName")
            rsTemp.MoveNext
        Wend
        GridCount fgExport
     End If
    
'     If fgExport.Rows = 1 Then fgExport.AddItem ""


End Sub



    Private Sub PopulateCompanySearch()
        
'      If strCallingForm = LCase("frmMeasurement") Then
''          If fgExport.Row > 0 Then
'          If fgExport.Row > 0 Then
'             frmMeasurement.PopulateMenuGroup fgExport.TextMatrix(fgExport.Row, 3)
'          End If
'      End If
    
    If strCallingForm = LCase("frmMenuGroup") Then
'        If fgExport.Row > 0 Then
        If fgExport.Row > 0 Then
            frmMenuGroup.PopulateMenuGroup fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End If
   
   End Sub






