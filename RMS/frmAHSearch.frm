VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAHSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Accounts Head Details"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   Icon            =   "frmAHSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMode 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   " "
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   7440
      Width           =   1935
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
      Height          =   630
      Left            =   6960
      Picture         =   "frmAHSearch.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   975
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
      Height          =   630
      Left            =   6000
      Picture         =   "frmAHSearch.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   975
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
      Height          =   630
      Left            =   4920
      Picture         =   "frmAHSearch.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Accounts Head Name Details"
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
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   6570
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7785
         _cx             =   13732
         _cy             =   11589
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAHSearch.frx":1A6A
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
      TabIndex        =   8
      Top             =   7200
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
      Left            =   2160
      TabIndex        =   7
      Top             =   7200
      Width           =   1815
   End
End
Attribute VB_Name = "frmAHSearch"
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
        
If cboMode.text = "Guest Card No" Then

      rsTemp.Open "SELECT CID,AHName,Department,AHType,Cphone " & _
                 "FROM AccountsHead WHERE AccountsHead.AHName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
ElseIf cboMode.text = "Guest Name" Then

      rsTemp.Open "SELECT CID,AHName,Department,AHType,Cphone " & _
                 "FROM AccountsHead WHERE AccountsHead.Department Like '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

'Else cboMode.text = "Guest Phone No" Then
Else

      rsTemp.Open "SELECT CID,AHName,Department,AHType,Cphone " & _
                 "FROM AccountsHead WHERE AccountsHead.CPhone LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

End If

fgExport.Rows = 1

    While Not rsTemp.EOF
    
    fgExport.AddItem "" & vbTab & rsTemp("CID") & vbTab & rsTemp("AHName") & _
         vbTab & rsTemp("Department") & vbTab & rsTemp("AHType") & vbTab & rsTemp("Cphone")
   
'        fgExport.AddItem "" & vbTab & rsTemp("CID") & vbTab & rsTemp("AHName")& vbTab & rsTemp("Department")& vbTab & rsTemp("C")

        rsTemp.MoveNext
    Wend
'     GridCount fgExport
End Sub

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Customer Name From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmAccountsHead = Nothing
End Sub

Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
     rsTemp.Open "SELECT AID,AHName,Department,AHType FROM AccountsHead", cn, adOpenStatic, adLockReadOnly
        
            fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("AID") & vbTab & rsTemp("AHName") & _
         vbTab & rsTemp("Department") & vbTab & rsTemp("AHType")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
     
       cboMode.AddItem "Accounts Head"
       cboMode.AddItem "Department Name"
       cboMode.AddItem "Accounts Type"
     
End Sub

    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmAccountsHead.PopulateCnf fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub txtSearch_Change()
 cmdFind_Click
End Sub





