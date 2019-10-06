VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReservation 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Booking & Reservation System"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmReservation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkActive 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Active"
      Height          =   195
      Left            =   9000
      TabIndex        =   43
      Top             =   10560
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton cmdActive 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Active"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9720
      Picture         =   "frmReservation.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9600
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12360
      Top             =   10320
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Driver={SQL Server};Server=NOTEBOOK;Database=RMS;Trusted_Connection=yes"""
      OLEDBString     =   "Driver={SQL Server};Server=NOTEBOOK;Database=RMS;Trusted_Connection=yes"""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Record Search"
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
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cha&nge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   10800
      MouseIcon       =   "frmReservation.frx":08D6
      Picture         =   "frmReservation.frx":0FC0
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   10800
      Top             =   10440
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      Picture         =   "frmReservation.frx":16AA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1200
      Picture         =   "frmReservation.frx":1F74
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2280
      Picture         =   "frmReservation.frx":283E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7560
      Picture         =   "frmReservation.frx":3108
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5520
      Picture         =   "frmReservation.frx":39D2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4440
      Picture         =   "frmReservation.frx":429C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3360
      Picture         =   "frmReservation.frx":4B66
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9600
      Width           =   1110
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6600
      Picture         =   "frmReservation.frx":5430
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9600
      Width           =   990
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "P&ost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8640
      MouseIcon       =   "frmReservation.frx":5CFA
      Picture         =   "frmReservation.frx":63E4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Reservation Details Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   15015
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   6045
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   14760
         _cx             =   26035
         _cy             =   10663
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReservation.frx":6ACE
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Reserbation Master"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   15015
      Begin VB.TextBox txtCActive 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   13440
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   42
         Text            =   "frmReservation.frx":6BDF
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtAdvance 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11760
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtCPost 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtUName 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13440
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   13440
         TabIndex        =   31
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtMenuItem 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmReservation.frx":6BE1
         Top             =   2160
         Width           =   10815
      End
      Begin VB.TextBox txtTAmount 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8880
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCPerson 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5520
         TabIndex        =   0
         Text            =   " "
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPerson 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7320
         TabIndex        =   1
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtHostAddress 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmReservation.frx":6BE3
         Top             =   1320
         Width           =   12735
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmReservation.frx":6BE5
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker BillDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   20774915
         CurrentDate     =   39739
      End
      Begin MSComCtl2.DTPicker PDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   20774915
         CurrentDate     =   39739
      End
      Begin VB.Label lblAdvance 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Advance Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11760
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13440
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMenuItem 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Menu Item Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblTAmount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblPartyDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Party Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblBillNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Bill No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblBookingDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Booking Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblGuest 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Person"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblContactPerson 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblVat 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8880
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblHost 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Reservation Person Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1080
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Adodc1                As ADODB.Recordset
Private rsfactory             As ADODB.Recordset
Private strFileName           As String
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim Tracer                    As Integer

Dim str As String
'--------------------------------------------------------------
Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition


Private objReportSub                        As CRPEAuto.Report 'sub
Private objReportDatabaseSub                As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition
Private rsDailyRpt                          As ADODB.Recordset


Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions

Private Sub cmdActive_Click()
cmdActive.Caption = "&Void"
If txtCActive.text = "Active" Then
If txtCPost.text = "Posted" Then
  If cmdActive.Caption = "&Void" Then
      If IsValidRecord Then
              If rcupdate Then
'                cmdPrint.Caption = "&Printing"
                cmdEdit.Enabled = False
                cmdCancel.Enabled = True
                cmdClose.Enabled = True
'                fgCashMemo.Enabled = False
                cmdFind.Enabled = True
                cmdPreview.Enabled = True
                cmdDelete.Enabled = True
                cmdChange.Enabled = False
                txtBillNo.Enabled = False
'                fgCashMemo.Editable = flexEDKbdMouse
'                txtCActive.text = "Active"
'                Call alldisable
'                s = " & Val(cmbATable.Columns(1).text) & "
'                rsAMaster.Requery
'                rsAMaster.MoveFirst
'                rsAMaster.Find "Post='" & parseQuotes(s) & "'"
'                FindRecord1
            End If
        End If
    End If
  End If
End If
    
    cmdActive.Caption = "&Active"
End Sub

Private Sub cmdChange_Click()
If cmdChange.Caption = "Cha&nge" Then
        cmdNew.Enabled = False
        Call allenable
        cmdChange.Caption = "&Modify"
        cmdCancel.Enabled = True
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
        cmdFind.Enabled = False
        cmdPreview.Enabled = False
        cmdClose.Enabled = False
        cmdPrint.Enabled = False
        txtBillNo.Enabled = False
        cmdPost.Enabled = False
'      End If
 
ElseIf cmdChange.Caption = "&Modify" Then
'  Call Calculation
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
'                 CmdChange.Enabled = False
                 cmdFind.Enabled = True
                 cmdPreview.Enabled = True
                 cmdDelete.Enabled = True
                 cmdPrint.Enabled = True
                 txtBillNo.Enabled = False
                Call alldisable
                rsfactory.Requery
'                Dim s As String
'                s = txtBillNo
'                rsfactory.MoveFirst
'                rsfactory.Find "BillNo='" & parseQuotes(s) & "'"
                FindRecord
            End If
        End If
    End If
    
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
     idelete = MsgBox("Do you want to delete this record?", vbYesNo)
     If txtUName.text = "Borhan" Then
    If idelete = vbYes Then
  
            cn.Execute "Delete From tblReservation Where BillNo ='" & parseQuotes(txtBillNo) & "'"
            Call allClear
    
    MsgBox "Please Call your System Administrator"
    End If
        
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdFirst_Click()
'Dim Adodc1 As ADODB.Recordset
'Adodc1.Recordset.MoveFirst
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
        txtBillNo = Adodc1.Recordset!BillNo
        BillDate.Value = Adodc1.Recordset!BDate
        PDate.Value = Adodc1.Recordset!PDate
        txtPerson = Adodc1.Recordset!Person
        txtCPerson = Adodc1.Recordset!CPerson
        txtRate = Adodc1.Recordset!Amount
        txtTAmount = Adodc1.Recordset!TAmount
        txtAdvance = Adodc1.Recordset!Advance
        txtTime = Adodc1.Recordset!strTime
        txtHostAddress = Adodc1.Recordset!HostName
        txtMenuItem = Adodc1.Recordset!MenuItem
        txtUName = Adodc1.Recordset!UName
        txtCPost.text = Adodc1.Recordset!Posted
End If
End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
       txtBillNo = Adodc1.Recordset!BillNo
        BillDate.Value = Adodc1.Recordset!BDate
        PDate.Value = Adodc1.Recordset!PDate
        txtPerson = Adodc1.Recordset!Person
        txtCPerson = Adodc1.Recordset!CPerson
        txtRate = Adodc1.Recordset!Amount
        txtTAmount = Adodc1.Recordset!TAmount
        txtAdvance = Adodc1.Recordset!Advance
        txtTime = Adodc1.Recordset!strTime
        txtHostAddress = Adodc1.Recordset!HostName
        txtMenuItem = Adodc1.Recordset!MenuItem
        txtUName = Adodc1.Recordset!UName
        txtCPost.text = Adodc1.Recordset!Posted
        
End If

End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
       txtBillNo = Adodc1.Recordset!BillNo
        BillDate.Value = Adodc1.Recordset!BDate
        PDate.Value = Adodc1.Recordset!PDate
        txtPerson = Adodc1.Recordset!Person
        txtCPerson = Adodc1.Recordset!CPerson
        txtRate = Adodc1.Recordset!Amount
        txtTAmount = Adodc1.Recordset!TAmount
        txtAdvance = Adodc1.Recordset!Advance
        txtTime = Adodc1.Recordset!strTime
        txtHostAddress = Adodc1.Recordset!HostName
        txtMenuItem = Adodc1.Recordset!MenuItem
        txtUName = Adodc1.Recordset!UName
        txtCPost.text = Adodc1.Recordset!Posted
        
End If

End Sub

Private Sub cmdPost_Click()

Dim s As String
cmdPost.Caption = "&Posted"

If cmdPost.Caption = "&Posted" Then
     If txtCPost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 cmdChange.Enabled = False
                 cmdFind.Enabled = True
                 cmdPreview.Enabled = True
                 cmdDelete.Enabled = True
                 cmdPrint.Enabled = True
                 txtBillNo.Enabled = False
                 Call alldisable
           End If
        End If
      End If
Else
 End If
cmdPost.Caption = "&Posted"
End Sub

Private Sub cmdPreview_Click()
Tracer = 0
If txtCPost.text = "Posted" Then
   Call printReport
   Else
   MsgBox "Please Post your Reservation Bill"
   End If
'RptReservation.Show vbModal
End Sub

Public Sub printReport()
On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Set rsDailyRpt = New ADODB.Recordset
    

If rsfactory.State <> 0 Then rsfactory.Close
        rsfactory.Open "select BillNo,BDate,PDate,CPerson,Person,Amount,TAmount,Advance,HostName,MenuItem,strTime, " & _
                          "UName from tblReservation", cn, adOpenStatic, adLockReadOnly
    
    If rsfactory.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Reservation.RPT"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        
        If rsDailyRpt.State <> 0 Then rsDailyRpt.Close

           
          rsDailyRpt.Open "SELECT tblReservation.BillNo,tblReservation.BDate,tblReservation.PDate," & _
                          "tblReservation.CPerson,tblReservation.Person," & _
                          "tblReservation.Amount,tblReservation.TAmount,tblReservation.Advance," & _
                          "tblReservation.HostName,tblReservation.MenuItem, tblReservation.strTime," & _
                          "tblReservation.UName FROM  tblReservation Where tblReservation.BillNo='" & txtBillNo.text & "' order by tblReservation.PDate", cn, adOpenStatic
                    
        
        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
            objReportFF.text = "'" + parseQuotes(txtUName.text) + " '"

        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        If Tracer = 0 Then
        objReport.Preview "Reservation Report", , , , , 16777216 Or 524288 Or 65536
        Else
        objReport.PrintOut
        End If
        
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Printing Cancel Information"
    End Select

End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdPrevious.Enabled = False
 Else
      cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
       txtBillNo = Adodc1.Recordset!BillNo
        BillDate.Value = Adodc1.Recordset!BDate
        PDate.Value = Adodc1.Recordset!PDate
        txtPerson = Adodc1.Recordset!Person
        txtCPerson = Adodc1.Recordset!CPerson
        txtRate = Adodc1.Recordset!Amount
        txtTAmount = Adodc1.Recordset!TAmount
        txtAdvance = Adodc1.Recordset!Advance
        txtTime = Adodc1.Recordset!strTime
        txtHostAddress = Adodc1.Recordset!HostName
        txtMenuItem = Adodc1.Recordset!MenuItem
        txtUName = Adodc1.Recordset!UName
        txtCPost.text = Adodc1.Recordset!Posted
        
End If

End Sub

Private Sub cmdPrint_Click()
Dim s As String
If cmdPrint.Caption = "&Print" Then
cmdPrint.Caption = "&Print"
        If IsValidRecord Then
            If rcupdate Then
'                cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 cmdChange.Enabled = False
                 cmdFind.Enabled = True
                 cmdPreview.Enabled = True
                 cmdDelete.Enabled = True
                 cmdPrint.Enabled = True
                txtBillNo.Enabled = False
                Call alldisable
'                txtWords = InWords(txtNPayable.text)

            End If
        End If
    End If
    
Tracer = 1
Screen.MousePointer = vbHourglass
If txtCPost.text = "Posted" Then
Call printReport
End If
Screen.MousePointer = vbDefault


cmdPrint.Caption = "&Print"
End Sub


Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from tblReservation", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If

    If Not rsfactory.EOF Then FindRecord

    txtBillNo.Enabled = False
    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
    
    Adodc1.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "tblReservation"

  Adodc1.Refresh

End Sub


Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdNew.Caption = "&New"
    cmdEdit.Caption = "&Edit"
    cmdPreview.Enabled = True
    cmdPrint.Enabled = True
    cmdDelete.Enabled = True
    cmdFind.Enabled = True
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdPost.Enabled = True
    cmdChange.Enabled = True
    txtBillNo.Enabled = False
'    Call allClear
    Call alldisable
    If Not rsfactory.BOF Then FindRecord

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdEdit_Click()
        
If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtCPerson.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdDelete.Enabled = False
        cmdPreview.Enabled = False
        cmdFind.Enabled = False
        cmdChange.Enabled = False
        cmdPost.Enabled = False
        cmdPrint.Enabled = False
        txtCPerson.SetFocus
        
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdDelete.Enabled = True
                cmdPreview.Enabled = True
                cmdFind.Enabled = True
                Call alldisable
                rsfactory.Requery

            End If
        End If
    End If

End Sub

Private Sub allenable()
    txtHostAddress.Enabled = True
    txtPerson.Enabled = True
    txtMenuItem.Enabled = True
    txtCPerson.Enabled = True
    txtPerson.Enabled = True
    txtRate.Enabled = True
    txtTAmount.Enabled = True
    txtAdvance.Enabled = True
    txtTime.Enabled = True
    BillDate.Enabled = True
    PDate.Enabled = True
End Sub

Private Sub alldisable()
    txtHostAddress.Enabled = False
    txtMenuItem.Enabled = False
    txtPerson.Enabled = False
    txtRate.Enabled = False
    txtTAmount.Enabled = False
    txtAdvance.Enabled = False
    txtTime.Enabled = False
    BillDate.Enabled = False
    txtCPerson.Enabled = False
    PDate.Enabled = False
End Sub

Private Sub allClear()
    txtHostAddress.text = ""
    txtMenuItem.text = ""
    txtPerson.text = ""
    txtCPerson.text = ""
    txtRate.text = ""
    txtTAmount.text = ""
    txtAdvance.text = ""
    BillDate.Value = Date
    PDate.Value = Date
    End Sub

Private Sub cmdNew_Click()
      
Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdDelete.Enabled = False
        cmdFind.Enabled = False
        cmdPreview.Enabled = False
        cmdPost.Enabled = False
        cmdChange.Enabled = False
        txtUName.text = frmLogin.txtUID.text
        txtCPost.text = "Not Posted"
        Call allClear

If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(BillNo),0) as InvNo from tblReservation"
            rs.Open str, cn, adOpenStatic, adLockReadOnly
                txtBillNo.text = Val(rs!InvNo) + 1

        Call allenable
            txtCPerson.SetFocus
        
    ElseIf cmdNew.Caption = "&Save" Then
        If IsValidRecord Then
            If rcupdate Then
                txtBillNo.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdDelete.Enabled = True
                cmdFind.Enabled = True
                cmdPreview.Enabled = True
                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
                cmdPost.Enabled = True
                Call alldisable
                
                FindRecord

            End If
        End If
    End If
 
ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Sub Timer1_Timer()
    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub

Public Sub FindRecord()
On Error Resume Next
If Not rsfactory.EOF Then
        txtBillNo = rsfactory("BillNo")
        BillDate = rsfactory("BDate")
        PDate = rsfactory("PDate")
        txtCPerson = rsfactory("CPerson")
        txtHostAddress = rsfactory("HostName")
        txtMenuItem = rsfactory("MenuItem")
        txtPerson = rsfactory("Person")
        txtTime = rsfactory("strTime")
        txtUName = rsfactory("UName")
        txtCPost = rsfactory("Posted") & ""
        txtRate = IIf(IsNull(rsfactory("Amount")), "", rsfactory("Amount"))
        txtTAmount = IIf(IsNull(rsfactory("TAmount")), "", rsfactory("TAmount"))
        txtAdvance = IIf(IsNull(rsfactory("Advance")), "", rsfactory("Advance"))
    End If
End Sub

Private Function rcupdate() As Boolean
On Error Resume Next

Dim ipost
Dim iprint

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then
    
    cn.Execute "INSERT INTO tblReservation(BillNo,BDate,PDate,CPerson,Person,Amount,TAmount,Advance,HostName,MenuItem,strTime, " & _
                   " Posted,UName) " & _
                   " VALUES ('" & txtBillNo & "','" & Format(BillDate, "dd-MMM-yyyy") & "','" & Format(PDate, "dd-MMM-yyyy") & "','" & parseQuotes(txtCPerson) & "'," & _
                   " " & Val(txtPerson.text) & "," & Val(txtRate.text) & "," & Val(txtTAmount.text) & "," & Val(txtAdvance.text) & "," & _
                   " '" & parseQuotes(txtHostAddress) & "','" & parseQuotes(txtMenuItem) & "','" & txtTime.text & "','" & txtCPost.text & "','" & txtUName.text & "') "

          rcupdate = True
          cn.CommitTrans
          MsgBox "Record Added", vbInformation, "Confirmation"

ElseIf (cmdEdit.Caption = "&Update") Then

cn.Execute "Update tblReservation SET BDate='" & Format(BillDate, "dd-mmm-yyyy") & "',PDate='" & Format(PDate, "dd-MMM-yyyy") & "',CPerson='" & parseQuotes(txtCPerson) & "', " & _
                   " Person=" & Val(txtPerson.text) & ",Amount=" & Val(txtRate.text) & ",TAmount=" & Val(txtTAmount.text) & ",Advance=" & Val(txtAdvance.text) & ",HostName='" & parseQuotes(txtHostAddress.text) & "', " & _
                   " MenuItem='" & parseQuotes(txtMenuItem) & "',strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE BillNo = '" & txtBillNo & "'"


        rcupdate = True
        cn.CommitTrans
        MsgBox "Record Updated", vbInformation, "Confirmation"
        
'----------------------------------------------Printing Start--------------------------
  ElseIf cmdPrint.Caption = "&Printing" Then
  
    txtCPost.text = "Posted"

    iprint = MsgBox("Do you want to Print this Reservation?", vbYesNo)

    If iprint = vbYes Then
                 
         cn.Execute "Update tblReservation SET BDate='" & Format(BillDate, "dd-mmm-yyyy") & "',PDate='" & Format(PDate, "dd-MMM-yyyy") & "',CPerson='" & parseQuotes(txtCPerson) & "', " & _
                   " Person=" & Val(txtPerson.text) & ",Amount=" & Val(txtRate.text) & ",TAmount=" & Val(txtTAmount.text) & ",Advance=" & Val(txtAdvance.text) & ",HostName='" & parseQuotes(txtHostAddress.text) & "', " & _
                   " MenuItem='" & parseQuotes(txtMenuItem) & "',strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE BillNo = '" & txtBillNo & "'"
'        Next
                
        rcupdate = True
        cn.CommitTrans
        
        End If
'----------------------------------Printing End---------------------------

'----------------------------------Posted Start--------------------------
    ElseIf cmdPost.Caption = "&Posted" Then

     txtCPost.text = "Posted"
     
     ipost = MsgBox("Do you want to Post this bill?", vbYesNo)

           If ipost = vbYes Then

     cn.Execute "Update tblReservation SET BDate='" & Format(BillDate, "dd-mmm-yyyy") & "',PDate='" & Format(PDate, "dd-MMM-yyyy") & "',CPerson='" & parseQuotes(txtCPerson) & "', " & _
                   " Person=" & Val(txtPerson.text) & ",Amount=" & Val(txtRate.text) & ",TAmount=" & Val(txtTAmount.text) & ",Advance=" & Val(txtAdvance.text) & ",HostName='" & parseQuotes(txtHostAddress.text) & "', " & _
                   " MenuItem='" & parseQuotes(txtMenuItem) & "',strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE BillNo = '" & txtBillNo & "'"

       rcupdate = True
       cn.CommitTrans
       MsgBox "Record Posted Successfully", vbInformation, "Confirmation"

    End If
    End If

'    cn.CommitTrans
'    Exit Sub
    Exit Function

End Function


Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If (txtPerson.text = "") Then
       MsgBox "Enter Person Quantity"
       txtPerson.SetFocus
       IsValidRecord = False
       Exit Function
    End If
    
If (txtCPerson.text = "") Then
       MsgBox "Enter Contact Person"
       txtCPerson.SetFocus
       IsValidRecord = False
       Exit Function
    End If
    
  If (txtHostAddress.text = "") Then
       MsgBox "Enter Person Details"
       txtHostAddress.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtMenuItem.text = "") Then
      MsgBox "Enter Menu Informations"
      txtMenuItem.SetFocus
      IsValidRecord = False
      Exit Function
    End If
        
    End Function

Private Sub txtAdvance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtHostAddress.SetFocus
    End If
End Sub

Private Sub txtCPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtPerson.SetFocus
    End If
End Sub

Private Sub txtHostAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtMenuItem.SetFocus
    End If
End Sub

Private Sub txtMenuItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    cmdNew.SetFocus
    End If
End Sub

Private Sub txtRate_Change()
txtTAmount = Val(txtPerson) * Val(txtRate)
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtAdvance.SetFocus
    End If
End Sub

Private Sub txtPerson_Change()
txtTAmount = Val(txtPerson) * Val(txtRate)
End Sub

Private Sub txtPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtRate.SetFocus
    End If
End Sub


