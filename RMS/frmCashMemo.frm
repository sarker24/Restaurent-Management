VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "ANIGIF.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCashMemo 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Restaurant Billing Systems [LOTUS ETANG]"
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19110
   FillColor       =   &H00008080&
   ForeColor       =   &H00000000&
   Icon            =   "frmCashMemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10905
   ScaleWidth      =   19110
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Bill Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2805
      Left            =   240
      TabIndex        =   84
      Top             =   7080
      Width           =   7305
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   5160
         TabIndex        =   94
         Text            =   " "
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtTotalDiscount 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtNPayable 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtTotalSCharge 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtTotalVat 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtTotalBill 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dcReceiveBy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtDue 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   88
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtPaid 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   87
         Text            =   " "
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cboMode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5160
         TabIndex        =   86
         Text            =   " "
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtTItem 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Text            =   " "
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Bill "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblTSCharge 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Service Charge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblTVat 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total VAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblPaid 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   101
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblDue 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Due"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   100
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Net Payable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblPMode 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Payment Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   98
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblRemark 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   97
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblTRCount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   95
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtNoVAT 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   12000
      TabIndex        =   83
      TabStop         =   0   'False
      Text            =   "No VAT"
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9165
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtItemName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7695
   End
   Begin VB.TextBox txtItemCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10080
      TabIndex        =   76
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtTips 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11400
      TabIndex        =   75
      TabStop         =   0   'False
      Text            =   "Tips"
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtNoDiscount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   12000
      TabIndex        =   74
      TabStop         =   0   'False
      Text            =   "No Discount"
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtItemGroup 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   12960
      TabIndex        =   73
      TabStop         =   0   'False
      Text            =   "Item Group"
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtItemCatagory 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   13920
      TabIndex        =   72
      TabStop         =   0   'False
      Text            =   "Item Catagory"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Find Last"
      Top             =   10200
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Find Next"
      Top             =   10200
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Find Previous"
      Top             =   10200
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Find First"
      Top             =   10200
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0B4A9&
      Height          =   2805
      Left            =   7560
      TabIndex        =   39
      Top             =   7080
      Width           =   3735
      Begin MSComCtl2.DTPicker AdDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   56
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65142787
         CurrentDate     =   41925
      End
      Begin VB.TextBox txtCashReceived 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   46
         Text            =   " "
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtCashBack 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   45
         Text            =   " "
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtMR_No 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   42
         Text            =   " "
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtAdvance 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """BDT""#,##0.00;(""BDT""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   40
         Text            =   " "
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCashPaid 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cash Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblCashBack 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cash Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblAdvance_Date 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Advance Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblMR_No 
         BackColor       =   &H00C0B4A9&
         Caption         =   "MR.No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblAdvance 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Advance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1455
      End
   End
   Begin AniGIFCtrl.AniGIF AniGIF1 
      Height          =   2055
      Left            =   12360
      TabIndex        =   34
      Top             =   7080
      Width           =   2775
      BackColor       =   12632256
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      GIF             =   "frmCashMemo.frx":058A
      ExtendWidth     =   4895
      ExtendHeight    =   3625
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0B4A9&
      Height          =   4095
      Left            =   240
      TabIndex        =   33
      Top             =   2880
      Width           =   11055
      Begin VSFlex7LCtl.VSFlexGrid fgCashMemo 
         Height          =   3735
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   10815
         _cx             =   19076
         _cy             =   6588
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   8421376
         ForeColor       =   -2147483640
         BackColorFixed  =   16744576
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16512
         ForeColorSel    =   -2147483638
         BackColorBkg    =   12629161
         BackColorAlternate=   8421504
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCashMemo.frx":53F2
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
   Begin VB.TextBox txtUName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   " "
      Top             =   9600
      Width           =   2775
   End
   Begin VB.CommandButton cmdPost2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "P&ost"
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
      Left            =   7680
      MouseIcon       =   "frmCashMemo.frx":55B4
      Picture         =   "frmCashMemo.frx":5C9E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9960
      Width           =   1095
   End
   Begin VB.TextBox txtWords 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   11520
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   " "
      Top             =   5760
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.CommandButton cmdActive 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Active"
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
      Left            =   8760
      Picture         =   "frmCashMemo.frx":6388
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Find"
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
      Left            =   5640
      Picture         =   "frmCashMemo.frx":6C52
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdLDelete 
      BackColor       =   &H00C0B4A9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12000
      Picture         =   "frmCashMemo.frx":751C
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Remove"
      Top             =   2550
      Width           =   540
   End
   Begin VB.CheckBox chkActive 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Active"
      Height          =   195
      Left            =   11520
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Change"
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
      Left            =   9840
      Picture         =   "frmCashMemo.frx":7AA6
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9960
      Width           =   1110
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
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
      Left            =   3480
      Picture         =   "frmCashMemo.frx":8370
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9960
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Print"
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
      Left            =   4560
      Picture         =   "frmCashMemo.frx":8C3A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9960
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Quit"
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
      Left            =   6600
      Picture         =   "frmCashMemo.frx":9504
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
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
      Left            =   2400
      Picture         =   "frmCashMemo.frx":9DCE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
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
      Left            =   1320
      Picture         =   "frmCashMemo.frx":A698
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
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
      Left            =   240
      Picture         =   "frmCashMemo.frx":AF62
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdtemSelected 
      BackColor       =   &H00C0B4A9&
      Height          =   405
      Left            =   11340
      Picture         =   "frmCashMemo.frx":B82C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Browse"
      Top             =   2550
      Width           =   555
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cash Master"
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
      Height          =   2295
      Left            =   240
      TabIndex        =   17
      Top             =   0
      Width           =   14775
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   13440
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   9000
         TabIndex        =   62
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtKOT 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11520
         TabIndex        =   61
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtBOT 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   12480
         TabIndex        =   60
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtVAT 
         Appearance      =   0  'Flat
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtServiceCharge 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7200
         TabIndex        =   58
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtPersonNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Text            =   " "
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDiscountAmt 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10200
         TabIndex        =   57
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Left            =   14280
         Top             =   840
      End
      Begin VB.TextBox txtCPost 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   12240
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "frmCashMemo.frx":BDB6
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtCActive 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   12240
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "frmCashMemo.frx":BDBA
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtCAddress 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   53
         Top             =   1850
         Width           =   10695
      End
      Begin VB.TextBox txtGName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   1440
         Width           =   10695
      End
      Begin VB.ComboBox cmbTName 
         Height          =   315
         Left            =   3000
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtBillSerialNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbDCard 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   10695
      End
      Begin VB.ComboBox cmbWaiter 
         Height          =   315
         Left            =   4680
         TabIndex        =   1
         Top             =   600
         Width           =   1695
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
         Left            =   1440
         TabIndex        =   31
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   65142787
         CurrentDate     =   39739
      End
      Begin VB.Label lblGuest 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Guest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   71
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSCharge 
         BackColor       =   &H00C0B4A9&
         Caption         =   "S. Charge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   70
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblVat 
         BackColor       =   &H00C0B4A9&
         Caption         =   "VAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   69
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblDiscount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Discount(%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13440
         TabIndex        =   67
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblKOT 
         BackColor       =   &H00C0B4A9&
         Caption         =   "KOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   66
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblBOT 
         BackColor       =   &H00C0B4A9&
         Caption         =   "BOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12480
         TabIndex        =   65
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDiscountamt 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Discount(Amt)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   64
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblDcard 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Discount card"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblWaiter 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Waiter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblTName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Table Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblBillNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Bill No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblGName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Guest Name"
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
         Left            =   120
         TabIndex        =   50
         Top             =   1440
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc dcCatagory 
      Height          =   360
      Left            =   11520
      Top             =   4800
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   635
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dcCatagory"
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
   Begin MSAdodcLib.Adodc DCRSearch 
      Height          =   330
      Left            =   11520
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
      Connect         =   "Driver={SQL Server};Server=MAS;Database=KG;Trusted_Connection=yes"
      OLEDBString     =   "Driver={SQL Server};Server=MAS;Database=KG;Trusted_Connection=yes"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblCashMaster"
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
   Begin MSAdodcLib.Adodc dcATable 
      Height          =   360
      Left            =   11520
      Top             =   5160
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   635
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Active Table"
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
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   1485
      Left            =   11520
      TabIndex        =   82
      Top             =   4080
      Visible         =   0   'False
      Width           =   4080
      _cx             =   7197
      _cy             =   2619
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
      FormatString    =   $"frmCashMemo.frx":BDBC
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
   Begin VB.Label lblRate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   81
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblQty 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9165
      TabIndex        =   80
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblItemName 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   79
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblItemCode 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Item Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   78
      Top             =   2280
      Width           =   1095
   End
   Begin MSForms.ComboBox cmbATable 
      Height          =   375
      Left            =   13080
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
      VariousPropertyBits=   746604571
      BackColor       =   -2147483645
      DisplayStyle    =   3
      Size            =   "3413;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Bill Created By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Left            =   12360
      TabIndex        =   32
      Top             =   9120
      Width           =   2775
   End
   Begin VB.Label Label63 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   11280
      TabIndex        =   24
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Active Table"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   13080
      TabIndex        =   16
      Top             =   3240
      Width           =   1935
   End
End
Attribute VB_Name = "frmCashMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                          As ADODB.Recordset
Private rs                              As New ADODB.Recordset
Private rscashmaster                    As New ADODB.Recordset
Private rsCashDetail                    As ADODB.Recordset
Private rsCashDetail1                   As ADODB.Recordset
Private rsATable                        As New ADODB.Recordset
Private rsAMaster                       As New ADODB.Recordset
Private rsCustomerMaster                As New ADODB.Recordset
Dim str                                 As String
Private rsfactory                       As ADODB.Recordset
Dim Tracer                              As Integer
Private rsTemp2                         As ADODB.Recordset
Dim flagSlNo                            As Integer
Dim strMood                             As String
''---------------------------------------------------------------------------
''----Add For Reporting Perpose----------------------------------------------
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
Private objReportFormulaFieldDefinationsSub As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                      As CRPEAuto.FormulaFieldDefinition


Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private rsDailyRpt                          As ADODB.Recordset
'Private Tracer                             As Integer
Private strGroupName                        As String
Dim temp As Double
Dim temp1 As Double
Dim temp2 As Double
Dim temp3 As Double

''--------------------------------------------------------------------------------

Private Sub cboMode_Click()
If txtPaid.text = "" Then

    ElseIf txtPaid.text > 0 Then
            cboMode.Enabled = False
    Else
            cboMode.Enabled = True
End If
End Sub


Private Sub cboMode_DropDown()
If txtPaid.text = "" Then

    ElseIf txtPaid.text > 0 Then
            cboMode.Enabled = False
    Else
            cboMode.Enabled = True
End If
End Sub

'================== Table Information ========================

Private Sub cmbTName_click()
If IsValidtATable Then
   
   rscashmaster.Requery
   End If
End Sub

Private Sub cmbTName_GotFocus()
'If cmbTName <> rsCashMaster!TableName Then
'MsgBox "Invalid Table Name.", vbInformation, "Confirmation"
'cmbTName.text = ""
'cmbTName.SetFocus
'End If
cmbTName.BackColor = &HFFFFC0
End Sub

Private Sub cmbTName_LostFocus()

If IsValidtATable Then
   
   rscashmaster.Requery
   End If
    cmbTName.BackColor = vbWhite
End Sub

Private Sub cmbTName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbTName, KeyAscii)
   If KeyAscii = 13 Then
   If IsValidtATable Then

   rscashmaster.Requery
   Else
       SendKeys Chr(9)
    End If
    End If
End Sub

Private Sub TableName()

    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT TableNo FROM tblRDetail ORDER BY TableNo ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbTName.AddItem rsTemp2("TableNo")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub



'================== Table Information End ==============

'============ Waiter Information =======================

Private Sub cmbWaiter_Change()
'If IsValidtTable Then
'
'   rsCashMaster.Requery
'   End If
'    cmbTName.BackColor = vbWhite
End Sub

Private Sub cmbWaiter_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbWaiter, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbWaiter_GotFocus()
cmbWaiter.BackColor = &HFFFFC0
End Sub

Private Sub cmbWaiter_LostFocus()
    cmbWaiter.BackColor = vbWhite
End Sub

Private Sub WaiterName()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT WaiterName FROM tblWaiterName ORDER BY WaiterName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbWaiter.AddItem rsTemp2("WaiterName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

'================== Waiter Information End ========================

Private Sub cmdPreview_Click()
    Tracer = 0
    If txtCPost.text = "Posted" Then
'   Call GuestCopy
   Call CashCopy
   
'   Call printReport

   Else
   MsgBox "Please Post your bill"
   End If
End Sub



Private Sub cmdAdd_Click(Index As Integer)
Dim i As Integer, j As Integer
Dim rs As New ADODB.Recordset
'-----------------------------------------
Select Case Index
Case 1
    fgCashMemo.AddItem ""
    fgCashMemo.Col = 1

    If flagSlNo = 0 Then
    rs.Open "Select SL=isnull(max(LedgerID),0) from tblCashDetail", cn, adOpenStatic
    j = rs!SL + 1
    fgCashMemo.TextMatrix(fgCashMemo.Rows - 1, 1) = j
    flagSlNo = 1
    Else
       fgCashMemo.TextMatrix(fgCashMemo.Rows - 1, 1) = fgCashMemo.TextMatrix(fgCashMemo.Rows - 2, 1) + 1
        j = j + 1
    End If
End Select
End Sub

Private Sub cmdCalculate_Click()
 Dim j As Integer
        temp = 0
        temp1 = 0
         For j = 1 To fgCashMemo.Rows - 1

        temp = temp + Val(fgCashMemo.TextMatrix(j, 6))
        temp1 = temp1 + Val(fgCashMemo.TextMatrix(j, 7))


         fgCashMemo.TextMatrix(j, 10) = temp1 - temp
   Next
End Sub


'----------------------- Restaurant Guest Related -------------------------------------------------------

Private Sub cmbDCard_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbDCard, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub CName()
    Dim rsTemp2 As New ADODB.Recordset
          
     rsTemp2.Open ("SELECT DISTINCT Dcard FROM RMSCustomer ORDER BY DCard ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbDCard.AddItem rsTemp2("DCard")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub

Private Sub cmbDCard_Click()
Set rsCustomerMaster = New ADODB.Recordset
    
    If rsCustomerMaster.State <> 0 Then rsCustomerMaster.Close
       rsCustomerMaster.Open "select CID,DCard,CName,CAddress,CPhone,BDate,CEmail,DiscountAmt,RPoint from RMSCustomer where DCard ='" & cmbDCard & "' ", cn, adOpenStatic, adLockReadOnly

   If rsCustomerMaster.RecordCount > 0 Then
      rsCustomerMaster.MoveFirst
    End If
      
    If Not rsCustomerMaster.EOF Then FindRecord2
End Sub

Private Sub cmbDCard_LostFocus()
Set rsCustomerMaster = New ADODB.Recordset
    
    If rsCustomerMaster.State <> 0 Then rsCustomerMaster.Close
       rsCustomerMaster.Open "select CID,DCard,CName,CAddress,CPhone,BDate,CEmail,DiscountAmt,RPoint from RMSCustomer where DCard ='" & cmbDCard & "' ", cn, adOpenStatic, adLockReadOnly

   If rsCustomerMaster.RecordCount > 0 Then
      rsCustomerMaster.MoveFirst
    End If
      
    If Not rsCustomerMaster.EOF Then FindRecord2
End Sub


Private Sub cmbDCard_DropDown()
cmbDCard.Refresh
End Sub

Private Sub FindRecord2()
    txtGName = rsCustomerMaster!CName
    txtCAddress = rsCustomerMaster!CAddress
    txtDiscount = rsCustomerMaster!DiscountAmt
End Sub

'--------------------------End Customer Informations------------------------------


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
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
'                CmdDelete.Enabled = True
                cmdChange.Enabled = False
                txtBillSerialNo.Enabled = False
                cmdtemSelected.Enabled = True
                cmdLDelete.Enabled = True
                fgCashMemo.Editable = flexEDKbdMouse
            
            End If
        End If
    End If
  End If
End If
    
    cmdActive.Caption = "&Active"
'    cmdChange.Enabled = True

End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*
' Routine:           cmbTName_InitColumnProps
' Description:       type_description_here
' Created by:        Administrator
' Machine:           MAS
' Date-Time:         8/2/200811:45:55 PM
' Last modification: last_modification_info_here
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*


Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdOpen.Enabled = True
    cmdClose.Enabled = True
    cmbATable.Enabled = True
    cmdtemSelected.Enabled = False
    cmdLDelete.Enabled = False
    cmdChange.Caption = "&Change"
    fgCashMemo.Enabled = False
    cmdPreview.Enabled = True
    cmdPost2.Enabled = True
    cmdActive.Enabled = True
    cmdPrint.Enabled = True
    cmdChange.Enabled = True
'    Call Clear
    Call alldisable
    
'    flagSlNo = 0
    cmdNew.Caption = "&New"
    If Not rscashmaster.EOF Then FindRecord

End Sub

Private Sub cmdChange_Click()
If cmdChange.Caption = "&Change" Then
     strMood = "U"
'    If txtPaid.text = "0" Then
        cmdNew.Enabled = False
        Call allenable
        cmdChange.Caption = "&Modify"
        cmdCancel.Enabled = True
        cmdEdit.Enabled = False
        cmbATable.Enabled = True
        cmdOpen.Enabled = False
        cmdPreview.Enabled = False
        cmdClose.Enabled = False
        cmdtemSelected.Enabled = True
        cmdLDelete.Enabled = True
        fgCashMemo.Enabled = True
        cmdActive.Enabled = False
        cmdPost2.Enabled = False
        cmdPrint.Enabled = False
        fgCashMemo.Editable = flexEDKbdMouse
        txtBillSerialNo.Enabled = False
        Call Calculation
'      End If
 
ElseIf cmdChange.Caption = "&Modify" Then
  Call Calculation
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Enabled = True
                cmdEdit.Enabled = True
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
                fgCashMemo.Enabled = True
                cmbATable.Enabled = True
                cmdtemSelected.Enabled = False
                cmdLDelete.Enabled = False
                cmdActive.Enabled = True
                cmdPost2.Enabled = True
                cmdPrint.Enabled = True
                cmdChange.Enabled = True
                cmdChange.Caption = "&Change"
                fgCashMemo.Editable = flexEDNone
                Call alldisable
                rscashmaster.Requery
            
            End If
        End If
    End If
'    If txtPaid.text <> "0" Then
'        MsgBox "Bill Can Not Change After Paid", vbInformation, "Confirmation"
'        End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
 If cmdEdit.Caption = "&Edit" Then
     strMood = "U"
    If txtCPost.text = "Not Posted" Then
        cmdNew.Enabled = False
        Call allenable
'        cmbTName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
'        CmdDelete.Enabled = False
        cmbATable.Enabled = True
        cmdOpen.Enabled = False
        cmdPreview.Enabled = False
        cmdClose.Enabled = False
        cmdtemSelected.Enabled = True
         cmdLDelete.Enabled = True
         fgCashMemo.Enabled = True
         cmdActive.Enabled = False
         cmdPost2.Enabled = False
         cmdPrint.Enabled = False
         cmdChange.Enabled = False
        fgCashMemo.Editable = flexEDKbdMouse
        txtBillSerialNo.Enabled = False
        Call Calculation
      End If

  ElseIf cmdEdit.Caption = "&Update" Then
  Call Calculation
'  Call ActiveTable
        
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
                cmdChange.Enabled = True
                fgCashMemo.Enabled = True
                cmbATable.Enabled = True
                cmdtemSelected.Enabled = False
                cmdLDelete.Enabled = False
                cmdActive.Enabled = True
                cmdPost2.Enabled = True
                cmdPrint.Enabled = True
                cmdChange.Enabled = False
                fgCashMemo.Editable = flexEDNone
                Call alldisable
                rscashmaster.Requery
'                Dim s As String
'                s = txtBillSerialNo
'                rsCashMaster.MoveFirst
'                rsCashMaster.Find "SerialNo='" & parseQuotes(s) & "'"
'                FindRecord11
            End If
        End If
End If

End Sub

Private Sub cmdDel_Click(Index As Integer)

Select Case Index
Case 0
    If fgCashMemo.Rows = 1 Then Exit Sub
    If fgCashMemo.Row >= 1 Then
        fgCashMemo.RemoveItem fgCashMemo.Row
    Else
        MsgBox "You have to select a row to delete.", vbInformation, "General"
    End If

End Select
End Sub


Private Sub cmdFind_Click(Index As Integer)
frmFind.Show vbModal
End Sub

Private Sub cmdFirst_Click()
Dim i As Integer
    Dim strLedgerDetail As String
    Set rsCashDetail = New ADODB.Recordset

DCRSearch.Recordset.MoveFirst
If DCRSearch.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    
    txtBillSerialNo = DCRSearch.Recordset!SerialNo
    BillDate = DCRSearch.Recordset!strDate
    cmbTName = DCRSearch.Recordset!TableName
    cmbWaiter = DCRSearch.Recordset!WaiterName
    txtPersonNo = DCRSearch.Recordset!Guest
    txtServiceCharge = DCRSearch.Recordset!ServiceCharge
    txtVAT = DCRSearch.Recordset!Vat
    txtDiscount = DCRSearch.Recordset!Discount
    txtDiscountAmt = DCRSearch.Recordset!DiscountAmt
    cmbDCard = DCRSearch.Recordset!DCard
    txtGName = DCRSearch.Recordset!GName
    txtCAddress = DCRSearch.Recordset!Address

    txtTotalBill = DCRSearch.Recordset!TotalBill
    txtTotalVat = DCRSearch.Recordset!TotalVat
    txtTotalSCharge = DCRSearch.Recordset!TSCharge
    txtTotalDiscount = DCRSearch.Recordset!TotalDiscount
    txtNPayable = DCRSearch.Recordset!NetPayable
    cboMode = DCRSearch.Recordset!PaymentMode
    txtRemarks = DCRSearch.Recordset!Remarks
    txtPaid = DCRSearch.Recordset!Paid
    txtDue = DCRSearch.Recordset!Due
    cmbATable = DCRSearch.Recordset!Post
    chkActive = DCRSearch.Recordset!Active
    txtCActive = DCRSearch.Recordset!CActive
    txtCPost = DCRSearch.Recordset!CPost
    txtUName = DCRSearch.Recordset!UName
    txtKOT = DCRSearch.Recordset!Kot
    txtBOT = DCRSearch.Recordset!Bot
    AdDate.Value = DCRSearch.Recordset!ADate
    txtAdvance = DCRSearch.Recordset!Advance
    txtMR_No = DCRSearch.Recordset!MrNo
 
          
        
        
        
        
    fgCashMemo.Rows = 1
    strLedgerDetail = "SELECT  BillSerialNo,SerialNo,ItemCode, ItemName, Qty, Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount" & _
                " FROM tblCashDetail " & _
                "WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo.text) & "'"
    rsCashDetail.CursorLocation = adUseClient
    rsCashDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsCashDetail.RecordCount <> 0 Then

        fgCashMemo.Rows = rsCashDetail.RecordCount + 1
                i = 0
        For i = 1 To rsCashDetail.RecordCount
            fgCashMemo.TextMatrix(i, 1) = rsCashDetail("BillSerialNo")
            fgCashMemo.TextMatrix(i, 2) = rsCashDetail("SerialNo")
            fgCashMemo.TextMatrix(i, 3) = rsCashDetail("ItemCode")
            fgCashMemo.TextMatrix(i, 4) = rsCashDetail("ItemName")
            fgCashMemo.TextMatrix(i, 5) = rsCashDetail("Qty")
            fgCashMemo.TextMatrix(i, 6) = rsCashDetail("Rate")
            fgCashMemo.TextMatrix(i, 7) = rsCashDetail("Tips")
            fgCashMemo.TextMatrix(i, 8) = rsCashDetail("ItemGroup")
            fgCashMemo.TextMatrix(i, 9) = rsCashDetail("ItemCatagory")
            fgCashMemo.TextMatrix(i, 10) = rsCashDetail("strDate")
            fgCashMemo.TextMatrix(i, 11) = rsCashDetail("CActive")
            fgCashMemo.TextMatrix(i, 12) = rsCashDetail("CPost")
            fgCashMemo.TextMatrix(i, 13) = rsCashDetail("NoDiscount")
            
    rsCashDetail.MoveNext
        Next
      End If
    rsCashDetail.Close

        
        
End If
End Sub

Private Sub cmdLast_Click()
Dim i As Integer
    Dim strLedgerDetail As String
    Set rsCashDetail = New ADODB.Recordset

DCRSearch.Recordset.MoveLast
If DCRSearch.Recordset.EOF = True Then
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    
    txtBillSerialNo = DCRSearch.Recordset!SerialNo
    BillDate = DCRSearch.Recordset!strDate
    cmbTName = DCRSearch.Recordset!TableName
    cmbWaiter = DCRSearch.Recordset!WaiterName
    txtPersonNo = DCRSearch.Recordset!Guest
    txtServiceCharge = DCRSearch.Recordset!ServiceCharge
    txtVAT = DCRSearch.Recordset!Vat
    txtDiscount = DCRSearch.Recordset!Discount
    txtDiscountAmt = DCRSearch.Recordset!DiscountAmt
    cmbDCard = DCRSearch.Recordset!DCard
    txtGName = DCRSearch.Recordset!GName
    txtCAddress = DCRSearch.Recordset!Address

    txtTotalBill = DCRSearch.Recordset!TotalBill
    txtTotalVat = DCRSearch.Recordset!TotalVat
    txtTotalSCharge = DCRSearch.Recordset!TSCharge
    txtTotalDiscount = DCRSearch.Recordset!TotalDiscount
    txtNPayable = DCRSearch.Recordset!NetPayable
    cboMode = DCRSearch.Recordset!PaymentMode
    txtRemarks = DCRSearch.Recordset!Remarks
    txtPaid = DCRSearch.Recordset!Paid
    txtDue = DCRSearch.Recordset!Due
    cmbATable = DCRSearch.Recordset!Post
    chkActive = DCRSearch.Recordset!Active
    txtCActive = DCRSearch.Recordset!CActive
    txtCPost = DCRSearch.Recordset!CPost
    txtUName = DCRSearch.Recordset!UName
    txtKOT = DCRSearch.Recordset!Kot
    txtBOT = DCRSearch.Recordset!Bot
    AdDate.Value = DCRSearch.Recordset!ADate
    txtAdvance = DCRSearch.Recordset!Advance
    txtMR_No = DCRSearch.Recordset!MrNo
          
        
        
        
        
        fgCashMemo.Rows = 1
    strLedgerDetail = "SELECT  BillSerialNo,SerialNo,ItemCode, ItemName, Qty, Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount" & _
                " FROM tblCashDetail " & _
                "WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo.text) & "'"
    rsCashDetail.CursorLocation = adUseClient
    rsCashDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsCashDetail.RecordCount <> 0 Then

        fgCashMemo.Rows = rsCashDetail.RecordCount + 1
                i = 0
        For i = 1 To rsCashDetail.RecordCount
            fgCashMemo.TextMatrix(i, 1) = rsCashDetail("BillSerialNo")
            fgCashMemo.TextMatrix(i, 2) = rsCashDetail("SerialNo")
            fgCashMemo.TextMatrix(i, 3) = rsCashDetail("ItemCode")
            fgCashMemo.TextMatrix(i, 4) = rsCashDetail("ItemName")
            fgCashMemo.TextMatrix(i, 5) = rsCashDetail("Qty")
            fgCashMemo.TextMatrix(i, 6) = rsCashDetail("Rate")
            fgCashMemo.TextMatrix(i, 7) = rsCashDetail("Tips")
            fgCashMemo.TextMatrix(i, 8) = rsCashDetail("ItemGroup")
            fgCashMemo.TextMatrix(i, 9) = rsCashDetail("ItemCatagory")
            fgCashMemo.TextMatrix(i, 10) = rsCashDetail("strDate")
            fgCashMemo.TextMatrix(i, 11) = rsCashDetail("CActive")
            fgCashMemo.TextMatrix(i, 12) = rsCashDetail("CPost")
            fgCashMemo.TextMatrix(i, 13) = rsCashDetail("NoDiscount")
            
    rsCashDetail.MoveNext
        Next
      End If
    rsCashDetail.Close

        
        
End If
End Sub

Private Sub cmdLDelete_Click()
'With fgCashMemo
'        If .Row = 0 Or .Row = -1 Then Exit Sub
'
'        If .Rows > 1 Then .RemoveItem .Row
'    End With
If fgCashMemo.Rows = 1 Then Exit Sub

     If fgCashMemo.Row >= 1 Then
      If MsgBox("Are you sure to delete the selected Item", vbYesNo, "General Setup") = vbYes Then fgCashMemo.RemoveItem fgCashMemo.Row
     Else
      MsgBox "You have to select a row to delete.", vbInformation, "General"
    End If
Call Calculation
End Sub


Private Sub cmdNew_Click()
'Call duplicate
    Set rs = New ADODB.Recordset
If cmdNew.Caption = "&New" Then
        strMood = "I"
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdClose.Enabled = False
        BillDate.Value = Date
        cmbATable.text = ""
        cmbATable.Enabled = False
        cmdtemSelected.Enabled = True
        cmdLDelete.Enabled = True
        cmdPost2.Enabled = False
        cmdActive.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmdOpen.Enabled = False
        txtBillSerialNo.Enabled = False
        cmdChange.Enabled = False
        Call Clear
        txtUName.text = frmLogin.txtUID.text
        txtCPost.text = "Not Posted"
        txtCActive.text = "Active"
        txtTime.text = Time

        fgCashMemo.Rows = 1
        fgCashMemo.Enabled = True
        fgCashMemo.Editable = flexEDKbdMouse
        Call allenable
        cmbTName.SetFocus
'           If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(SerialNo),0) as InvNo from tblCashMaster"
'           rs.Open str, cn, adOpenStatic, adLockReadOnly
'           txtBillSerialNo.text = Val(rs!InvNo) + 1

    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
'        fgLedger_AfterEdit
    If IsValidRecord Then
'        If IsValidtATable Then
           If rcupdate Then
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                fgCashMemo.Enabled = False
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
'                CmdDelete.Enabled = True
                txtBillSerialNo.Enabled = False
                cmbATable.Enabled = True
                cmdtemSelected.Enabled = False
                cmdLDelete.Enabled = False
                cmdPost2.Enabled = True
                cmdActive.Enabled = True
                cmdPrint.Enabled = True
                cmdChange.Enabled = False
'                cmdMovePrevious.Enabled = True
'                cmdMoveNext.Enabled = True
                Call alldisable
            If rs.State <> 0 Then rs.Close
               str = "Select ISNULL(max(SerialNo),0) as InvNo from tblCashMaster"
               rs.Open str, cn, adOpenStatic, adLockReadOnly
               txtBillSerialNo.text = Val(rs!InvNo)

                s = txtBillSerialNo
                rscashmaster.Requery
                rscashmaster.MoveFirst
                rscashmaster.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord
            End If
'        End If
    End If
End If

End Sub


Private Sub cmdNext_Click()
Dim i As Integer
    Dim strLedgerDetail As String
    Set rsCashDetail = New ADODB.Recordset

DCRSearch.Recordset.MoveNext
If DCRSearch.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    
    txtBillSerialNo = DCRSearch.Recordset!SerialNo
    BillDate = DCRSearch.Recordset!strDate
    cmbTName = DCRSearch.Recordset!TableName
    cmbWaiter = DCRSearch.Recordset!WaiterName
    txtPersonNo = DCRSearch.Recordset!Guest
    txtServiceCharge = DCRSearch.Recordset!ServiceCharge
    txtVAT = DCRSearch.Recordset!Vat
    txtDiscount = DCRSearch.Recordset!Discount
    txtDiscountAmt = DCRSearch.Recordset!DiscountAmt
    cmbDCard = DCRSearch.Recordset!DCard
    txtGName = DCRSearch.Recordset!GName
    txtCAddress = DCRSearch.Recordset!Address

    txtTotalBill = DCRSearch.Recordset!TotalBill
    txtTotalVat = DCRSearch.Recordset!TotalVat
    txtTotalSCharge = DCRSearch.Recordset!TSCharge
    txtTotalDiscount = DCRSearch.Recordset!TotalDiscount
    txtNPayable = DCRSearch.Recordset!NetPayable
    cboMode = DCRSearch.Recordset!PaymentMode
    txtRemarks = DCRSearch.Recordset!Remarks
    txtPaid = DCRSearch.Recordset!Paid
    txtDue = DCRSearch.Recordset!Due
    cmbATable = DCRSearch.Recordset!Post
    chkActive = DCRSearch.Recordset!Active
    txtCActive = DCRSearch.Recordset!CActive
    txtCPost = DCRSearch.Recordset!CPost
    txtUName = DCRSearch.Recordset!UName
    txtKOT = DCRSearch.Recordset!Kot
    txtBOT = DCRSearch.Recordset!Bot
    AdDate.Value = DCRSearch.Recordset!ADate
    txtAdvance = DCRSearch.Recordset!Advance
    txtMR_No = DCRSearch.Recordset!MrNo
          
        
        
        
        
        fgCashMemo.Rows = 1
    strLedgerDetail = "SELECT  BillSerialNo,SerialNo,ItemCode, ItemName, Qty, Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount" & _
                " FROM tblCashDetail " & _
                "WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo.text) & "'"
    rsCashDetail.CursorLocation = adUseClient
    rsCashDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsCashDetail.RecordCount <> 0 Then

        fgCashMemo.Rows = rsCashDetail.RecordCount + 1
                i = 0
        For i = 1 To rsCashDetail.RecordCount
            fgCashMemo.TextMatrix(i, 1) = rsCashDetail("BillSerialNo")
            fgCashMemo.TextMatrix(i, 2) = rsCashDetail("SerialNo")
            fgCashMemo.TextMatrix(i, 3) = rsCashDetail("ItemCode")
            fgCashMemo.TextMatrix(i, 4) = rsCashDetail("ItemName")
            fgCashMemo.TextMatrix(i, 5) = rsCashDetail("Qty")
            fgCashMemo.TextMatrix(i, 6) = rsCashDetail("Rate")
            fgCashMemo.TextMatrix(i, 7) = rsCashDetail("Tips")
            fgCashMemo.TextMatrix(i, 8) = rsCashDetail("ItemGroup")
            fgCashMemo.TextMatrix(i, 9) = rsCashDetail("ItemCatagory")
            fgCashMemo.TextMatrix(i, 10) = rsCashDetail("strDate")
            fgCashMemo.TextMatrix(i, 11) = rsCashDetail("CActive")
            fgCashMemo.TextMatrix(i, 12) = rsCashDetail("CPost")
            fgCashMemo.TextMatrix(i, 13) = rsCashDetail("NoDiscount")
            
    rsCashDetail.MoveNext
        Next
      End If
    rsCashDetail.Close

        
        
End If
End Sub

Private Sub cmdOpen_Click()
    frmCashMasterSearch.Show vbModal
    Call Calculation
    txtPaid.Enabled = True
    cboMode.Enabled = True
    
End Sub

Private Sub cmdPost2_Click()

Dim s As String
If txtCPost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 fgCashMemo.Enabled = False
                 cmdOpen.Enabled = True
                 cmdPreview.Enabled = True
                 cmdChange.Enabled = True
                 txtBillSerialNo.Enabled = False
                 Call alldisable
                 txtWords = InWords(txtNPayable.text)
'                 txtCPost.text = "Not Posted"
'                s = " & Val(cmbATable.Columns(1).text) & "
'                rsAMaster.Requery
'                rsAMaster.MoveFirst
'                rsAMaster.Find "Post='" & parseQuotes(s) & "'"
'                FindRecord1
            End If
        End If
'    End If
Else
    cmdtemSelected.Enabled = False
    cmdLDelete.Enabled = False
    
End If
    
End Sub

Private Sub cmdPrevious_Click()
Dim i As Integer
    Dim strLedgerDetail As String
    Set rsCashDetail = New ADODB.Recordset

DCRSearch.Recordset.MovePrevious
If DCRSearch.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdPrevious.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    
    txtBillSerialNo = DCRSearch.Recordset!SerialNo
    BillDate = DCRSearch.Recordset!strDate
    cmbTName = DCRSearch.Recordset!TableName
    cmbWaiter = DCRSearch.Recordset!WaiterName
    txtPersonNo = DCRSearch.Recordset!Guest
    txtServiceCharge = DCRSearch.Recordset!ServiceCharge
    txtVAT = DCRSearch.Recordset!Vat
    txtDiscount = DCRSearch.Recordset!Discount
    txtDiscountAmt = DCRSearch.Recordset!DiscountAmt
    cmbDCard = DCRSearch.Recordset!DCard
    txtGName = DCRSearch.Recordset!GName
    txtCAddress = DCRSearch.Recordset!Address

    txtTotalBill = DCRSearch.Recordset!TotalBill
    txtTotalVat = DCRSearch.Recordset!TotalVat
    txtTotalSCharge = DCRSearch.Recordset!TSCharge
    txtTotalDiscount = DCRSearch.Recordset!TotalDiscount
    txtNPayable = DCRSearch.Recordset!NetPayable
    cboMode = DCRSearch.Recordset!PaymentMode
    txtRemarks = DCRSearch.Recordset!Remarks
    txtPaid = DCRSearch.Recordset!Paid
    txtDue = DCRSearch.Recordset!Due
    cmbATable = DCRSearch.Recordset!Post
    chkActive = DCRSearch.Recordset!Active
    txtCActive = DCRSearch.Recordset!CActive
    txtCPost = DCRSearch.Recordset!CPost
    txtUName = DCRSearch.Recordset!UName
    txtKOT = DCRSearch.Recordset!Kot
    txtBOT = DCRSearch.Recordset!Bot
    AdDate.Value = DCRSearch.Recordset!ADate
    txtAdvance = DCRSearch.Recordset!Advance
    txtMR_No = DCRSearch.Recordset!MrNo
          
        
        
        
        
    fgCashMemo.Rows = 1
    strLedgerDetail = "SELECT  BillSerialNo,SerialNo,ItemCode, ItemName, Qty, Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount" & _
                " FROM tblCashDetail " & _
                "WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo.text) & "'"
    rsCashDetail.CursorLocation = adUseClient
    rsCashDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsCashDetail.RecordCount <> 0 Then

        fgCashMemo.Rows = rsCashDetail.RecordCount + 1
                i = 0
        For i = 1 To rsCashDetail.RecordCount
            fgCashMemo.TextMatrix(i, 1) = rsCashDetail("BillSerialNo")
            fgCashMemo.TextMatrix(i, 2) = rsCashDetail("SerialNo")
            fgCashMemo.TextMatrix(i, 3) = rsCashDetail("ItemCode")
            fgCashMemo.TextMatrix(i, 4) = rsCashDetail("ItemName")
            fgCashMemo.TextMatrix(i, 5) = rsCashDetail("Qty")
            fgCashMemo.TextMatrix(i, 6) = rsCashDetail("Rate")
            fgCashMemo.TextMatrix(i, 7) = rsCashDetail("Tips")
            fgCashMemo.TextMatrix(i, 8) = rsCashDetail("ItemGroup")
            fgCashMemo.TextMatrix(i, 9) = rsCashDetail("ItemCatagory")
            fgCashMemo.TextMatrix(i, 10) = rsCashDetail("strDate")
            fgCashMemo.TextMatrix(i, 11) = rsCashDetail("CActive")
            fgCashMemo.TextMatrix(i, 12) = rsCashDetail("CPost")
            fgCashMemo.TextMatrix(i, 13) = rsCashDetail("NoDiscount")
            
            
    rsCashDetail.MoveNext
        Next
      End If
    rsCashDetail.Close

        
        
End If
End Sub

Private Sub cmdPrint_Click()
Dim s As String
If cmdPrint.Caption = "&Print" Then
cmdPrint.Caption = "&Printing"
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Enabled = True
                cmdCancel.Enabled = True
                cmdClose.Enabled = True
                fgCashMemo.Enabled = False
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
                cmdChange.Enabled = True
                txtBillSerialNo.Enabled = False
                Call alldisable
                txtWords = InWords(txtNPayable.text)

            End If
        End If
    End If
    
Tracer = 1
Screen.MousePointer = vbHourglass
If txtCPost.text = "Posted" Then

Call GuestCopy
'Call CashCopy
'Call printReport
End If
Screen.MousePointer = vbDefault

cmdPrint.Caption = "&Print"

End Sub

Public Sub printReport()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Set rsDailyRpt = New ADODB.Recordset
    

If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "select SerialNo,strDate,TableName,WaiterName,Guest,ServiceCharge,Vat,Discount,GName,Address,strTime, " & _
                          "TotalBill,TotalVat,TSCharge,TotalDiscount,NetPayable,PaymentMode,Remarks," & _
                          "Paid,Due,Active,Post,CActive,CPost,UName,BOT,KOT,Advance from tblCashMaster", cn, adOpenStatic, adLockReadOnly
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\CashMemoFinal.RPT"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        


'    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close

           
rsDailyRpt.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate,tblCashMaster.TableName," & _
                "tblCashMaster.WaiterName,tblCashMaster.Guest," & _
                "tblCashMaster.ServiceCharge,tblCashMaster.Vat,tblCashMaster.Discount," & _
                "tblCashMaster.GName, tblCashMaster.Address,tblCashMaster.PaymentMode," & _
                "tblCashMaster.NetPayable,tblCashDetail.ItemCode, tblCashDetail.ItemName," & _
                "tblCashDetail.Qty,tblCashDetail.Rate,((tblCashDetail.Qty)*(tblCashDetail.Rate))as Amount,tblCashDetail.NoVAT,tblCashMaster.TotalVat, " & _
                "tblCashMaster.TSCharge,tblCashMaster.TotalDiscount,tblCashMaster.Due FROM tblCashMaster,tblCashDetail Where tblCashMaster.SerialNo=tblCashDetail.BillSerialNo " & _
                "AND tblCashMaster.SerialNo ='" & txtBillSerialNo.text & "' order by tblCashDetail.ItemCode", cn, adOpenStatic
                    
                   
        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"
            
            Set objReportFF = objReportFormulaFieldDefinations.Item(2)
            objReportFF.text = "'" + parseQuotes(txtUName.text) + " '"
            
'-------------End Add Discount-------------------
        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        If Tracer = 0 Then
        objReport.Preview "Cash Memo Report", , , , , 16777216 Or 524288 Or 65536
        Else
        objReport.PrintOut (False)
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

Private Sub cmdtemSelected_Click()
frmCashSelectIteam.Show vbModal
Call Calculation
End Sub

Private Sub fgCashMemo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim j As Integer
     Select Case Col
        Case 9
'------------------duplicate folio-----------
Case 4
Dim k As Integer
If fgCashMemo.Rows > 2 Then
    For k = 1 To fgCashMemo.Rows - 1
 If (fgCashMemo.TextMatrix(k, 4)) = fgCashMemo.TextMatrix(Row, 4) And k <> fgCashMemo.Row Then
        MsgBox "Duplicate Item Name.", vbInformation
        fgCashMemo.TextMatrix(Row, 4) = ""

     End If

   Next
End If
'-------------------------------------------
End Select

Call Calculation
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If Trim(cmbTName) = "" Then
        MsgBox "Please Select Table Number", vbInformation
        cmbTName.SetFocus
        IsValidRecord = False
        Exit Function
'    End If
    
    
    ElseIf Trim(cmbWaiter) = "" Then
        MsgBox "Please Select Waiter Name", vbInformation
        cmbWaiter.SetFocus
        IsValidRecord = False
        Exit Function
'    End If
    
    
    ElseIf Trim(cboMode) = "" Then
        MsgBox "Please Select Payment Mode", vbInformation
        cboMode.SetFocus
        IsValidRecord = False
        Exit Function
'    End If

 ElseIf Trim(fgCashMemo.Rows) < 2 Then
        MsgBox "Please input Item Name", vbInformation
        txtItemCode.SetFocus
        IsValidRecord = False
        Exit Function
'    End If

 ElseIf cmdPrint.Caption = "&Printing" And Trim(cmbATable) = "" Or Trim(cmbATable) = 0 Then
        MsgBox "Please Select Active Table Number.", vbInformation
        cmbATable.SetFocus
        IsValidRecord = False
        Exit Function
        
        
ElseIf cmdEdit.Caption = "&Update" And Trim(cmbATable) = "" Or Trim(cmbATable) = 0 Then
        MsgBox "Please Select Active Table Number.", vbInformation
        cmbATable.SetFocus
        IsValidRecord = False
        Exit Function
        
'ElseIf Duplicate = True Then Exit Sub

   ElseIf cmdNew.Caption = "&Save" Or cmdEdit.Caption = "&Update" Then
         
 Dim j As Integer

  For j = 1 To fgCashMemo.Rows - 2

If Trim(fgCashMemo.TextMatrix(j, 3)) = Trim(fgCashMemo.TextMatrix(j + 1, 3)) Then
    MsgBox "Duplicate Item Code Number.", vbInformation, "Confirmation"
    '             fgItem.TextMatrix(j, 4) = ""
    fgCashMemo.Select j, 3
    '             fgItem.RemoveItem fgItem.Row
    IsValidRecord = False

End If

Next

Exit Function
Else
If Trim(cmbATable) = "" Or Trim(cmbATable) = 0 Then
   MsgBox "Please Select Active Table.", vbInformation
   cmbATable.SetFocus
   IsValidRecord = True
   cmbATable.SetFocus
   Exit Function
        
        End If

    End If
    
        
    End Function
    
Private Function IsValidtATable() As Boolean

IsValidtATable = False

        If rscashmaster.State <> 0 Then rscashmaster.Close
            rscashmaster.Open "select * from tblCashMaster where CPost='Not Posted'", cn

If rscashmaster.RecordCount > 0 Then
        rscashmaster.MoveFirst
    End If

While Not rscashmaster.EOF
    
      If (rscashmaster!CPost) = "Not Posted" And cmbTName = rscashmaster!TableName Then
            MsgBox "This Table Already Engaged, Please Select Another Table.", vbInformation, "Confirmation"
            cmbTName.SetFocus
            cmbTName.text = ""
            IsValidtATable = True
            Exit Function
        End If
'    End If
    rscashmaster.MoveNext
Wend

End Function


Private Sub fgCashMemo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case Col
   Case 3, 4, 6, 7

      Cancel = True
End Select
End Sub

Private Sub fgCashMemo_CellChanged(ByVal Row As Long, ByVal Col As Long)
'Dim j As Integer
'
'If fgCashMemo.Cell(flexcpChecked, j, 13) = flexUnchecked Then
'If fgCashMemo.Cell(flexcpChecked, j, 14) = flexUnchecked Then
'Call Calculation
'
'End If
'End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call Connect
   ModFunction.StartUpPosition Me
   
        
         Call alldisable
         Call TableName
         Call WaiterName
         Call CName
         Call ActiveTable
         
         cmdtemSelected.Enabled = False
         cmdLDelete.Enabled = False
         
         txtCPost.text = "Not Posted"
         txtCActive.text = "Active"
         
    Set rscashmaster = New ADODB.Recordset

    If rscashmaster.State <> 0 Then rscashmaster.Close
    rscashmaster.Open "select TOP 1 SerialNo,strDate,TableName,WaiterName,Guest,ServiceCharge,Vat,Discount,DiscountAmt,DCard,GName,Address,strTime, " & _
                          "TotalBill,TotalVat,TSCharge,TotalDiscount,NetPayable,PaymentMode,Remarks," & _
                          "Paid,Due,Active,Post,CActive,CPost,UName,Kot,Bot,ADate,Advance,CReceived,CashBack,MrNo from tblCashMaster ORDER BY SerialNo DESC", cn, adOpenStatic, adLockReadOnly


   If rscashmaster.RecordCount > 0 Then
        rscashmaster.MoveLast
    End If


    If Not rscashmaster.EOF Then FindRecord
        txtBillSerialNo.Enabled = False

       cboMode.AddItem "CASH"
       cboMode.AddItem "CREDIT"
       cboMode.AddItem "CITY BANK"
       cboMode.AddItem "BRAC BANK"
       cboMode.AddItem "CREDIT CARD"
       cboMode.AddItem "ADVANCE"
       cboMode.AddItem "CHEQUE"
       

       cboMode.text = "Cash"
       Call Calculation

txtPaid.Enabled = True

'-----------------For Record Search----------
DCRSearch.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  DCRSearch.CommandType = adCmdTable
  DCRSearch.RecordSource = "tblCashMaster"

  DCRSearch.Refresh
'-------------------End Record Search---------
Call changeVisible
AdDate.Value = Date

Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     fgExport.Editable = flexEDKbdMouse
     fgExport.ColDataType(1) = flexDTBoolean
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
'            rsTemp.CursorLocation = adUseClient
     rsTemp.Open "SELECT TOP 50 SerialNo,ItemCode,ItemName,ItemQty,ItemPrice,Tips,NoDiscount,ItemGroup,ItemCatagory FROM tblItemDetail", cn, adOpenStatic, adLockReadOnly
          
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ItemCode") & _
         vbTab & rsTemp("ItemName") & vbTab & rsTemp("ItemQty") & vbTab & rsTemp("ItemPrice") & _
         vbTab & rsTemp("Tips") & vbTab & rsTemp("NoDiscount") & vbTab & rsTemp("ItemGroup") & vbTab & rsTemp("ItemCatagory")
         
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
    For i = 0 To fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next
'====================================================================
End Sub


Private Sub Calculation()

Dim j As Integer
  Dim i As Integer
 
       temp = 0
       temp1 = 0
       temp2 = 0
       temp3 = 0
    For j = 1 To fgCashMemo.Rows - 1

temp = temp + CDbl(Val(fgCashMemo.TextMatrix(j, 5)) * CDbl(Val(fgCashMemo.TextMatrix(j, 6))))

If fgCashMemo.Cell(flexcpChecked, j, 13) = flexUnchecked Then
   temp2 = temp2 + CDbl(Val(fgCashMemo.TextMatrix(j, 5)) * CDbl(Val(fgCashMemo.TextMatrix(j, 6))))
End If

If fgCashMemo.Cell(flexcpChecked, j, 14) = flexUnchecked Then
   temp3 = temp3 + CDbl(Val(fgCashMemo.TextMatrix(j, 5)) * CDbl(Val(fgCashMemo.TextMatrix(j, 6))))

End If

   Next
   txtTotalBill = temp

txtTotalVat = (temp3 * CDbl(Val(txtVAT) / 100))
txtTotalSCharge = temp * CDbl(Val(txtServiceCharge) / 100)
txtTotalDiscount = temp2 * CDbl(Val(txtDiscount) / 100) + CDbl(Val(txtDiscountAmt))
txtNPayable = CDbl(txtTotalBill) + CDbl(txtTotalVat) + CDbl(txtTotalSCharge) - CDbl(txtTotalDiscount)
txtDue = (CDbl(txtTotalBill) + CDbl(txtTotalVat) + CDbl(txtTotalSCharge) - CDbl(txtTotalDiscount)) - CDbl(Val(txtPaid)) - CDbl(Val(txtAdvance))
txtWords = InWords(txtNPayable.text)

'  Dim j As Integer
'  Dim i As Integer
'
'       temp = 0
'       temp1 = 0
'       temp2 = 0
'       temp3 = 0
'    For j = 1 To fgCashMemo.Rows - 1
'
'temp = temp + CDbl(Val(fgCashMemo.TextMatrix(j, 5)) * CDbl(Val(fgCashMemo.TextMatrix(j, 6))))
'
'If fgCashMemo.Cell(flexcpChecked, j, 13) = flexUnchecked Then
'    If Val(fgCashMemo.TextMatrix(j, 13)) = False Then
'         temp2 = temp2 + CDbl(Val(fgCashMemo.TextMatrix(j, 5)) * CDbl(Val(fgCashMemo.TextMatrix(j, 6))))
'End If
'
'If fgCashMemo.Cell(flexcpChecked, j, 14) = flexUnchecked Then
'    If Val(fgCashMemo.TextMatrix(j, 14)) = False Then
'         temp3 = temp3 + CDbl(Val(fgCashMemo.TextMatrix(j, 5)) * CDbl(Val(fgCashMemo.TextMatrix(j, 6))))
'    End If
'End If
'
'End If
'
'Next
'txtTotalBill = temp
'
'txtTotalVat = (temp3 * CDbl(Val(txtVAT) / 100))
'txtTotalSCharge = temp * CDbl(Val(txtServiceCharge) / 100)
'txtTotalDiscount = temp2 * CDbl(Val(txtDiscount) / 100) + CDbl(Val(txtDiscountAmt))
'txtNPayable = CDbl(txtTotalBill) + CDbl(txtTotalVat) + CDbl(txtTotalSCharge) - CDbl(txtTotalDiscount)
'txtDue = (CDbl(txtTotalBill) + CDbl(txtTotalVat) + CDbl(txtTotalSCharge) - CDbl(txtTotalDiscount)) - CDbl(Val(txtPaid)) - CDbl(Val(txtAdvance))
'txtWords = InWords(txtNPayable.text)
'
'End Sub
'
'
'Private Sub cmbATable_DropDown()
'Call ActiveTable
'dcATable.Refresh
'txtWords = InWords(txtNPayable.text)

End Sub

Private Sub cmbATable_Click()

 Set rsAMaster = New ADODB.Recordset
'    Call ActiveTable
    If rsAMaster.State <> 0 Then rsAMaster.Close
        rsAMaster.Open "select SerialNo,strDate,TableName,WaiterName,Guest,ServiceCharge,Vat,Discount,DiscountAmt,DCard,GName,Address,strTime, " & _
                          "TotalBill,TotalVat,TSCharge,TotalDiscount,NetPayable,PaymentMode,Remarks," & _
                          "Paid,Due,Active,Post,CActive,CPost,UName,Kot,Bot,ADate,Advance,CReceived,CashBack,MrNo from tblCashMaster where TableName= '" & parseQuotes(cmbATable) & "' and CPost='Not Posted'", cn, adOpenStatic, adLockReadOnly

   If rsAMaster.RecordCount > 0 Then
        rsAMaster.MoveFirst
    End If
    
    If Not rsAMaster.EOF Then FindRecord1

txtWords = InWords(txtNPayable.text)
'
  
End Sub

Private Sub ActiveTable()

Dim rsTemp2 As New ADODB.Recordset
    cmbATable.Clear
     rsTemp2.Open ("SELECT DISTINCT TableName FROM tblCashMaster where CPost='Not Posted' ORDER BY TableName ASC"), cn, adOpenStatic
      
    While Not rsTemp2.EOF
        cmbATable.AddItem rsTemp2("TableName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
    
End Sub


Private Function rcupdate() As Boolean
'    On Error GoTo ErrHandler
      Dim strSQL As String
      Dim iRow As Integer
      Dim j As Integer
      Dim strDate As String
      Dim strActive As String
      Dim blnNoDiscount, blnNoVat    As Boolean
'     cmbATable.BackColorOdd = &HFFFF00
      
'   cn.BeginTrans
    flagSlNo = 0
    If cmdNew.Caption = "&Save" Then
      
strSQL = "INSERT INTO tblCashMaster (strDate,TableName,WaiterName,Guest,ServiceCharge,Vat,Discount,DiscountAmt, " & _
         "DCard,GName,Address,strTime,TotalBill,TotalVat,TSCharge,TotalDiscount,NetPayable,PaymentMode,Remarks, " & _
         "Paid,Due,Post,Active,CActive,CPost,UName,Kot,Bot,ADate,Advance,CReceived,CashBack,MrNo) " & _
         "VALUES ('" & Format(BillDate, "dd-mmm-yyyy") & "','" & parseQuotes(cmbTName) & "'," & _
         " '" & parseQuotes(cmbWaiter) & "', " & _
         " " & Val(txtPersonNo.text) & "," & Val(txtServiceCharge.text) & "," & Val(txtVAT.text) & "," & _
         " " & Val(txtDiscount.text) & "," & Val(txtDiscountAmt.text) & ", " & _
         " '" & parseQuotes(cmbDCard) & "','" & parseQuotes(txtGName) & "','" & parseQuotes(txtCAddress) & "','" & txtTime.text & "'," & Val(txtTotalBill.text) & ", " & _
         " " & Val(txtTotalVat.text) & "," & Val(txtTotalSCharge.text) & "," & Val(txtTotalDiscount.text) & "," & Val(txtNPayable.text) & ",'" & cboMode & "','" & parseQuotes(txtRemarks) & "', " & _
         " " & Val(txtPaid.text) & "," & Val(txtDue.text) & ",'" & parseQuotes(cmbATable) & "'," & chkActive & ",'" & txtCActive.text & "','" & txtCPost.text & "','" & txtUName.text & "', " & _
         "'" & txtKOT.text & "','" & txtBOT.text & "','" & Format(AdDate, "dd-mmm-yyyy") & "'," & Val(txtAdvance.text) & "," & Val(txtCashReceived.text) & "," & Val(txtCashBack.text) & ",'" & txtMR_No.text & "')"
        
        
    cn.Execute strSQL
    
    If rs.State <> 0 Then rs.Close
       str = "Select ISNULL(max(SerialNo),0) as InvNo from tblCashMaster"
       rs.Open str, cn, adOpenStatic, adLockReadOnly
       txtBillSerialNo.text = Val(rs!InvNo)

            j = 0
            For j = 1 To fgCashMemo.Rows - 1
            
        If fgCashMemo.Cell(flexcpChecked, j, 13) = flexChecked Then
               blnNoDiscount = True
            Else
                blnNoDiscount = False
            End If
            
            
            If fgCashMemo.Cell(flexcpChecked, j, 14) = flexChecked Then
               blnNoVat = True
            Else
                blnNoVat = False
            End If

cn.Execute "INSERT INTO tblCashDetail (BillSerialNo,SerialNo,ItemCode,ItemName,Qty,Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT) " & _
            "Values ('" & parseQuotes(txtBillSerialNo) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 2)) & "'," & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 3)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 4)) & "'," & _
            IIf(fgCashMemo.TextMatrix(j, 5) = "", "0", fgCashMemo.TextMatrix(j, 5)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 6) = "", "0", fgCashMemo.TextMatrix(j, 6)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 7) = "", "0", fgCashMemo.TextMatrix(j, 7)) & ", " & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 8)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 9)) & "', " & _
            " '" & Format(BillDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtCActive) & "', " & _
            " '" & parseQuotes(txtCPost) & "'," & _
            IIf(blnNoDiscount, 1, 0) & "," & _
            IIf(blnNoVat, 1, 0) & ")"
'
               Next
       rcupdate = True
'       cn.CommitTrans
       MsgBox "Record added Successfully", vbInformation, "Confirmation"
     Call ActiveTable
    
    ElseIf (cmdEdit.Caption = "&Update") Then
    ' Update Information

        cn.Execute "UPDATE tblCashMaster SET strDate='" & Format(BillDate, "dd-mmm-yyyy") & "',TableName='" & parseQuotes(cmbTName) & "', " & _
                   " WaiterName='" & parseQuotes(cmbWaiter) & "',Guest=" & Val(txtPersonNo.text) & ",ServiceCharge=" & Val(txtServiceCharge.text) & ", " & _
                   "Vat=" & Val(txtVAT.text) & ",Discount=" & Val(txtDiscount.text) & ",DiscountAmt=" & Val(txtDiscountAmt.text) & ",DCard='" & parseQuotes(cmbDCard.text) & "'," & _
                   "GName='" & parseQuotes(txtGName) & "',Address='" & parseQuotes(txtCAddress) & "',strTime='" & (txtTime.text) & "',TotalBill=" & Val(txtTotalBill.text) & ", " & _
                   "TotalVat=" & Val(txtTotalVat.text) & ",TSCharge=" & Val(txtTotalSCharge.text) & ",TotalDiscount=" & Val(txtTotalDiscount.text) & ", " & _
                   "NetPayable=" & Val(txtNPayable.text) & ",PaymentMode='" & parseQuotes(cboMode) & "', " & _
                   "Remarks='" & parseQuotes(txtRemarks) & "',Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ",Post='" & parseQuotes(cmbATable) & "',Active=" & chkActive & ", " & _
                   "CActive='" & parseQuotes(txtCActive.text) & "',CPost='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "', Kot='" & txtKOT.text & "',Bot='" & txtBOT.text & "', " & _
                   "ADate='" & Format(AdDate, "dd-mmm-yyyy") & "',Advance=" & Val(txtAdvance.text) & ",CReceived=" & Val(txtCashReceived.text) & ",CashBack=" & Val(txtCashBack.text) & ",MrNo='" & txtMR_No.text & "' WHERE SerialNo = '" & txtBillSerialNo & "'"

'''    'ImpYarnDetail Information
        cn.Execute "DELETE FROM tblCashDetail WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo) & "'"

                j = 0
            For j = 1 To fgCashMemo.Rows - 1
          
          If fgCashMemo.Cell(flexcpChecked, j, 13) = flexChecked Then
               blnNoDiscount = True
            Else
                blnNoDiscount = False
            End If
            
            
            If fgCashMemo.Cell(flexcpChecked, j, 14) = flexChecked Then
               blnNoVat = True
            Else
                blnNoVat = False
            End If

cn.Execute "INSERT INTO tblCashDetail (BillSerialNo,SerialNo,ItemCode,ItemName,Qty,Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT) " & _
            "Values ('" & parseQuotes(txtBillSerialNo) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 2)) & "'," & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 3)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 4)) & "'," & _
            IIf(fgCashMemo.TextMatrix(j, 5) = "", "0", fgCashMemo.TextMatrix(j, 5)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 6) = "", "0", fgCashMemo.TextMatrix(j, 6)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 7) = "", "0", fgCashMemo.TextMatrix(j, 7)) & ", " & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 8)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 9)) & "', " & _
            " '" & Format(BillDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtCActive) & "', " & _
            " '" & parseQuotes(txtCPost) & "'," & _
            IIf(blnNoDiscount, 1, 0) & "," & _
            IIf(blnNoVat, 1, 0) & ")"
                           
               Next
        Call ActiveTable
        rcupdate = True
        
        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
'        Call ActiveTable
'----------------------------------------------Printing Start--------------------------
  ElseIf cmdPrint.Caption = "&Printing" Then

Dim iprint

iprint = MsgBox("Do you want to Print this bill?", vbYesNo)

If iprint = vbYes Then
'If iprint Then
         
txtCPost.text = "Posted"
 cn.Execute "UPDATE tblCashMaster SET strDate='" & Format(BillDate, "dd-mmm-yyyy") & "',TableName='" & parseQuotes(cmbTName) & "', " & _
          " WaiterName='" & parseQuotes(cmbWaiter) & "',Guest=" & Val(txtPersonNo.text) & ",ServiceCharge=" & Val(txtServiceCharge.text) & ", " & _
          "Vat=" & Val(txtVAT.text) & ",Discount=" & Val(txtDiscount.text) & ",DiscountAmt=" & Val(txtDiscountAmt.text) & ",DCard='" & parseQuotes(cmbDCard.text) & "'," & _
          "GName='" & parseQuotes(txtGName) & "',Address='" & parseQuotes(txtCAddress) & "',strTime='" & (txtTime.text) & "',TotalBill=" & Val(txtTotalBill.text) & ", " & _
          "TotalVat=" & Val(txtTotalVat.text) & ",TSCharge=" & Val(txtTotalSCharge.text) & ",TotalDiscount=" & Val(txtTotalDiscount.text) & ", " & _
          "NetPayable=" & Val(txtNPayable.text) & ",PaymentMode='" & parseQuotes(cboMode) & "', " & _
          "Remarks='" & parseQuotes(txtRemarks) & "',Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ",Post='" & parseQuotes(cmbATable) & "',Active=" & chkActive & ", " & _
          "CActive='" & parseQuotes(txtCActive.text) & "',CPost='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "', Kot='" & txtKOT.text & "',Bot='" & txtBOT.text & "', " & _
          "ADate='" & Format(AdDate, "dd-mmm-yyyy") & "',Advance=" & Val(txtAdvance.text) & ",CReceived=" & Val(txtCashReceived.text) & ",CashBack=" & Val(txtCashBack.text) & ",MrNo='" & txtMR_No.text & "' WHERE SerialNo = '" & txtBillSerialNo & "'"

        cn.Execute "DELETE FROM tblCashDetail WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo) & "'"

                j = 0
            For j = 1 To fgCashMemo.Rows - 1
            
            If fgCashMemo.Cell(flexcpChecked, j, 13) = flexChecked Then
               blnNoDiscount = True
            Else
                blnNoDiscount = False
            End If
            
            
            If fgCashMemo.Cell(flexcpChecked, j, 14) = flexChecked Then
               blnNoVat = True
            Else
                blnNoVat = False
            End If


cn.Execute "INSERT INTO tblCashDetail (BillSerialNo,SerialNo,ItemCode,ItemName,Qty,Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT) " & _
            "Values ('" & parseQuotes(txtBillSerialNo) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 2)) & "'," & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 3)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 4)) & "'," & _
            IIf(fgCashMemo.TextMatrix(j, 5) = "", "0", fgCashMemo.TextMatrix(j, 5)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 6) = "", "0", fgCashMemo.TextMatrix(j, 6)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 7) = "", "0", fgCashMemo.TextMatrix(j, 7)) & ", " & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 8)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 9)) & "', " & _
            " '" & Format(BillDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtCActive) & "', " & _
            " '" & parseQuotes(txtCPost) & "'," & _
            IIf(blnNoDiscount, 1, 0) & "," & _
            IIf(blnNoVat, 1, 0) & ")"
               Next
              
        rcupdate = True
'        MsgBox "Record is Printing Now", vbInformation, "Confirmation"
        
        End If
        Call ActiveTable
'------------------------------Printing End-----------------------------
      
'------------------------------Active Bill------------------------------
  ElseIf cmdActive.Caption = "&Void" Then

Dim iActive
Dim iCoppy
chkActive.Value = 1
Set rs = New ADODB.Recordset
iActive = MsgBox("Do you want to Void this bill?", vbYesNo)

If iActive = vbYes Then
       
        txtCActive.text = "Void"
        txtCPost.text = "Posted"
        cmdChange.Enabled = False
        txtDue = Val(txtNPayable)
        txtPaid = Val(txtNPayable) - Val(txtDue)
        
cn.Execute "UPDATE tblCashMaster SET strDate='" & Format(BillDate, "dd-mmm-yyyy") & "',TableName='" & parseQuotes(cmbTName) & "', " & _
             " WaiterName='" & parseQuotes(cmbWaiter) & "',Guest=" & Val(txtPersonNo.text) & ",ServiceCharge=" & Val(txtServiceCharge.text) & ", " & _
             "Vat=" & Val(txtVAT.text) & ",Discount=" & Val(txtDiscount.text) & ",DiscountAmt=" & Val(txtDiscountAmt.text) & ",DCard='" & parseQuotes(cmbDCard.text) & "'," & _
             "GName='" & parseQuotes(txtGName) & "',Address='" & parseQuotes(txtCAddress) & "',strTime='" & (txtTime.text) & "',TotalBill=" & Val(txtTotalBill.text) & ", " & _
             "TotalVat=" & Val(txtTotalVat.text) & ",TSCharge=" & Val(txtTotalSCharge.text) & ",TotalDiscount=" & Val(txtTotalDiscount.text) & ", " & _
             "NetPayable=" & Val(txtNPayable.text) & ",PaymentMode='" & parseQuotes(cboMode) & "', " & _
             "Remarks='" & parseQuotes(txtRemarks) & "',Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ",Post='" & parseQuotes(cmbATable) & "',Active=" & chkActive & ", " & _
             "CActive='" & parseQuotes(txtCActive.text) & "',CPost='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "', Kot='" & txtKOT.text & "',Bot='" & txtBOT.text & "', " & _
             "ADate='" & Format(AdDate, "dd-mmm-yyyy") & "',Advance=" & Val(txtAdvance.text) & ",CReceived=" & Val(txtCashReceived.text) & ",CashBack=" & Val(txtCashBack.text) & ",MrNo='" & txtMR_No.text & "' WHERE SerialNo = '" & txtBillSerialNo & "'"


'''    'ImpYarnDetail Information
        cn.Execute "DELETE FROM tblCashDetail WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo) & "'"

                j = 0
            For j = 1 To fgCashMemo.Rows - 1
            
            If fgCashMemo.Cell(flexcpChecked, j, 13) = flexChecked Then
               blnNoDiscount = True
            Else
                blnNoDiscount = False
            End If
            
            
            If fgCashMemo.Cell(flexcpChecked, j, 14) = flexChecked Then
               blnNoVat = True
            Else
                blnNoVat = False
            End If


cn.Execute "INSERT INTO tblCashDetail (BillSerialNo,SerialNo,ItemCode,ItemName,Qty,Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT) " & _
            "Values ('" & parseQuotes(txtBillSerialNo) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 2)) & "'," & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 3)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 4)) & "'," & _
            IIf(fgCashMemo.TextMatrix(j, 5) = "", "0", fgCashMemo.TextMatrix(j, 5)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 6) = "", "0", fgCashMemo.TextMatrix(j, 6)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 7) = "", "0", fgCashMemo.TextMatrix(j, 7)) & ", " & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 8)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 9)) & "', " & _
            " '" & Format(BillDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtCActive) & "', " & _
            " '" & parseQuotes(txtCPost) & "'," & _
            IIf(blnNoDiscount, 1, 0) & "," & _
            IIf(blnNoVat, 1, 0) & ")"
               
               Next
                
        rcupdate = True
'        cn.CommitTrans
        MsgBox "Record is Void Now", vbInformation, "Confirmation"
        chkActive.Value = 0
        
        Else
        chkActive.Value = 0
        rcupdate = True
        cn.CommitTrans
        End If

iCoppy = MsgBox("Do you want to coppy this bill?", vbYesNo)

      If iCoppy = vbYes Then
      
           txtCPost.text = "Not Posted"
           txtCActive.text = "Active"
           txtPaid.text = 0
         
           cmdNew.Caption = "&Save"
           Call allenable
'           fgCashMemo.Editable = flexEDKbdMouse
           If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),0) as InvNo from tblCashMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtBillSerialNo.text = Val(rs!InvNo)

       End If
' End If
'-------------------------------End Of Active Bill-----------------------
    
Else
'    ---------------------------Post Information--------------------------

Dim ipost

ipost = MsgBox("Do you want to post this bill?", vbYesNo)

If ipost = vbYes Then

  txtCPost.text = "Posted"
cn.Execute "UPDATE tblCashMaster SET strDate='" & Format(BillDate, "dd-mmm-yyyy") & "',TableName='" & parseQuotes(cmbTName) & "', " & _
            " WaiterName='" & parseQuotes(cmbWaiter) & "',Guest=" & Val(txtPersonNo.text) & ",ServiceCharge=" & Val(txtServiceCharge.text) & ", " & _
            "Vat=" & Val(txtVAT.text) & ",Discount=" & Val(txtDiscount.text) & ",DiscountAmt=" & Val(txtDiscountAmt.text) & ",DCard='" & parseQuotes(cmbDCard.text) & "'," & _
            "GName='" & parseQuotes(txtGName) & "',Address='" & parseQuotes(txtCAddress) & "',strTime='" & (txtTime.text) & "',TotalBill=" & Val(txtTotalBill.text) & ", " & _
            "TotalVat=" & Val(txtTotalVat.text) & ",TSCharge=" & Val(txtTotalSCharge.text) & ",TotalDiscount=" & Val(txtTotalDiscount.text) & ", " & _
            "NetPayable=" & Val(txtNPayable.text) & ",PaymentMode='" & parseQuotes(cboMode) & "', " & _
            "Remarks='" & parseQuotes(txtRemarks) & "',Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ",Post='" & parseQuotes(cmbATable) & "',Active=" & chkActive & ", " & _
            "CActive='" & parseQuotes(txtCActive.text) & "',CPost='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "', Kot='" & txtKOT.text & "',Bot='" & txtBOT.text & "', " & _
            "ADate='" & Format(AdDate, "dd-mmm-yyyy") & "',Advance=" & Val(txtAdvance.text) & ",CReceived=" & Val(txtCashReceived.text) & ",CashBack=" & Val(txtCashBack.text) & ",MrNo='" & txtMR_No.text & "' WHERE SerialNo = '" & txtBillSerialNo & "'"


'''    'ImpYarnDetail Information
        cn.Execute "DELETE FROM tblCashDetail WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo) & "'"

                j = 0
            For j = 1 To fgCashMemo.Rows - 1
            
            If fgCashMemo.Cell(flexcpChecked, j, 13) = flexChecked Then
               blnNoDiscount = True
            Else
                blnNoDiscount = False
            End If
            
            
            If fgCashMemo.Cell(flexcpChecked, j, 14) = flexChecked Then
               blnNoVat = True
            Else
                blnNoVat = False
            End If


cn.Execute "INSERT INTO tblCashDetail (BillSerialNo,SerialNo,ItemCode,ItemName,Qty,Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT) " & _
            "Values ('" & parseQuotes(txtBillSerialNo) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 2)) & "'," & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 3)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 4)) & "'," & _
            IIf(fgCashMemo.TextMatrix(j, 5) = "", "0", fgCashMemo.TextMatrix(j, 5)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 6) = "", "0", fgCashMemo.TextMatrix(j, 6)) & ", " & _
            IIf(fgCashMemo.TextMatrix(j, 7) = "", "0", fgCashMemo.TextMatrix(j, 7)) & ", " & _
            " '" & parseQuotes(fgCashMemo.TextMatrix(j, 8)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 9)) & "', " & _
            " '" & Format(BillDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtCActive) & "', " & _
            " '" & parseQuotes(txtCPost) & "'," & _
            IIf(blnNoDiscount, 1, 0) & "," & _
            IIf(blnNoVat, 1, 0) & ")"
'                           "('" & IIf(fgCashMemo.TextMatrix(j, 13), "1", "0") & "'))"
                           
                           Next
                
        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record is Posting Now", vbInformation, "Confirmation"
        
        Else

        rcupdate = True
        
        End If
        Call ActiveTable
'   ------------------------End of Post Information-------------------------- -2147168237
    
    End If
    
    Exit Function

End Function

Private Sub FindRecord()
If Not rscashmaster.BOF Then

    Dim i As Integer
    Dim strLedgerDetail As String
    Set rsCashDetail = New ADODB.Recordset
    txtBillSerialNo = rscashmaster!SerialNo
    BillDate = rscashmaster!strDate
    cmbTName = rscashmaster!TableName
    cmbWaiter = rscashmaster!WaiterName
    txtPersonNo = rscashmaster!Guest
    txtServiceCharge = rscashmaster!ServiceCharge
    txtVAT = rscashmaster!Vat
    txtDiscount = rscashmaster!Discount
    txtDiscountAmt = rscashmaster!DiscountAmt
    txtTime = rscashmaster!strTime
'    cmbDCard = rscashmaster!DCard
    txtGName = rscashmaster!GName
    txtCAddress = rscashmaster!Address

    txtTotalBill = rscashmaster!TotalBill
    txtTotalVat = rscashmaster!TotalVat
    txtTotalSCharge = rscashmaster!TSCharge
    txtTotalDiscount = rscashmaster!TotalDiscount
    txtNPayable = rscashmaster!NetPayable
    cboMode = rscashmaster!PaymentMode
    txtRemarks = rscashmaster!Remarks
    txtPaid = rscashmaster!Paid
    txtDue = rscashmaster!Due
    cmbATable = rscashmaster!Post
    chkActive = rscashmaster!Active
    txtCActive = rscashmaster!CActive
    txtCPost = rscashmaster!CPost
    txtUName = rscashmaster!UName
    txtKOT = IIf(IsNull(rscashmaster!Kot), "", rscashmaster!Kot)
    txtBOT = IIf(IsNull(rscashmaster!Bot), "", rscashmaster!Bot)
    AdDate.Value = rscashmaster!ADate
    txtAdvance = rscashmaster!Advance
    txtCashReceived = rscashmaster!CReceived
    txtCashBack = rscashmaster!CashBack
    txtMR_No = rscashmaster!MrNo

     
fgCashMemo.Rows = 1
strLedgerDetail = "SELECT  BillSerialNo,SerialNo,ItemCode, ItemName, Qty, Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT" & _
                " FROM tblCashDetail " & _
                "WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo.text) & "'"
    
rsCashDetail.CursorLocation = adUseClient
rsCashDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsCashDetail.RecordCount <> 0 Then

        fgCashMemo.Rows = rsCashDetail.RecordCount + 1
                i = 0
        For i = 1 To rsCashDetail.RecordCount
            fgCashMemo.TextMatrix(i, 1) = rsCashDetail("BillSerialNo")
            fgCashMemo.TextMatrix(i, 2) = rsCashDetail("SerialNo")
            fgCashMemo.TextMatrix(i, 3) = rsCashDetail("ItemCode")
            fgCashMemo.TextMatrix(i, 4) = rsCashDetail("ItemName")
            fgCashMemo.TextMatrix(i, 5) = rsCashDetail("Qty")
            fgCashMemo.TextMatrix(i, 6) = rsCashDetail("Rate")
            fgCashMemo.TextMatrix(i, 7) = rsCashDetail("Tips")
            fgCashMemo.TextMatrix(i, 8) = rsCashDetail("ItemGroup")
            fgCashMemo.TextMatrix(i, 9) = rsCashDetail("ItemCatagory")
            fgCashMemo.TextMatrix(i, 10) = rsCashDetail("strDate")
            fgCashMemo.TextMatrix(i, 11) = rsCashDetail("CActive")
            fgCashMemo.TextMatrix(i, 12) = rsCashDetail("CPost")
            fgCashMemo.TextMatrix(i, 13) = rsCashDetail("NoDiscount")
            fgCashMemo.TextMatrix(i, 14) = rsCashDetail("NoVAT")
            
            
            
    rsCashDetail.MoveNext
        Next
      End If
    rsCashDetail.Close
    End If
End Sub

Private Sub FindRecord1()

    Dim i As Integer
    Dim strLedgerDetail1 As String
    
'    Set rsCashDetail1 = New ADODB.Recordset
    txtBillSerialNo = rsAMaster!SerialNo
    BillDate = rsAMaster!strDate
    cmbTName = rsAMaster!TableName
    cmbWaiter = rsAMaster!WaiterName
    txtPersonNo = rsAMaster!Guest
    txtServiceCharge = rsAMaster!ServiceCharge
    txtVAT = rsAMaster!Vat
    txtDiscount = rsAMaster!Discount
    txtDiscountAmt = rsAMaster!DiscountAmt
    cmbDCard = rsAMaster!DCard
    txtGName = rsAMaster!GName
    txtCAddress = rsAMaster!Address

    txtTotalBill = rsAMaster!TotalBill
    txtTotalVat = rsAMaster!TotalVat
    txtTotalSCharge = rsAMaster!TSCharge
    txtTotalDiscount = rsAMaster!TotalDiscount
    txtNPayable = rsAMaster!NetPayable
    cboMode = rsAMaster!PaymentMode
    txtRemarks = rsAMaster!Remarks
    txtPaid = rsAMaster!Paid
    txtDue = rsAMaster!Due
    txtCActive = rsAMaster!CActive
    txtCPost = rsAMaster!CPost
    txtUName = rsAMaster!UName
    '    txtKOT = rsCashMaster!Kot
    txtKOT = IIf(IsNull(rsAMaster!Kot), "", rsAMaster!Kot)
'    txtBOT = rsCashMaster!Bot
    txtBOT = IIf(IsNull(rsAMaster!Bot), "", rsAMaster!Bot)
    AdDate.Value = rsAMaster!ADate
    txtAdvance = rsAMaster!Advance
    txtMR_No = rsAMaster!MrNo
    
    
    
    Set rsCashDetail1 = New ADODB.Recordset
    fgCashMemo.Rows = 1
    If rsCashDetail1.State <> 0 Then rsCashDetail1.Close
    strLedgerDetail1 = "SELECT  BillSerialNo,SerialNo,ItemCode, ItemName,Qty, Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT " & _
                " FROM tblCashDetail " & _
                "WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo.text) & "'"
    
    rsCashDetail1.CursorLocation = adUseClient
    rsCashDetail1.Open strLedgerDetail1, cn, adOpenStatic, adLockReadOnly

 If rsCashDetail1.RecordCount <> 0 Then
          

        fgCashMemo.Rows = rsCashDetail1.RecordCount + 1
                i = 0
        For i = 1 To rsCashDetail1.RecordCount
            fgCashMemo.TextMatrix(i, 1) = rsCashDetail1("BillSerialNo")
            fgCashMemo.TextMatrix(i, 2) = rsCashDetail1("SerialNo")
            fgCashMemo.TextMatrix(i, 3) = rsCashDetail1("ItemCode")
            fgCashMemo.TextMatrix(i, 4) = rsCashDetail1("ItemName")
            fgCashMemo.TextMatrix(i, 5) = rsCashDetail1("Qty")
            fgCashMemo.TextMatrix(i, 6) = rsCashDetail1("Rate")
            fgCashMemo.TextMatrix(i, 7) = rsCashDetail1("Tips")
            fgCashMemo.TextMatrix(i, 8) = rsCashDetail1("ItemGroup")
            fgCashMemo.TextMatrix(i, 9) = rsCashDetail1("ItemCatagory")
            fgCashMemo.TextMatrix(i, 10) = rsCashDetail1("strDate")
            fgCashMemo.TextMatrix(i, 11) = rsCashDetail1("CActive")
            fgCashMemo.TextMatrix(i, 12) = rsCashDetail1("CPost")
            fgCashMemo.TextMatrix(i, 13) = rsCashDetail1("NoDiscount")
            fgCashMemo.TextMatrix(i, 14) = rsCashDetail1("NoVAT")
            
            
    rsCashDetail1.MoveNext
        Next
      End If
    rsCashDetail1.Close
End Sub

Private Sub Clear()
    cmbTName.text = ""
    cmbWaiter.text = ""
    txtPersonNo.text = ""
    txtServiceCharge.text = "10"
    txtVAT.text = "15"
    txtDiscount.text = "0"
    txtDiscountAmt.text = "0"
    cmbDCard.text = ""
    txtGName.text = ""
    txtCAddress.text = ""
    txtRemarks.text = ""
    txtPaid.text = ""
    txtDue.text = ""
    txtTotalBill.text = "0"
    txtTotalVat.text = "0"
    txtTotalSCharge.text = "0"
    txtTotalDiscount.text = "0"
    txtNPayable.text = "0"
    txtPaid.text = "0"
    txtDue.text = "0"
    txtKOT = "0"
    txtBOT = "0"
    txtTItem.text = ""
    txtAdvance = "0"
    txtMR_No = "0"
    txtCashReceived.text = "0"
    txtCashBack.text = "0"
    
    txtQty.text = ""
    txtItemCode.text = ""
    txtItemName.text = ""
    txtRate.text = ""
    txtTips.text = ""
    txtItemCatagory.text = ""
    txtItemGroup.text = ""
    txtNoDiscount.text = ""
    txtNoVAT.text = ""
    
End Sub

Private Sub allenable()
    cmbTName.Enabled = True
    cmbWaiter.Enabled = True
    txtPersonNo.Enabled = True
    txtServiceCharge.Enabled = True
    txtVAT.Enabled = True
    txtDiscount.Enabled = True
    txtDiscountAmt.Enabled = True
    cmbDCard.Enabled = True
    txtGName.Enabled = True
    txtCAddress.Enabled = True
    txtRemarks.Enabled = True
    txtPaid.Enabled = True
    txtDue.Enabled = True
    BillDate.Enabled = True
    cboMode.Enabled = True
'    fgCashMemo.Editable = flexEDKbdMouse
     txtKOT.Enabled = True
     txtBOT.Enabled = True
     txtAdvance.Enabled = True
     txtMR_No.Enabled = True
     txtTItem.Enabled = True
     
     txtQty.Enabled = True
    txtItemCode.Enabled = True
    txtItemName.Enabled = True
    txtRate.Enabled = True
    txtTips.Enabled = True
    txtItemCatagory.Enabled = True
    txtItemGroup.Enabled = True
    txtNoDiscount.Enabled = True
    txtNoVAT.Enabled = True
    End Sub

Private Sub alldisable()
    cmbTName.Enabled = False
    cmbWaiter.Enabled = False
    txtPersonNo.Enabled = False
    txtServiceCharge.Enabled = False
    txtVAT.Enabled = False
    txtDiscount.Enabled = False
    txtDiscountAmt.Enabled = False
    cmbDCard.Enabled = False
    txtCAddress.Enabled = False
    txtGName.Enabled = False
    txtTime.Enabled = False
    txtDue.Enabled = False
    BillDate.Enabled = False
    AdDate.Enabled = False

     txtKOT.Enabled = False
     txtBOT.Enabled = False
     txtAdvance.Enabled = False
     txtMR_No.Enabled = False
     txtTItem.Enabled = False
    
    txtQty.Enabled = False
    txtItemCode.Enabled = False
    txtItemName.Enabled = False
    txtRate.Enabled = False
    txtTips.Enabled = False
    txtItemCatagory.Enabled = False
    txtItemGroup.Enabled = False
    txtNoDiscount.Enabled = False
    txtNoVAT.Enabled = False
End Sub

Private Sub duplicate()

Dim j As Integer
        
         For j = 1 To fgCashMemo.Rows - 2
        
        If Val(fgCashMemo.TextMatrix(j, 3)) = Val(fgCashMemo.TextMatrix(j + 1, 3)) Then
        MsgBox "Duplicate Item Code Number.", vbInformation
         fgCashMemo.TextMatrix(j, 3) = ""
         End If

Next

End Sub
''------------------------------Add for Reporting Purpose------------------------
Public Sub PopulateForm(StrID As String)

 Set rscashmaster = New ADODB.Recordset

    If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "select SerialNo,strDate,TableName,WaiterName,Guest,ServiceCharge,Vat,Discount,DiscountAmt,GName,Address,strTime, " & _
                          "TotalBill,TotalVat,TSCharge,TotalDiscount,NetPayable,PaymentMode,Remarks," & _
                          "Paid,Due,Active,Post,CActive,CPost,UName,Kot,Bot,ADate,Advance,CReceived,CashBack,MrNo from tblCashMaster", cn, adOpenStatic, adLockReadOnly
                          

        rscashmaster.MoveFirst
'End If
    
    rscashmaster.Find "SerialNo=" & parseQuotes(StrID)
    If rscashmaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


Public Sub CashCopy()
On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Set rsDailyRpt = New ADODB.Recordset


If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "select SerialNo,strDate,TableName,WaiterName,Guest,ServiceCharge,Vat,Discount,DiscountAmt,DCard,GName,Address,strTime, " & _
                          "TotalBill,TotalVat,TSCharge,TotalDiscount,NetPayable,PaymentMode,Remarks," & _
                          "Paid,Due,Active,Post,CActive,CPost,UName,BOT,KOT,Advance,CReceived,CashBack,MrNo from tblCashMaster", cn, adOpenStatic, adLockReadOnly



    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If


        strPath = App.Path + "\reports\CashMemo Cash.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields



'    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close

 rsDailyRpt.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate,tblCashMaster.TableName," & _
                "tblCashMaster.WaiterName,tblCashMaster.Guest,tblCashMaster.ServiceCharge," & _
                "tblCashMaster.Vat,tblCashMaster.Discount,tblCashMaster.GName," & _
                "tblCashMaster.Address, tblCashMaster.strTime,tblCashMaster.PaymentMode,tblCashMaster.Remarks," & _
                "tblCashMaster.NetPayable,tblCashDetail.ItemCode, tblCashDetail.ItemName," & _
                "tblCashDetail.Qty,tblCashDetail.Rate,((tblCashDetail.Qty)*(tblCashDetail.Rate))as Amount,tblCashMaster.TotalVat, " & _
                "tblCashMaster.TSCharge,tblCashMaster.TotalDiscount,tblCashMaster.Due, " & _
                "tblCashMaster.Paid,tblCashMaster.CActive,tblCashMaster.CPost,tblCashMaster.UName, " & _
                "tblCashMaster.CReceived,tblCashMaster.CashBack FROM  tblCashMaster,tblCashDetail Where tblCashMaster.SerialNo=tblCashDetail.BillSerialNo " & _
                "AND tblCashMaster.SerialNo ='" & txtBillSerialNo.text & "' order by tblCashDetail.ItemCode", cn, adOpenStatic


'rsDailyRpt.Open "SELECT tblCashMaster.SerialNo, tblCashMaster.strDate, tblCashMaster.TableName, tblCashMaster.WaiterName, tblCashMaster.Guest," & _
'                          "tblCashMaster.ServiceCharge, tblCashMaster.Vat, tblCashMaster.Discount, tblCashMaster.DiscountAmt, tblCashMaster.DCard," & _
'                          "tblCashMaster.GName, tblCashMaster.Address, tblCashMaster.strTime, tblCashMaster.TotalBill, tblCashMaster.TotalVat," & _
'                          "tblCashMaster.TSCharge, tblCashMaster.TotalDiscount, tblCashMaster.NetPayable, tblCashMaster.PaymentMode," & _
'                          "tblCashMaster.Remarks, tblCashMaster.Paid, tblCashMaster.Due, tblCashDetail.ItemCode, tblCashDetail.ItemName," & _
'                          "tblCashDetail.Qty, tblCashDetail.Rate, ((tblCashDetail.Qty) * (tblCashDetail.Rate)) AS Amout, tblCashMaster.CPost," & _
'                          "tblCashMaster.UName,tblCashMaster.Advance FROM  tblCashMaster,tblCashDetail Where tblCashMaster.SerialNo=tblCashDetail.BillSerialNo " & _
'                          "AND tblCashMaster.SerialNo ='" & txtBillSerialNo.text & "' order by tblCashDetail.ItemCode", cn, adOpenStatic

        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"

            Set objReportFF = objReportFormulaFieldDefinations.Item(2)
            objReportFF.text = "'" + parseQuotes(txtUName.text) + " '"

'-------------End Add Discunt-------------------
        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt

        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True

        objReport.DiscardSavedData
        If Tracer = 0 Then
        objReport.Preview "Cash Memo Report", , , , , 16777216 Or 524288 Or 65536
        Else
        objReport.PrintOut (False)
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

Public Sub GuestCopy()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Set rsDailyRpt = New ADODB.Recordset


If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "select SerialNo,strDate,TableName,WaiterName,Guest,ServiceCharge,Vat,Discount,DiscountAmt,DCard,GName,Address,strTime, " & _
                          "TotalBill,TotalVat,TSCharge,TotalDiscount,NetPayable,PaymentMode,Remarks," & _
                          "Paid,Due,Active,Post,CActive,CPost,UName,BOT,KOT,Advance from tblCashMaster", cn, adOpenStatic, adLockReadOnly



    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If


        strPath = App.Path + "\reports\CashMemo Guest.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields



'    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close

rsDailyRpt.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate,tblCashMaster.TableName," & _
                "tblCashMaster.WaiterName,tblCashMaster.Guest,tblCashMaster.ServiceCharge," & _
                "tblCashMaster.Vat,tblCashMaster.Discount,tblCashMaster.GName," & _
                "tblCashMaster.Address, tblCashMaster.strTime,tblCashMaster.PaymentMode,tblCashMaster.Remarks," & _
                "tblCashMaster.NetPayable,tblCashDetail.ItemCode, tblCashDetail.ItemName," & _
                "tblCashDetail.Qty,tblCashDetail.Rate,((tblCashDetail.Qty)*(tblCashDetail.Rate))as Amount,tblCashMaster.TotalVat, " & _
                "tblCashMaster.TSCharge,tblCashMaster.TotalDiscount,tblCashMaster.Due, " & _
                "tblCashMaster.Paid,tblCashMaster.CActive,tblCashMaster.CPost,tblCashMaster.UName, " & _
                "tblCashMaster.CReceived,tblCashMaster.CashBack FROM  tblCashMaster,tblCashDetail Where tblCashMaster.SerialNo=tblCashDetail.BillSerialNo " & _
                "AND tblCashMaster.SerialNo ='" & txtBillSerialNo.text & "' order by tblCashDetail.ItemCode", cn, adOpenStatic

'rsDailyRpt.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate,tblCashMaster.TableName," & _
'                          "tblCashMaster.WaiterName,tblCashMaster.Guest," & _
'                          "tblCashMaster.ServiceCharge,tblCashMaster.Vat,tblCashMaster.Discount," & _
'                          "tblCashMaster.GName, tblCashMaster.Address,tblCashMaster.PaymentMode," & _
'                          "tblCashMaster.NetPayable,tblCashDetail.ItemCode, tblCashDetail.ItemName," & _
'                          "tblCashDetail.Qty,tblCashDetail.Rate,((tblCashDetail.Qty)*(tblCashDetail.Rate))as Amount,tblCashMaster.TotalVat, " & _
'                          "tblCashMaster.TSCharge,tblCashMaster.TotalDiscount,tblCashMaster.Due FROM  tblCashMaster,tblCashDetail Where tblCashMaster.SerialNo=tblCashDetail.BillSerialNo " & _
'                          "AND tblCashMaster.SerialNo ='" & txtBillSerialNo.text & "' order by tblCashDetail.ItemCode", cn, adOpenStatic


        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"

'-------------End Add Discunt-------------------
        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt

        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True

        objReport.DiscardSavedData
        If Tracer = 0 Then
        objReport.Preview "CashMemo Cash Copy", , , , , 16777216 Or 524288 Or 65536
        Else
        objReport.PrintOut (False)
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

'-----------------This code is only used for calculate in word purpos--------------------
'Option Base 0
Function InWords(ByVal GetAmount As Variant) As String
On Error GoTo Kick_Errors 'if error goto Kick_Errors Labels
'Declare some necessary variable
Dim tempNum As Integer
Dim getTaka As Variant
Dim getPaisa As Integer, getPaisaainWords As String
Dim AmountinWords As String
Dim Arrindex As Integer
Static NumInWord1 As Variant
'Check whether getAmount contain valid number
If Not IsNumeric(GetAmount) Then Exit Function

'Check whether getAmount>999999999.99
If GetAmount > 999999999.99 Then Exit Function

'array for thousand and million only that calculate here
NumInWord1 = Array(" ", "Thousand", "Million")

GetAmount = Abs(GetAmount)              'make positive
getTaka = Int(GetAmount)                'get taka part
getPaisa = (GetAmount - getTaka) * 100  'get taka part
If getTaka > 0 Then         'if there is taka,
                            'the following Loop use to get
                            'hundreds,thousands, then millions.

Do
    tempNum = getTaka Mod 1000
    getTaka = Int(getTaka / 1000)
'Set output
If tempNum <> 0 Then
    AmountinWords = GetAmWords(tempNum) & " " & _
                    NumInWord1(Arrindex) & " " & AmountinWords
    End If
    Arrindex = Arrindex + 1
Loop While getTaka > 0
If getPaisa > 0 Then
    getPaisaainWords = GetAmWords(getPaisa)
    AmountinWords = RTrim(AmountinWords) & _
                    "Taka and" & getPaisaainWords & " Paisa Only"
Else
    AmountinWords = RTrim(AmountinWords) & " Taka Only."
    End If
End If
getOut: 'label getOut
    InWords = AmountinWords
    Exit Function

Kick_Errors: 'label Kick_Errors
    'If text box contain wrong data just return empty string
    AmountinWords = " "
    Resume getOut
End Function

Function GetAmWords(ByVal GetAmount As Integer) As String
    Static UnitOnes As Variant
    Static UnitTens As Variant
    Dim AmountinWords As String
    Dim getNumDigit As Integer
    'Set UnitOnes if have no elements
If IsEmpty(UnitOnes) Then
    UnitOnes = Array(" ", "One", "Two", "Three", "Four", _
                     "Five", "Six", "Seven", "Eight", "Nine", "Ten", _
                     "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", _
                     "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty")

End If

'What about others
If IsEmpty(UnitTens) Then
    UnitTens = Array(" ", " ", "Twenty", "Thirty", "Forty", "Fifty", _
                     "Sixty", "Seventy", "Eighty", "Ninety")

End If

'Calculate hundreds and rest value
getNumDigit = GetAmount \ 100
GetAmount = GetAmount Mod 100
'If hundred found
If getNumDigit > 0 Then
    AmountinWords = UnitOnes(getNumDigit) & "  Hundred"
End If
'Select Word for Unit Ones and Tens
Select Case GetAmount
    Case 1 To 20 'get from UnitOnes array
            AmountinWords = AmountinWords & _
                            " " & UnitOnes(GetAmount)

    Case 21 To 99 'get from UnitOnes array
            getNumDigit = GetAmount \ 10
            GetAmount = GetAmount Mod 10

 If getNumDigit > 0 Then
    AmountinWords = AmountinWords & _
                    " " & UnitTens(getNumDigit)

 End If

 If GetAmount > 0 Then
    AmountinWords = AmountinWords & _
                    " " & UnitOnes(GetAmount)
    End If
 End Select
    GetAmWords = AmountinWords
 End Function

Private Sub txtAdvance_Change()
Call Calculation
End Sub

Private Sub txtCashBack_Change()
'CDbl(txtCashBack) = CDbl(txtCashReceived) - CDbl(txtDue)
End Sub


Private Sub txtCashReceived_Change()
txtCashBack.text = Val(txtCashReceived) - Val(txtNPayable)
End Sub

Private Sub txtDiscount_Change()
Dim i As Integer
Call Calculation
'End If
End Sub

Private Sub txtDiscountAmt_Change()
Dim i As Integer
Call Calculation
'End If
End Sub


'========================Direct Data Through in Data Grid===========================================================
Private Sub txtItemCode_Change()
Set rsCustomerMaster = New ADODB.Recordset
    
    If rsCustomerMaster.State <> 0 Then rsCustomerMaster.Close
       rsCustomerMaster.Open "select SerialNo,ItemCode,ItemName,ItemQty,ItemPrice,Tips,NoDiscount,NoVAT,ItemGroup,ItemCatagory from tblItemDetail where ItemCode ='" & txtItemCode & "' ", cn, adOpenStatic, adLockReadOnly

   If rsCustomerMaster.RecordCount > 0 Then
      rsCustomerMaster.MoveFirst
    End If
    
    If Not rsCustomerMaster.EOF Then FindRecord3
cmdFind1_Click

End Sub

Private Sub FindRecord3()
    txtItemName = rsCustomerMaster!ItemName
    txtQty = rsCustomerMaster!ItemQty
    txtRate = rsCustomerMaster!ItemPrice
    txtTips = rsCustomerMaster!Tips
    txtNoDiscount = rsCustomerMaster!NoDiscount
    txtNoVAT = rsCustomerMaster!NoVAT
    txtItemGroup = rsCustomerMaster!ItemGroup
    txtItemCatagory = rsCustomerMaster!ItemCatagory
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If

End Sub

Private Sub txtItemCode_LostFocus()
Set rsCustomerMaster = New ADODB.Recordset
    
    If rsCustomerMaster.State <> 0 Then rsCustomerMaster.Close
       rsCustomerMaster.Open "select SerialNo,ItemCode,ItemName,ItemQty,ItemPrice,Tips,NoDiscount,NoVAT,ItemGroup,ItemCatagory from tblItemDetail where ItemCode ='" & txtItemCode & "' ", cn, adOpenStatic, adLockReadOnly

   If rsCustomerMaster.RecordCount > 0 Then
      rsCustomerMaster.MoveFirst
    End If
      
    If Not rsCustomerMaster.EOF Then FindRecord3

End Sub

Private Sub txtQty_Change()
cmdFind1_Click
End Sub

Private Sub cmdFind1_Click()
  
        If rsTemp.State <> 0 Then rsTemp.Close
              
If txtItemCode.text <> "" Then

rsTemp.Open "SELECT TOP 50 SerialNo,ItemCode,ItemName,ItemQty='" & txtQty & "',ItemPrice,Tips,NoDiscount,NoVAT,ItemGroup,ItemCatagory " & _
                "FROM tblItemDetail WHERE tblItemDetail.ItemCode= '" & parseQuotes(txtItemCode.text) & "'", cn, adOpenStatic, adLockReadOnly
        
fgExport.Rows = 1
    
    While Not rsTemp.EOF

 fgExport.AddItem "" & vbTab & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ItemCode") & _
         vbTab & rsTemp("ItemName") & vbTab & rsTemp("ItemQty") & vbTab & rsTemp("ItemPrice") & vbTab & rsTemp("Tips") & vbTab & rsTemp("NoDiscount") & vbTab & rsTemp("NoVAT") & vbTab & rsTemp("ItemGroup") & vbTab & rsTemp("ItemCatagory")
        rsTemp.MoveNext

        Wend
        
        If fgExport.Rows = 0 Then fgExport.AddItem ""
        On Error Resume Next
        If fgExport.Rows <= 1 Then Exit Sub
Dim i As Integer
Dim j As Integer
    For i = 0 To fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next
        
GridCount fgExport
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyBack Then
txtQty.text = ""
End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If KeyAscii = 13 And txtQty = "" Then
SendKeys Chr(9)
Call cmdOk_Click
End If
   
   If txtQty <> "" Then
 
   Dim iRows As Integer
   Dim i As Integer
   Dim temp As Double

    temp = 0
    iRows = fgExport.Rows
    If fgExport.Rows <= 1 Then Exit Sub
    
    For i = 1 To iRows - 1
fgExport.Cell(flexcpChecked, i, 1) = flexChecked
fgCashMemo.AddItem "" & vbTab & vbTab & fgExport.TextMatrix(i, 2) & vbTab & fgExport.TextMatrix(i, 3) & _
        vbTab & fgExport.TextMatrix(i, 4) & vbTab & fgExport.TextMatrix(i, 5) & vbTab & fgExport.TextMatrix(i, 6) & vbTab & fgExport.TextMatrix(i, 7) & _
        vbTab & fgExport.TextMatrix(i, 10) & vbTab & fgExport.TextMatrix(i, 11) & _
        vbTab & vbTab & vbTab & vbTab & fgExport.TextMatrix(i, 8) & vbTab & fgExport.TextMatrix(i, 9)

txtQty.text = ""
txtItemCode.text = ""
txtItemName.text = ""
txtRate.text = ""
txtTips.text = ""
txtItemCatagory.text = ""
txtItemGroup.text = ""
txtNoDiscount.text = ""
txtNoVAT.text = ""

txtItemCode.SetFocus
Next
End If
End If
Call deleteRow
Call Calculation
End Sub

Private Sub deleteRow()
If fgExport.Rows = 1 Then fgExport.AddItem ""
        On Error Resume Next
        If fgExport.Rows <= 1 Then Exit Sub
Dim i As Integer
Dim j As Integer
    For i = 0 To fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next
End Sub

Private Sub txtQty_GotFocus()
txtQty.SelStart = 0
txtQty.SelLength = Len(txtQty)
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
    For i = 0 To fgCashMemo.Rows - 1
             j = 1
             For j = 1 To fgExport.Rows - 1

                If fgCashMemo.TextMatrix(i, 3) = fgExport.TextMatrix(j, 3) Then
                    fgExport.RemoveItem j
                End If
             Next
    Next
    
'    ---------------------
'Unload Me

Set frmItemGroupSearch = Nothing

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
                    vbTab & fgExport.TextMatrix(i, 7) & vbTab & fgExport.TextMatrix(i, 9) & vbTab & fgExport.TextMatrix(i, 10) & vbTab & vbTab & vbTab & vbTab & (fgExport.TextMatrix(i, 8))
                     
        End If
        
    Next

End Sub

'========================End Data Through in Data Grid===========================================================

''-----------------End of calculate in word purpose---------------------------------------
'Private Sub cmdCalculate_Click()
'txtWords.text = InWords(txtSumBD.text)
'End Sub
''-----------------------------------------------------------------------------------------
'Private Sub txtPaid_Change()
'Call txtPaid_click
'
'End Sub


Private Sub txtPaid_GotFocus()
Call txtPaid_click

End Sub


Private Sub txtPaid_click()
Dim ipaid
Dim j As Integer
Dim blnNoDiscount, blnNoVat    As Boolean
'On Error GoTo ErrH
If txtCPost.text = "Posted" And txtCActive.text = "Active" Then
       If Val(txtPaid) = 0 Then
                        txtPaid.Enabled = True
                        ipaid = MsgBox("Do you want to full payment This bill?", vbYesNo)
       
              If ipaid = vbYes Then
                          txtPaid = CDbl(Val(txtDue))
                          txtDue = 0
'                          cn.BeginTrans
            
   cn.Execute "UPDATE tblCashMaster SET strDate='" & Format(BillDate, "dd-mmm-yyyy") & "',TableName='" & parseQuotes(cmbTName) & "', " & _
                " WaiterName='" & parseQuotes(cmbWaiter) & "',Guest=" & Val(txtPersonNo.text) & ",ServiceCharge=" & Val(txtServiceCharge.text) & ", " & _
                "Vat=" & Val(txtVAT.text) & ",Discount=" & Val(txtDiscount.text) & ",DiscountAmt=" & Val(txtDiscountAmt.text) & ",DCard='" & parseQuotes(cmbDCard.text) & "'," & _
                "GName='" & parseQuotes(txtGName) & "',Address='" & parseQuotes(txtCAddress) & "',strTime='" & (txtTime.text) & "',TotalBill=" & Val(txtTotalBill.text) & ", " & _
                "TotalVat=" & Val(txtTotalVat.text) & ",TSCharge=" & Val(txtTotalSCharge.text) & ",TotalDiscount=" & Val(txtTotalDiscount.text) & ", " & _
                "NetPayable=" & Val(txtNPayable.text) & ",PaymentMode='" & parseQuotes(cboMode) & "', " & _
                "Remarks='" & parseQuotes(txtRemarks) & "',Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ",Post='" & parseQuotes(cmbATable) & "',Active=" & chkActive & ", " & _
                "CActive='" & parseQuotes(txtCActive.text) & "',CPost='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "', Kot='" & txtKOT.text & "',Bot='" & txtBOT.text & "', " & _
                "ADate='" & Format(AdDate, "dd-mmm-yyyy") & "',Advance=" & Val(txtAdvance.text) & ",CReceived=" & Val(txtCashReceived.text) & ",CashBack=" & Val(txtCashBack.text) & ",MrNo='" & txtMR_No.text & "' WHERE SerialNo = '" & txtBillSerialNo & "'"

    
                   cn.Execute "DELETE FROM tblCashDetail WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo) & "'"
        
                        j = 0
                                For j = 1 To fgCashMemo.Rows - 1
                                
            If fgCashMemo.Cell(flexcpChecked, j, 13) = flexChecked Then
               blnNoDiscount = True
            Else
                blnNoDiscount = False
            End If
            
            
            If fgCashMemo.Cell(flexcpChecked, j, 14) = flexChecked Then
               blnNoVat = True
            Else
                blnNoVat = False
            End If
                    
cn.Execute "INSERT INTO tblCashDetail (BillSerialNo,SerialNo,ItemCode,ItemName,Qty,Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT) " & _
           "Values ('" & parseQuotes(txtBillSerialNo) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 2)) & "'," & _
           " '" & parseQuotes(fgCashMemo.TextMatrix(j, 3)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 4)) & "'," & _
           IIf(fgCashMemo.TextMatrix(j, 5) = "", "0", fgCashMemo.TextMatrix(j, 5)) & ", " & _
           IIf(fgCashMemo.TextMatrix(j, 6) = "", "0", fgCashMemo.TextMatrix(j, 6)) & ", " & _
           IIf(fgCashMemo.TextMatrix(j, 7) = "", "0", fgCashMemo.TextMatrix(j, 7)) & ", " & _
           " '" & parseQuotes(fgCashMemo.TextMatrix(j, 8)) & "','" & parseQuotes(fgCashMemo.TextMatrix(j, 9)) & "', " & _
           " '" & Format(BillDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtCActive) & "', " & _
           " '" & parseQuotes(txtCPost) & "'," & _
           IIf(blnNoDiscount, 1, 0) & ", " & _
            IIf(blnNoVat, 1, 0) & ")"
                       
                       Next
                
                
               '        rcupdate = True
'                        cn.CommitTrans
                        MsgBox "Bill is Paid Now", vbInformation, "Confirmation"
                        cboMode.Enabled = False
              End If


        Else
           txtPaid.Enabled = True
           
        '   cboMode.Enabled = True
        End If
Else
  If txtCPost.text <> "Posted" Then
  MsgBox "Please Post The bill First .", vbInformation, "Confirmation"
  Else
  MsgBox "You Can't paid because the bill is Void .", vbInformation, "Confirmation"
'ErrH:
  End If
End If

End Sub


Private Sub FindRecord11()

    Dim i As Integer
    Dim strLedgerDetail As String
    Set rsCashDetail = New ADODB.Recordset
    txtBillSerialNo = rscashmaster!SerialNo
    BillDate = rscashmaster!strDate
    cmbTName = rscashmaster!TableName
    cmbWaiter = rscashmaster!WaiterName
    txtPersonNo = rscashmaster!Guest
    txtServiceCharge = rscashmaster!ServiceCharge
    txtVAT = rscashmaster!Vat
    txtDiscount = rscashmaster!Discount
    txtDiscountAmt = rscashmaster!DiscountAmt
    cmbDCard = rscashmaster!DCard
    txtGName = rscashmaster!GName
    txtCAddress = rscashmaster!Address

    txtTotalBill = rscashmaster!TotalBill
    txtTotalVat = rscashmaster!TotalVat
    txtTotalSCharge = rscashmaster!TSCharge
    txtTotalDiscount = rscashmaster!TotalDiscount
    txtNPayable = rscashmaster!NetPayable
    cboMode = rscashmaster!PaymentMode
    txtRemarks = rscashmaster!Remarks
    txtPaid = rscashmaster!Paid
    txtDue = rscashmaster!Due
    cmbATable = rscashmaster!Post
    chkActive = rscashmaster!Active
    txtCActive = rscashmaster!CActive
    txtCPost = rscashmaster!CPost
'    txtUName = rsCashMaster!UName
    txtUName = IIf(IsNull(rscashmaster!UName), "", rscashmaster!UName)
    '    txtKOT = rsCashMaster!Kot
    txtKOT = IIf(IsNull(rscashmaster!Kot), "", rscashmaster!Kot)
'    txtBOT = rsCashMaster!Bot
    txtBOT = IIf(IsNull(rscashmaster!Bot), "", rscashmaster!Bot)
    AdDate.Value = rscashmaster!ADate
    txtAdvance = rscashmaster!Advance
    txtMR_No = rscashmaster!MrNo
    
    
    
    
    fgCashMemo.Rows = 1
    strLedgerDetail = "SELECT  BillSerialNo,SerialNo,ItemCode, ItemName, Qty, Rate,Tips,ItemGroup,ItemCatagory,strDate,CActive,CPost,NoDiscount,NoVAT" & _
                " FROM tblCashDetail " & _
                "WHERE BillSerialNo='" & parseQuotes(txtBillSerialNo.text) & "'"
    rsCashDetail.CursorLocation = adUseClient
    rsCashDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsCashDetail.RecordCount <> 0 Then

        fgCashMemo.Rows = rsCashDetail.RecordCount + 1
                i = 0
        For i = 1 To rsCashDetail.RecordCount
            fgCashMemo.TextMatrix(i, 1) = rsCashDetail("BillSerialNo")
            fgCashMemo.TextMatrix(i, 2) = rsCashDetail("SerialNo")
            fgCashMemo.TextMatrix(i, 3) = rsCashDetail("ItemCode")
            fgCashMemo.TextMatrix(i, 4) = rsCashDetail("ItemName")
            fgCashMemo.TextMatrix(i, 5) = rsCashDetail("Qty")
            fgCashMemo.TextMatrix(i, 6) = rsCashDetail("Rate")
            fgCashMemo.TextMatrix(i, 7) = rsCashDetail("Tips")
            fgCashMemo.TextMatrix(i, 8) = rsCashDetail("ItemGroup")
            fgCashMemo.TextMatrix(i, 9) = rsCashDetail("ItemCatagory")
            fgCashMemo.TextMatrix(i, 10) = rsCashDetail("strDate")
            fgCashMemo.TextMatrix(i, 11) = rsCashDetail("CActive")
            fgCashMemo.TextMatrix(i, 12) = rsCashDetail("CPost")
            fgCashMemo.TextMatrix(i, 13) = rsCashDetail("NoDiscount")
            fgCashMemo.TextMatrix(i, 14) = rsCashDetail("NoVAT")
            
            
    rsCashDetail.MoveNext
        Next
      End If
    rsCashDetail.Close
End Sub


Private Sub Timer1_Timer()
    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub


Private Sub changeVisible()
Dim str As String
Set rs = New ADODB.Recordset
str = "select UID,UPassword,Upper(UID)as Name  from RMSUser where UID ='" & frmLogin.txtUID.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
'           If rs!Name = "ADMIN" And rs!UPassword = "VB5DELPHI2" Then
           If rs!UID = "ADMIN" Then
              cmdChange.Visible = True
              cmdActive.Visible = True
            
'        Else If rs!Name = "BORHAN" Then
              cmdChange.Visible = True
           Else
               cmdChange.Visible = False
               cmdActive.Visible = False
               
           End If
End Sub


Private Sub txtPersonNo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtServiceCharge_Change()
Dim i As Integer
Call Calculation
'End If
End Sub

Private Sub txtVAT_Change()
Dim i As Integer

Call Calculation
'End If
End Sub

