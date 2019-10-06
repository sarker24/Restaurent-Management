VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Restaurant Billing System [LOTUS ETANG]"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   Picture         =   "frmMain.frx":058A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10350
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Current User :"
            TextSave        =   "Current User :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   13626
            Text            =   "Software Developed by ""MAS IT SOLUTIONS"". Hot Line : 02-9031260, 01915682291"
            TextSave        =   "Software Developed by ""MAS IT SOLUTIONS"". Hot Line : 02-9031260, 01915682291"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu First 
      Caption         =   "............."
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "S&etup"
      Begin VB.Menu mnuUser 
         Caption         =   "&User Information"
      End
      Begin VB.Menu mnuRGuest 
         Caption         =   "Restaurant &Guest"
      End
      Begin VB.Menu mnuMenuGroupSetup 
         Caption         =   "Menu &Group Setup"
      End
      Begin VB.Menu mnuRestaurentCatagory 
         Caption         =   "Menu &Category Setup"
      End
      Begin VB.Menu mnuMenuItemSetup 
         Caption         =   "Menu &Item Setup"
      End
      Begin VB.Menu mnuRRoom 
         Caption         =   "&Table Setup"
      End
      Begin VB.Menu mnuWaiterNameSetup 
         Caption         =   "&Waiter Name Setup"
      End
      Begin VB.Menu mnuCashMemoModify 
         Caption         =   "Cash &Memo Modify"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRIC 
         Caption         =   "Restaurent Items & Consumption"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuquit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuRRM 
         Caption         =   "Restaurent Raw Materials"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuCashMemo 
      Caption         =   "Cash &Memo"
   End
   Begin VB.Menu mnuReservation 
      Caption         =   "Reservation"
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuDailySales 
         Caption         =   "&Daily Sales Statement"
      End
      Begin VB.Menu mnuSalesSummery 
         Caption         =   "Sales Su&mmery"
      End
      Begin VB.Menu mnuItemWiseSales 
         Caption         =   "Item &Wise Sales"
      End
      Begin VB.Menu mnuSBPMode 
         Caption         =   "Sales By Payment Mode"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRStatement 
         Caption         =   "Reservation Statement"
      End
      Begin VB.Menu mnuSDStatement 
         Caption         =   "Sales Due Statement"
      End
      Begin VB.Menu mnuWSStatement 
         Caption         =   "Waiter wise Sales Statement"
      End
      Begin VB.Menu mnuWTStatement 
         Caption         =   "Waiter wise Tips Statement"
      End
      Begin VB.Menu mnuWTSummary 
         Caption         =   "Waiter Wise Tips Summary"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "Accounts"
      Begin VB.Menu mnuCAHead 
         Caption         =   "Chart of Accounts Head"
      End
      Begin VB.Menu mnuVEntry 
         Caption         =   "Voucher Entry"
      End
      Begin VB.Menu mnuFSStatement 
         Caption         =   "Floor Sheet Statement"
      End
      Begin VB.Menu mnuCBook 
         Caption         =   "Cash Book"
      End
      Begin VB.Menu mnuGLedger 
         Caption         =   "General Ledger"
      End
      Begin VB.Menu mnuPLAccounts 
         Caption         =   "PL Accounts"
      End
   End
   Begin VB.Menu mnuBackUp 
      Caption         =   "Back&Up"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalculator 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuCommunication 
         Caption         =   "Commun&ication"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpPage 
         Caption         =   "Help Page"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Log Off"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Me.StatusBar1.Panels(2) = frmLogin.txtUID
Me.StatusBar1.Panels(3) = Date
Me.StatusBar1.Panels(4) = Time
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
ques = MsgBox("Do you want to exit the Application", vbQuestion + vbYesNo, "Restaurant Management System.....")
If ques = vbYes Then
    End
Else
    Cancel = 1
End If
End Sub

Private Sub mnuAbout_Click()
FrmAbout.Show vbModal
End Sub

Private Sub mnuBackUp_Click()
frmBackUp.Show vbModal
End Sub

Private Sub mnuCAHead_Click()
frmAccountsHead.Show vbModal
End Sub

Private Sub mnuCalculator_Click()
On Error Resume Next
   Shell "calc.exe"
End Sub

Private Sub mnuCashMemo_Click()
frmCashMemo.Show vbModal
End Sub

Private Sub mnuCashMemoModify_Click()
frmCMEdit.Show vbModal
End Sub

Private Sub mnuCBook_Click()
RptCashBook.Show vbModal
End Sub

Private Sub mnuCommunication_Click()
Call Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE", vbMaximizedFocus)
End Sub

Private Sub mnuDailySales_Click()
RptDSales.Show vbModal
End Sub

Private Sub mnuExit_Click()
Dim res As VbMsgBoxResult
    res = MsgBox("Are you sure you want to log off?", vbYesNo + vbQuestion)
    If res = vbYes Then
    Unload Me
'    frmLogin.txtUID.SetFocus
    frmLogin.Show
    frmLogin.txtUID = ""
    frmLogin.txtPassword = ""
    frmLogin.txtUID.SetFocus
    Else
    End If
End Sub

Private Sub mnuFSStatement_Click()
RptFloorSheet.Show vbModal
End Sub

Private Sub mnuGLedger_Click()
RptGLedger.Show vbModal
End Sub

Private Sub mnuItemWiseSales_Click()
RptItemwiseSales.Show vbModal
End Sub

Private Sub mnuMenuGroupSetup_Click()
frmMenuGroup.Show vbModal
End Sub

Private Sub mnuPmode_Click()
frmPaymentMode.Show vbModal
End Sub

Private Sub mnuMenuItemSetup_Click()
frmItemEntry.Show vbModal
End Sub

Private Sub mnuquit_Click()
End
End Sub

Private Sub mnuReservation_Click()
frmReservation.Show vbModal
End Sub

Private Sub mnuRestaurentCatagory_Click()
frmItemGroup.Show vbModal
End Sub

Private Sub mnuRestaurentItemType_Click()
frmItemEntry.Show vbModal
End Sub

Private Sub mnuRestaurentMenuGroup_Click()
frmItemGroup.Show vbModal
End Sub

Private Sub mnuRGuest_Click()
frmCustomer.Show vbModal
End Sub

Private Sub mnuRRoom_Click()
frmRRoom.Show vbModal
End Sub

Private Sub mnuSales_Click()
RptSalesSummery.Show vbModal
End Sub

Private Sub mnuRStatement_Click()
RptReservation.Show vbModal
End Sub

Private Sub mnuSalesSummery_Click()
RptSalesSummery.Show vbModal
End Sub

Private Sub mnuSalesWaiter_Click()
Rptwaiterwise.Show vbModal
End Sub

Private Sub mnuSBPMode_Click()
RptSBPMode.Show vbModal
End Sub

Private Sub mnuSDStatement_Click()
RptSalesDue.Show vbModal
End Sub

Private Sub mnuUser_Click()
frmUser.Show vbModal
End Sub

Private Sub mnuWNEntry_Click()
frmItemEntry.Show vbModal
End Sub

Private Sub mnuVEntry_Click()
frmMoneyReceipt.Show vbModal
End Sub

Private Sub mnuWaiterNameSetup_Click()
frmRestaurantWaiters.Show vbModal
End Sub

Private Sub mnuWSStatement_Click()
RptWaiterSales.Show vbModal
End Sub

Private Sub mnuWTStatement_Click()
RptwaiterTips.Show vbModal
End Sub

Private Sub mnuWTSummary_Click()
RptWTStatement.Show vbModal
End Sub
