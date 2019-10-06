VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmItemGroup 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Restaurent Menu Group"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frmRestaurentMGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Find"
      Height          =   795
      Left            =   5280
      Picture         =   "frmRestaurentMGroup.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   945
   End
   Begin VB.CommandButton chameleonButton1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      Height          =   795
      Left            =   4425
      Picture         =   "frmRestaurentMGroup.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   825
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "C&lose"
      Height          =   795
      Left            =   3600
      Picture         =   "frmRestaurentMGroup.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   825
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1920
      Picture         =   "frmRestaurentMGroup.frx":1FE8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   825
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   1080
      Picture         =   "frmRestaurentMGroup.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   795
      Left            =   2760
      Picture         =   "frmRestaurentMGroup.frx":317C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   825
   End
   Begin VB.TextBox txtMenuCatagory 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Text            =   " "
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtSerialNo 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   " "
      Top             =   240
      Width           =   3975
   End
   Begin MSForms.ComboBox cmbItemCatagory 
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   840
      Width           =   3975
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7011;873"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Menu Catagory"
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
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblSerialNo 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Serial No"
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
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblMenuGroup 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Menu Group"
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
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmItemGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsItemGroup             As ADODB.Recordset
Private strFileName           As String
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim str As String
'--------------------------------------------------------------
Private oReportApp                        As CRPEAuto.Application
Private oReport                           As CRPEAuto.Report
Private oReportDatabase                   As CRPEAuto.Database
Private oReportDatabaseTables             As CRPEAuto.DatabaseTables
Private oReportDatabaseTable              As CRPEAuto.DatabaseTable
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions

Private Sub chameleonButton1_Click()
'    Call printReport
End Sub

Private Sub cmdClose_Click()

    Unload Me
End Sub
'
Private Sub cmdCancel_Click()

   cmdCancel.Enabled = False
   cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    chameleonButton1.Enabled = True
    txtSerialNo.Enabled = False
    Call allClear
    Call alldisable
    If Not rsItemGroup.EOF Then FindRecord
End Sub
'
Private Sub cmdnew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
        chameleonButton1.Enabled = False
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),0) as SerialNo from ItemCatagory"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo.text = Val(rs!SerialNo) + 1
            
        Call allenable
        cmbItemCatagory.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSerialNo.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                chameleonButton1.Enabled = True
                Call alldisable
                s = cmbItemCatagory
                rsItemGroup.Requery
                rsItemGroup.MoveFirst
                rsItemGroup.Find "MenuGroup='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
'
    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
        chameleonButton1.Enabled = False
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdOpen.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                chameleonButton1.Enabled = True
                Call alldisable
                rsItemGroup.Requery

                Dim s As String
                s = cmbItemCatagory
                rsItemGroup.Find "MenuGroup='" & parseQuotes(s) & "'"
'                Call search
'                Call countrysearch
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
   
    frmItemGroupSearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub




Private Sub Find_Click()
    frmItemGroup.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub Form_Load()

    Call Connect
    Call ItemCatagory
       ModFunction.StartUpPosition Me
    Set rsItemGroup = New ADODB.Recordset
    rsItemGroup.Open "select * from ItemCatagory", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsItemGroup.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
   
    If Not rsItemGroup.EOF Then FindRecord
    
    txtSerialNo.Enabled = False
    
End Sub

Private Sub ItemCatagory()
    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT MGName FROM MenuGroup ORDER BY MGName ASC"), cn, adOpenStatic
    While Not rsTemp2.EOF
        cmbItemCatagory.AddItem rsTemp2("MGName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
    
End Sub


Private Sub allenable()
    txtSerialNo.Enabled = True
    cmbItemCatagory.Enabled = True
    txtMenuCatagory.Enabled = True
End Sub

Private Sub alldisable()
    txtSerialNo.Enabled = False
    cmbItemCatagory.Enabled = False
    txtMenuCatagory.Enabled = False
    
End Sub


Private Sub allClear()
'    ModFunction.TextClear Me
    cmbItemCatagory.text = ""
    txtMenuCatagory.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO ItemCatagory(SerialNo,MenuGroup,MenuCatagory) " & _
                   " VALUES ('" & parseQuotes(txtSerialNo) & "','" & parseQuotes(cmbItemCatagory) & "', " & _
                   " '" & parseQuotes(txtMenuCatagory) & "')"
                   
                   
         rcupdate = True
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    Else

        cn.Execute "Update ItemCatagory Set MenuGroup='" & parseQuotes(cmbItemCatagory) & _
                  "',MenuCatagory='" & parseQuotes(txtMenuCatagory) & "' where SerialNo='" & parseQuotes(txtSerialNo) & "'"
                  
                 
     rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans

    Exit Function



ErrHandler:
    cn.RollbackTrans
    rsItemGroup.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Iteam Group Name"
            cmbItemCatagory = ""
            cmbItemCatagory.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsItemGroup.EOF Then
        txtSerialNo = rsItemGroup("SerialNo")
        cmbItemCatagory = rsItemGroup("MenuGroup")
        txtMenuCatagory = rsItemGroup("MenuCatagory")
End If
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (cmbItemCatagory.text = "") Then
       MsgBox "Enter Menu Group Name"
       cmbItemCatagory.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtMenuCatagory.text = "") Then
      MsgBox "Enter Menu Catagory"
      txtMenuCatagory.SetFocus
      IsValidRecord = False
      Exit Function
    End If
    
'If CmdEdit.Caption <> "&Update" Or CmdEdit.Caption = "&Update" Then
'If cmdNew.Caption <> "&New" Or cmdNew.Caption <> "&New" Then
'        If rsItemGroup.RecordCount > 0 Then
'        If rsItemGroup.State <> 0 Then rsItemGroup.Close
'            rsItemGroup.Open "select * from ItemCatagory where upper(MenuGroup)='" & Strings.UCase(Strings.Trim(parseQuotes(cmbItemCatagory))) & "'", cn
'
'             If Not rsItemGroup.EOF Then
'        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
'          cmbItemCatagory.SetFocus
'          IsValidRecord = False
'         Exit Function
'            End If
'
'         End If
'    End If
End Function
'.............................................................................

Public Sub PrintReport()
'On Error GoTo ErrorHan
Dim strPath            As String
Dim rsFactProf         As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\MenuGroup.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select ItemCatagory.SerialNo,ItemCatagory.MenuGroup,ItemCatagory.MenuCatagory"
             
    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True
oReport.DiscardSavedData
oReport.Preview "Item Group Infromation of '" & cmbItemCatagory.text & "'", , , , , 16777216 Or 524288 Or 65536


End Sub


Public Sub PopulateIteam(StrID As String)


    rsItemGroup.MoveFirst
    rsItemGroup.Find "SerialNo=" & parseQuotes(StrID)
    If rsItemGroup.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub










