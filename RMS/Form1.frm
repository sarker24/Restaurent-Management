VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Height          =   1575
      Left            =   480
      TabIndex        =   6
      Top             =   5280
      Width           =   7575
      _Version        =   196616
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   13361
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "SSDBGrid1"
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton btnResult 
      Caption         =   "Result"
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton btnconnect 
      Caption         =   "Connect"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   4320
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim Rec As New ADODB.Recordset





Private Sub btnconnect_Click()



' conn.Open "Data Source=192.168.1.35,1433;Network Library=DBMSSOCN;Initial Catalog=AddressBook;User ID=sa;Password=00020297;"
 conn.Open "provider=sqloledb;Data Source=192.168.1.36,1433;Initial Catalog=AddressBook;User ID=sa;Password=00020297"
           Rec.Open "Select * from Patient where PatientID = 3 ", conn, adOpenStatic
            

''
''          conn.BeginTrans
''            Rec.Open "INSERT INTO Patient(PatientID,PatientName,Sex,Comments)" & _
''                     "SELECT PatientID,PatientName,Sex,Comments" & _
''                     " from Patient1 where PatientID = 1", conn, adOpenStatic
''
                     
'           conn.Close
           If Rec.RecordCount > O Then
           Text1.text = Rec.Fields!PatientID
           Text2.text = Rec.Fields!PatientName
           Text3.text = Rec.Fields!Sex
           Text4.text = Rec.Fields!Comments
           
'              MsgBox "Found result", vbInformation, "confirmation"
'            conn.CommitTrans
             Call insert

          
           
           Else
            MsgBox "No record exist", vbInformation, "confirmation"
            
            
            End If

conn.Close


           
End Sub

Private Function insert() As Boolean

Call Connect
 On Error GoTo ErrHandler

    cn.BeginTrans
    
cn.Execute "INSERT INTO Patient1(PatientID,PatientName,Sex,Comments)" & _
                    " VALUES (" & Text1.text & ",'" & parseQuotes(Text2.text) & "', " & _
                   " '" & parseQuotes(Text3.text) & "','" & parseQuotes(Text4.text) & "')"
                   
                   
'        cn.Execute "INSERT INTO AccountsHead(AID,AHName,Department,AHType) " & _
'                   " VALUES ('" & parseQuotes(txtAID) & "','" & parseQuotes(txtAHName) & "', " & _
'                   " '" & parseQuotes(cmbDepartment) & "','" & parseQuotes(cmbAHType) & "')"
'
             insert = True
             
         
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    
        
    

    cn.CommitTrans
    

 Exit Function


ErrHandler:
    cn.RollbackTrans
'    rsAccountsHead.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "No such record exit"
           
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select
           
End Function

