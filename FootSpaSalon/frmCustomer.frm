VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomer 
   Caption         =   "Customer File Maintenance"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "&Customer Information:"
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   7095
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Left            =   2280
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2280
         TabIndex        =   14
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   2280
         TabIndex        =   13
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2280
         TabIndex        =   12
         Top             =   3600
         Width           =   2295
      End
      Begin VB.ComboBox cboSex 
         Appearance      =   0  'Flat
         Height          =   405
         ItemData        =   "frmCustomer.frx":0000
         Left            =   2280
         List            =   "frmCustomer.frx":000A
         TabIndex        =   11
         Text            =   "Male"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Customer No"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   855
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Name"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1455
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Home Address"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2055
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Sex (M/F)"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3015
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Tel No. / Cell No."
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3615
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adoCust 
      Height          =   375
      Left            =   4080
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\FootSpaSalon\FSS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\FootSpaSalon\FSS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Customer"
      Caption         =   "Customer Table"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frmAsk 
      Caption         =   "Save?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8400
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdNo 
         Caption         =   "&No"
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "&Yes"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox dummy 
      DataField       =   "CustNo"
      DataSource      =   "adoCust"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   7560
      X2              =   11880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   8400
      Picture         =   "frmCustomer.frx":001C
      Top             =   480
      Width           =   2460
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   7560
      X2              =   7560
      Y1              =   0
      Y2              =   8640
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   240
      Picture         =   "frmCustomer.frx":0BC9
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7065
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varCno, sw

Sub clrtxt()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    cboSex = "Male"
End Sub

Sub txt(X As Boolean)
    Text2.Enabled = X
    Text3.Enabled = X
    Text4.Enabled = X
    cboSex.Enabled = X
End Sub

Sub cmd(X As Boolean)
    cmdAdd.Enabled = X
    cmdEdit.Enabled = X
    cmdDelete.Enabled = X
    cmdSearch.Enabled = X
    cmdBrowse.Enabled = X
    cmdExit.Enabled = X
End Sub

Private Sub cmdAdd_Click()
    sw = 1
    clrtxt
    txt (True)
    frmAsk.Visible = True
    cmd (False)
    
    If adoCust.Recordset.RecordCount <> 0 Then
        adoCust.Recordset.MoveLast
        varCno = adoCust.Recordset!CustNo
        varCno = varCno + 1
    Else
        varCno = 1
    End If
    Text1.Text = Format(varCno, "00000")
    Text2.SetFocus
End Sub

Private Sub cmdBrowse_Click()
    Load frmBrowseCust
    frmBrowseCust.Show 1
End Sub

Private Sub cmdDelete_Click()
    sw = 3
    varFound = False
    varCno = InputBox("Enter the Customer Number to Delete", "Delete Customer")
    
    If varCno = Empty Then
        MsgBox "Customer number is required!"
        Exit Sub
    End If
    
    adoCust.Recordset.MoveFirst
    Do While Not adoCust.Recordset.EOF
        If adoCust.Recordset!CustNo = varCno Then
            varFound = True
            Exit Do
        End If
        adoCust.Recordset.MoveNext
    Loop
    
    If varFound = True Then
        Text1 = adoCust.Recordset!CustNo
        Text2 = adoCust.Recordset!custname
        Text3 = adoCust.Recordset!custadd
        cboSex = adoCust.Recordset!custsex
        Text4 = adoCust.Recordset!custtelno
        ans = MsgBox("Delete this record?", vbQuestion + vbYesNo, "Confirm")
        If ans = vbYes Then
            adoCust.Recordset.Delete
            adoCust.Recordset.MoveFirst
            MsgBox "Record successfully deleted!"
        Else
            MsgBox "Record is not deleted!"
        End If
    Else
        MsgBox "Customer Number does not exist!"
    End If
End Sub

Private Sub cmdEdit_Click()
    sw = 2
    varFound = False
    varCno = 0
    varCno = InputBox("Enter the Customer Number to Edit", "Edit Customer")
    
    If varCno = Empty Then
        MsgBox "Customer number is required!"
        Exit Sub
    End If
    
    adoCust.Recordset.MoveFirst
    Do While Not adoCust.Recordset.EOF
        If adoCust.Recordset!CustNo = varCno Then
            varFound = True
            Exit Do
        End If
        adoCust.Recordset.MoveNext
    Loop
    
    If varFound = True Then
        txt (True)
        cmd (False)
        frmAsk.Visible = True
        
        Text1 = adoCust.Recordset!CustNo
        Text2 = adoCust.Recordset!custname
        Text3 = adoCust.Recordset!custadd
        cboSex = adoCust.Recordset!custsex
        Text4 = adoCust.Recordset!custtelno
        MsgBox "Make the necessary changes!"
    Else
        MsgBox "Customer Number does not exist!"
    End If
End Sub

Private Sub cmdExit_Click()
    ans = MsgBox("Exit Customer File Maintenance?", vbCritical + vbYesNo, "Confirm")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdNo_Click()
    MsgBox "Record not saved!"
    txt (False)
    cmd (True)
    frmAsk.Visible = False
End Sub

Private Sub cmdSearch_Click()
    sw = 4
    varFound = False
    varCno = InputBox("Enter the Customer Number to Search", "Search Customer")
    
    If varCno = Empty Then
        MsgBox "Customer number is required!"
        Exit Sub
    End If
    
    adoCust.Recordset.MoveFirst
    Do While Not adoCust.Recordset.EOF
        If adoCust.Recordset!CustNo = varCno Then
            varFound = True
            Exit Do
        End If
        adoCust.Recordset.MoveNext
    Loop
    
    If varFound = True Then
        Text1 = adoCust.Recordset!CustNo
        Text2 = adoCust.Recordset!custname
        Text3 = adoCust.Recordset!custadd
        cboSex = adoCust.Recordset!custsex
        Text4 = adoCust.Recordset!custtelno
    Else
        MsgBox "Customer Number does not exist!"
    End If
End Sub

Private Sub cmdYes_Click()
    If sw = 1 Then
        adoCust.Recordset.AddNew
    End If
        adoCust.Recordset!CustNo = Text1
        adoCust.Recordset!custname = Text2
        adoCust.Recordset!custadd = Text3
        adoCust.Recordset!custsex = Left(cboSex, 1)
        adoCust.Recordset!custtelno = Text4
        adoCust.Recordset.Update
        MsgBox "Record successfully saved!"
    
    cmd (True)
    txt (False)
    frmAsk.Visible = False
End Sub

Private Sub Form_Load()
    txt (False)
End Sub
