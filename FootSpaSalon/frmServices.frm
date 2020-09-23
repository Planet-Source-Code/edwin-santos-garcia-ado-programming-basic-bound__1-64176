VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmServices 
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
   ScaleHeight     =   5505
   ScaleWidth      =   6630
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "&Service Information:"
      ForeColor       =   &H00C00000&
      Height          =   4455
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   7095
      Begin VB.ComboBox cboCat 
         Height          =   405
         ItemData        =   "frmServices.frx":0000
         Left            =   2280
         List            =   "frmServices.frx":000D
         TabIndex        =   1
         Text            =   "SPA"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2280
         TabIndex        =   4
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   0
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   2280
         TabIndex        =   2
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2280
         TabIndex        =   3
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Service Code"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   855
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Service Category"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1455
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Descriptions"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2055
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Price / Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3015
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Remark(s)"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3650
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adoServ 
      Height          =   375
      Left            =   3960
      Top             =   6360
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
      RecordSource    =   "Services"
      Caption         =   "Service Table"
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
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdNo 
         Caption         =   "&No"
         Height          =   495
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "&Yes"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Service"
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
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse Services"
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
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search Service"
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
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Service"
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
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Service"
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
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New Service"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox dummy 
      DataField       =   "ServCode"
      DataSource      =   "adoServ"
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
      TabIndex        =   14
      Top             =   6360
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
      Picture         =   "frmServices.frx":0026
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
      Picture         =   "frmServices.frx":0BD3
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7065
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varServNo, sw

Sub clrtxt()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    cboCat = "SPA"
End Sub

Sub txt(X As Boolean)
    Text1.Enabled = X
    Text2.Enabled = X
    Text3.Enabled = X
    Text4.Enabled = X
    cboCat.Enabled = X
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
    Text1.SetFocus
End Sub

Private Sub cmdBrowse_Click()
    Load frmBrowseServ
    frmBrowseServ.Show 1
End Sub

Private Sub cmdDelete_Click()
    sw = 3
    varFound = False
    varServNo = InputBox("Enter the Service Code to Delete", "Delete Service")
    
    If varServNo = Empty Then
        MsgBox "Service number is required!"
        Exit Sub
    End If
    
    adoServ.Recordset.MoveFirst
    Do While Not adoServ.Recordset.EOF
        If adoServ.Recordset!ServCode = varServNo Then
            varFound = True
            Exit Do
        End If
        adoServ.Recordset.MoveNext
    Loop
    
    If varFound = True Then
        Text1 = adoServ.Recordset!ServCode
        cboCat = adoServ.Recordset!ServCat
        Text2 = adoServ.Recordset!ServDesc
        Text3 = adoServ.Recordset!Price
        Text4 = adoServ.Recordset!rem
        ans = MsgBox("Delete this record?", vbQuestion + vbYesNo, "Confirm")
        If ans = vbYes Then
            adoServ.Recordset.Delete
            adoServ.Recordset.MoveFirst
            MsgBox "Record successfully deleted!"
        Else
            MsgBox "Record is not deleted!"
        End If
    Else
        MsgBox "Service Code does not exist!"
    End If
End Sub

Private Sub cmdEdit_Click()
    sw = 2
    varFound = False
    varServNo = InputBox("Enter the Service Code to Edit", "Edit Service")
    
    If varServNo = Empty Then
        MsgBox "Service number is required!"
        Exit Sub
    End If
    
    adoServ.Recordset.MoveFirst
    Do While Not adoServ.Recordset.EOF
        If adoServ.Recordset!ServCode = varServNo Then
            varFound = True
            Exit Do
        End If
        adoServ.Recordset.MoveNext
    Loop
    
    If varFound = True Then
        txt (True)
        cmd (False)
        frmAsk.Visible = True
        
        Text1 = adoServ.Recordset!ServCode
        cboCat = adoServ.Recordset!ServCat
        Text2 = adoServ.Recordset!ServDesc
        Text3 = adoServ.Recordset!Price
        Text4 = adoServ.Recordset!rem
        MsgBox "Make the necessary changes!"
    Else
        MsgBox "Service Code does not exist!"
    End If
End Sub

Private Sub cmdExit_Click()
    ans = MsgBox("Exit Service File Maintenance?", vbCritical + vbYesNo, "Confirm")
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
    varServNo = InputBox("Enter the Service Code to Search", "Search Service")
    
    If varServNo = Empty Then
        MsgBox "Service number is required!"
        Exit Sub
    End If
    
    adoServ.Recordset.MoveFirst
    Do While Not adoServ.Recordset.EOF
        If adoServ.Recordset!ServCode = varServNo Then
            varFound = True
            Exit Do
        End If
        adoServ.Recordset.MoveNext
    Loop
    
    If varFound = True Then
        Text1 = adoServ.Recordset!ServCode
        cboCat = adoServ.Recordset!ServCat
        Text2 = adoServ.Recordset!ServDesc
        Text3 = adoServ.Recordset!Price
        Text4 = adoServ.Recordset!rem
    Else
        MsgBox "Service Code does not exist!"
    End If
End Sub

Private Sub cmdYes_Click()

    If sw = 1 Then
        adoServ.Recordset.AddNew
    End If
        adoServ.Recordset!ServCode = Text1
        adoServ.Recordset!ServCat = cboCat
        adoServ.Recordset!ServDesc = Text2
        adoServ.Recordset!Price = Val(Text3)
        If Text4 <> "" Then
            adoServ.Recordset!rem = Text4
        Else
            adoServ.Recordset!rem = "None"
        End If
        adoServ.Recordset.Update
        MsgBox "Record successfully saved!"
    
    cmd (True)
    txt (False)
    frmAsk.Visible = False
End Sub

Private Sub Form_Load()
    txt (False)
End Sub
