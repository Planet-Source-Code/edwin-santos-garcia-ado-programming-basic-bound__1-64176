VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSystemUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System User Maintenance"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSystemUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   6135
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back to Main Menu"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Record"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Record"
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
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Record"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox dummy 
      DataField       =   "empno"
      DataSource      =   "adoUsers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adoUsers 
      Height          =   375
      Left            =   120
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
      ForeColor       =   0
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\FootSpaSalon\FSS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\FootSpaSalon\FSS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employee"
      Caption         =   "Users File"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information:"
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8655
      Begin VB.Frame frmPass 
         Caption         =   "Password Settings:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2535
         Left            =   4800
         TabIndex        =   12
         Top             =   240
         Width           =   3735
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel Changes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   6
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save Changes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   5
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtVerPass 
            Appearance      =   0  'Flat
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   2040
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   1320
            Width           =   1600
         End
         Begin VB.TextBox txtPass 
            Appearance      =   0  'Flat
            DataField       =   "password"
            DataSource      =   "adoUsers"
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   2040
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   840
            Width           =   1600
         End
         Begin VB.TextBox txtUser 
            Appearance      =   0  'Flat
            DataField       =   "username"
            DataSource      =   "adoUsers"
            Height          =   390
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   2
            Top             =   360
            Width           =   1600
         End
         Begin VB.Label Label3 
            Caption         =   "Verify Password"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   1365
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Password"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   885
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "User Name"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   405
            Width           =   2535
         End
      End
      Begin MSDataGridLib.DataGrid dgEmp 
         Bindings        =   "frmSystemUser.frx":06EA
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "username"
            Caption         =   "username"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "password"
            Caption         =   "password"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "empno"
            Caption         =   "Emp No"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "empname"
            Caption         =   "Employee Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "add"
            Caption         =   "Address"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "pos"
            Caption         =   "Position"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "telno"
            Caption         =   "Tel No."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   120
      Picture         =   "frmSystemUser.frx":0701
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmSystemUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    dgEmp.AllowAddNew = True
    dgEmp.AllowUpdate = True
    MsgBox "Enter the Employee Informations!"
    frmPass.Enabled = True
    dgEmp.SetFocus
End Sub

Private Sub cmdBack_Click()
    ans = MsgBox("Exit System Users Maintenance?", vbCritical + vbYesNo, "Confirm")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    adoUsers.Recordset.CancelUpdate
    dgEmp.AllowAddNew = False
    dgEmp.AllowUpdate = False
    frmPass.Enabled = False
    MsgBox "Record changes not saved!"
End Sub

Private Sub cmdDelete_Click()
    ans = MsgBox("Delete this record? Are you sure?", vbCritical + vbYesNo, "Warning!")
    If ans = vbYes Then
        adoUsers.Recordset.Delete
        adoUsers.Recordset.MoveFirst
        MsgBox "Record succesfully deleted!"
    Else
        MsgBox "Record is not deleted!"
    End If
End Sub

Private Sub cmdEdit_Click()
    MsgBox "Make the necessary changes!"
    frmPass.Enabled = True
    dgEmp.AllowAddNew = False
    dgEmp.AllowUpdate = True
    dgEmp.SetFocus
End Sub

Private Sub cmdSave_Click()
    If txtPass <> txtVerPass Then
        MsgBox "Password is not the same!"
        txtVerPass.SetFocus
    Else
        adoUsers.Recordset.Update
        dgEmp.AllowAddNew = False
        dgEmp.AllowUpdate = False
        frmPass.Enabled = False
        MsgBox "Record successfully saved"
    End If
End Sub
