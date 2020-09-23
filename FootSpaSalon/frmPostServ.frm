VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPostServ 
   Caption         =   "Post Customer Services"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
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
   Begin MSAdodcLib.Adodc adoCust 
      Height          =   375
      Left            =   3360
      Top             =   8160
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
   Begin VB.Frame Frame1 
      Height          =   880
      Left            =   240
      TabIndex        =   44
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   45
         Top             =   320
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   5400
         MaxLength       =   5
         TabIndex        =   1
         Top             =   320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Tran Date"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Number"
         Height          =   375
         Left            =   3000
         TabIndex        =   46
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1935
      Left            =   8520
      TabIndex        =   34
      Top             =   8400
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         DataField       =   "CustNo"
         DataSource      =   "adoTemp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         DataField       =   "ServCode"
         DataSource      =   "adoTemp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         DataField       =   "ServDesc"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoTemp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   36
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         DataField       =   "Price"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoTemp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Cust No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Service Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear All Fields"
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
      Left            =   10320
      TabIndex        =   14
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
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
      Left            =   8880
      TabIndex        =   16
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "P&rint Receipt"
      Enabled         =   0   'False
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
      Left            =   7440
      TabIndex        =   15
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit Post Service"
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
      Left            =   10320
      TabIndex        =   17
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process Transaction"
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
      Left            =   8880
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Service"
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
      Left            =   7440
      TabIndex        =   12
      ToolTipText     =   "Select Service to Remove"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Transaction:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4095
      Left            =   7440
      TabIndex        =   28
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtTotPrice 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmPostServ.frx":0000
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   5318
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "CustNo"
            Caption         =   "CustNo"
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
            DataField       =   "ServCode"
            Caption         =   "ServCode"
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
            DataField       =   "ServDesc"
            Caption         =   "Service Description"
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
            DataField       =   "Price"
            Caption         =   "Price"
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
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3600
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   2775.118
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Total Service Price"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   3660
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select SERVICE:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   240
      TabIndex        =   22
      Top             =   4200
      Width           =   6735
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         DataField       =   "ServDesc"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoServ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAddServ 
         Caption         =   "&Add Service"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   27
         ToolTipText     =   "Enter the Customer number first."
         Top             =   2160
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmPostServ.frx":0016
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "ServCode"
            Caption         =   "ServCode"
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
            DataField       =   "ServCat"
            Caption         =   "ServCat"
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
            DataField       =   "ServDesc"
            Caption         =   "Description"
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
            DataField       =   "Price"
            Caption         =   "Price"
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
            DataField       =   "Rem"
            Caption         =   "Rem"
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
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   2415.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   2775.118
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         DataField       =   "Price"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoServ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         DataField       =   "ServCat"
         DataSource      =   "adoServ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         DataField       =   "ServCode"
         DataSource      =   "adoServ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Service Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Customer Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   6735
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Sex"
         Height          =   375
         Left            =   4680
         TabIndex        =   21
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Tel. No."
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1005
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   520
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc adoServ 
      Height          =   375
      Left            =   120
      Top             =   8160
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
      Index           =   0
      Left            =   1560
      TabIndex        =   23
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoTemp 
      Height          =   375
      Left            =   120
      Top             =   7680
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
      RecordSource    =   "Temp"
      Caption         =   "Temp Table"
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
   Begin VB.TextBox dummy 
      DataField       =   "CustNo"
      DataSource      =   "adoTemp"
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
      Index           =   2
      Left            =   1560
      TabIndex        =   31
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoTranDet 
      Height          =   375
      Left            =   3360
      Top             =   7680
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
      RecordSource    =   "TranDetails"
      Caption         =   "Tran Details"
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
   Begin VB.TextBox dummy 
      DataField       =   "TranNo"
      DataSource      =   "adoTranDet"
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
      Index           =   3
      Left            =   4800
      TabIndex        =   43
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoTran 
      Height          =   375
      Left            =   6480
      Top             =   8160
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
      RecordSource    =   "TranRecords"
      Caption         =   "Tran Details"
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
   Begin VB.TextBox dummy 
      DataField       =   "TranNo"
      DataSource      =   "adoTran"
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
      Index           =   4
      Left            =   7920
      TabIndex        =   48
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
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
      Index           =   1
      Left            =   4680
      TabIndex        =   29
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   7200
      X2              =   7200
      Y1              =   0
      Y2              =   8640
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   240
      Picture         =   "frmPostServ.frx":002C
      Top             =   120
      Width           =   6690
   End
End
Attribute VB_Name = "frmPostServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varCno, varTotPrice, varTranNo

Private Sub cmdAddServ_Click()
    If txtDesc <> "" Then
        adoTemp.Recordset.AddNew
        adoTemp.Recordset!TranNo = varTranNo
        adoTemp.Recordset!CustNo = Text2
        adoTemp.Recordset!custname = Text3
        adoTemp.Recordset!ServCode = txtCode
        adoTemp.Recordset!ServDesc = txtDesc
        adoTemp.Recordset!Price = Val(txtPrice)
        adoTemp.Recordset.Update
        varTotPrice = varTotPrice + Val(txtPrice)
        txtTotPrice = Format(varTotPrice, "P #,###,##0.00")
    End If
End Sub

Private Sub cmdClear_Click()
    varTotPrice = 0
    varCno = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    txtTotPrice = ""
    If adoTemp.Recordset.BOF = True And adoTemp.Recordset.EOF = True Then Exit Sub
    adoTemp.Recordset.MoveFirst
    Do While Not adoTemp.Recordset.EOF
        adoTemp.Recordset.Delete
        adoTemp.Recordset.MoveNext
    Loop
    txtTotPrice = Format(varTotPrice, "P #,###,##0.00")
    cmdProcess.Enabled = True
    Text2.SetFocus
End Sub

Private Sub cmdExit_Click()
    ans = MsgBox("Exit Posting of Services?", vbCritical + vbYesNo, "Confirm")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdNew_Click()
    cmdClear_Click
End Sub

Private Sub cmdPrint_Click()
    adoTemp.Recordset.Requery
    Load drReceipt
    drReceipt.Show 1
End Sub

Private Sub cmdProcess_Click()
    adoTran.Recordset.AddNew
    adoTran.Recordset!TranDate = Date
    adoTran.Recordset!TranTime = Time
    adoTran.Recordset!CustNo = Text2
    adoTran.Recordset!TotPrice = varTotPrice
    adoTran.Recordset.Update
    
    adoTran.Recordset.MoveLast
    varTranNo = adoTran.Recordset!TranNo
    
    adoTemp.Recordset.MoveFirst
    Do While Not adoTemp.Recordset.EOF
        adoTranDet.Recordset.AddNew
        adoTranDet.Recordset!TranNo = varTranNo
        adoTranDet.Recordset!TranDate = Date
        adoTranDet.Recordset!CustNo = adoTemp.Recordset!CustNo
        adoTranDet.Recordset!ServCode = adoTemp.Recordset!ServCode
        adoTranDet.Recordset!Price = adoTemp.Recordset!Price
        adoTranDet.Recordset.Update
        
        adoTemp.Recordset!TranNo = varTranNo
        adoTemp.Recordset.Update
        adoTemp.Recordset.MoveNext
    Loop
    MsgBox "All transactions were processed and recorded!"
    cmdPrint.Enabled = True
    cmdProcess.Enabled = False
End Sub

Private Sub cmdRemove_Click()
    varTotPrice = varTotPrice - Val(Text7)
    txtTotPrice = Format(varTotPrice, "P #,###,##0.00")
    
    If adoTemp.Recordset.BOF = True And adoTemp.Recordset.EOF = True Then Exit Sub
    MsgBox Text8 & " - Service was deleted!"
    
    On Error Resume Next
    adoTemp.Recordset.Delete
    adoTemp.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
    Text1.Text = Date
    If adoTemp.Recordset.BOF = True And adoTemp.Recordset.EOF = True Then Exit Sub
    adoTemp.Recordset.MoveFirst
    Do While Not adoTemp.Recordset.EOF
        adoTemp.Recordset.Delete
        adoTemp.Recordset.MoveNext
    Loop
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    varFound = False
    If KeyCode = 13 Then
        varCno = Text2
        
        adoCust.Recordset.MoveFirst
        Do While Not adoCust.Recordset.EOF
            If adoCust.Recordset!CustNo = varCno Then
                varFound = True
                Exit Do
            End If
            adoCust.Recordset.MoveNext
        Loop
        If varFound = True Then
            Text3 = adoCust.Recordset!custname
            Text4 = adoCust.Recordset!custadd
            Text5 = adoCust.Recordset!custtelno
            Text6 = adoCust.Recordset!custsex
            cmdAddServ.Enabled = True
        Else
            MsgBox "Customer Number does not exist!"
            Text2 = ""
        End If
    End If

End Sub

