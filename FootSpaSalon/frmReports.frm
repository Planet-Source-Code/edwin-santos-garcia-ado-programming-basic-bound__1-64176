VERSION 5.00
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports - Main Menu"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back to Main"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1440
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sales/Income Reports:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4095
      Begin VB.CommandButton cmdSalesRpt 
         Caption         =   "Sales and Income"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2160
         Picture         =   "frmReports.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdTranRpt 
         Caption         =   "Detailed &Transactions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         Picture         =   "frmReports.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
      Begin VB.CommandButton cmdEmpRpt 
         Caption         =   "&Employee Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2760
         Picture         =   "frmReports.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdServRpt 
         Caption         =   "&Services Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1440
         Picture         =   "frmReports.frx":1278
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCustRpt 
         Caption         =   "&Customer Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Picture         =   "frmReports.frx":1582
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   120
      Picture         =   "frmReports.frx":188C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4080
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    ans = MsgBox("Exit Printing Reports?", vbCritical + vbYesNo, "Confirm")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdCustRpt_Click()
    Load drCustomers
    drCustomers.Show 1
    Unload drCustomers
End Sub

Private Sub cmdEmpRpt_Click()
    Load drEmployee
    drEmployee.Show 1
    Unload drEmployee
End Sub

Private Sub cmdSalesRpt_Click()
    Load drTranRecords
    drTranRecords.Show 1
    Unload drTranRecords
End Sub

Private Sub cmdServRpt_Click()
    Load drServices
    drServices.Show 1
    Unload drServices
End Sub

Private Sub cmdTranRpt_Click()
    Load drTranDetails
    drTranDetails.Show 1
    Unload drTranDetails
End Sub
