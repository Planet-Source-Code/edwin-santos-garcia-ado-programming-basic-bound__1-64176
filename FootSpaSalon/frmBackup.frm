VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Facility"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4680
      Top             =   240
   End
   Begin VB.Frame Frame3 
      Height          =   1020
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   4935
      Begin MSComctlLib.ProgressBar pb 
         Height          =   220
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Max             =   50
      End
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
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Cl&ear All"
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
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Backup &Now"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Location where to Save and New Filename:"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   4935
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdBrowse2 
         Caption         =   "B&rowse"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "C&lear"
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
         Left            =   3360
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Database Name"
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
         Left            =   240
         TabIndex        =   12
         Top             =   400
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select the Database to Backup:"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4935
      Begin VB.CommandButton cmdClear1 
         Caption         =   "&Clear"
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
         Left            =   3360
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdBrowse1 
         Caption         =   "&Browse"
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
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Database Name"
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
         Left            =   240
         TabIndex        =   10
         Top             =   400
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4680
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   120
      Picture         =   "frmBackup.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    ans = MsgBox("Exit Backup Facility?", vbCritical + vbYesNo, "Confirm")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdBackup_Click()
    If Text1 = "" Or Text2 = "" Then
        MsgBox "Filenames and Locations are required!"
        Text1.SetFocus
    Else
        FileCopy Text1, Text2
        pb.Visible = True
        Timer1.Enabled = True
    End If
End Sub

Private Sub cmdBrowse1_Click()
    cd.ShowOpen
    Text1 = cd.FileName
End Sub

Private Sub cmdBrowse2_Click()
    cd.ShowSave
    Text2 = cd.FileName
End Sub

Private Sub cmdClear1_Click()
    Text1 = ""
End Sub

Private Sub cmdClear2_Click()
    Text2 = ""
End Sub

Private Sub cmdClearAll_Click()
    cmdClear1_Click
    cmdClear2_Click
End Sub

Private Sub Timer1_Timer()
    pb.Value = pb.Value + 1
    If pb.Value = 50 Then
        MsgBox "Backup Database Completed!"
        Timer1.Enabled = False
        pb.Visible = False
    End If
End Sub
