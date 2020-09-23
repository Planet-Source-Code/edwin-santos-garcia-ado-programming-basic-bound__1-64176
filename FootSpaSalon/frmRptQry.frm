VERSION 5.00
Begin VB.Form frmRptQry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptQry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Summary by:"
      ForeColor       =   &H000000C0&
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "&By Date && CustNo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   360
         TabIndex        =   13
         Top             =   2760
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Br&owse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4080
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCustNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   3720
         TabIndex        =   9
         Top             =   2760
         Width           =   1500
      End
      Begin VB.CommandButton cmdClear 
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
         Height          =   550
         Left            =   2040
         TabIndex        =   8
         Top             =   2760
         Width           =   1500
      End
      Begin VB.TextBox txtServCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Height          =   420
         Left            =   4080
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
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
         Height          =   420
         Left            =   4080
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Customer Number"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1755
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   5520
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label3 
         Caption         =   "Service Code"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1155
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Date: (mm/dd/yy)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   550
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Report Query"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmRptQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtDate = ""
    txtServCode = ""
    txtCustNo = ""
End Sub

Private Sub cmdExit_Click()
    ans = MsgBox("Exit Printing Query Reports?", vbCritical + vbYesNo, "Confirm")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    If deFSS.rscmdByDate.State = adStateOpen Then
        deFSS.rscmdByDate.Close
    End If
        
    If txtDate <> "" Then
        deFSS.cmdByDate Trim(txtDate.Text)
        Load drByDate
        drByDate.Show 1
    Else
        MsgBox "Date is required!"
        txtDate.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    If deFSS.rscmdByServCode.State = adStateOpen Then
        deFSS.rscmdByServCode.Close
    End If
        
    If txtServCode <> "" Then
        deFSS.cmdByServCode Trim(txtServCode.Text)
        Load drByServCode
        drByServCode.Show 1
    Else
        MsgBox "Service Code is required!"
        txtServCode.SetFocus
    End If
End Sub

Private Sub Command3_Click()
    If deFSS.rscmdByCustNo.State = adStateOpen Then
        deFSS.rscmdByCustNo.Close
    End If
        
    If txtCustNo <> "" Then
        deFSS.cmdByCustNo Trim(txtCustNo.Text)
        Load drByCustNo
        drByCustNo.Show 1
    Else
        MsgBox "Customer Number is required!"
        txtCustNo.SetFocus
    End If
End Sub

Private Sub Command4_Click()
    If txtDate = "" Or txtCustNo = "" Then
        MsgBox "Date / Customer No is required!"
        txtDate.SetFocus
    Else
        If deFSS.rscmdDateCust.State = adStateOpen Then
            deFSS.rscmdDateCust.Close
        End If
        deFSS.cmdDateCust Trim(txtDate.Text), Trim(txtCustNo.Text)
        Load drByDC
        drByDC.Show 1
    End If
End Sub
