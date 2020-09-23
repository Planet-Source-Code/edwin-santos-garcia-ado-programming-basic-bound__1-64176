VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login Security System"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   650
      Width           =   3855
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "User Name"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      TabIndex        =   4
      Top             =   2020
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   600
      TabIndex        =   3
      Top             =   2020
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Security System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":030A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    ans = MsgBox("Exit the System?", vbQuestion + vbYesNo, "Confirm!")
    If ans = vbYes Then
        End
    End If
End Sub

Private Sub cmdLogin_Click()
    varFound = False
    varName = Trim(Text1)
    varPass = Trim(Text2)
    
    If varName = "" Or varPass = "" Then
        MsgBox "Username and Password are required!", vbCritical, "Warning!"
        Text1.SetFocus
    Else
        adoRS.MoveFirst
        Do While Not adoRS.EOF
            If adoRS!UserName = varName And adoRS!Password = varPass Then
                varFound = True
                Exit Do
            End If
            adoRS.MoveNext
        Loop
        
        If varFound = False Then
            MsgBox "Username and Password does not exist!", vbExclamation, "Warning!"
            Text1 = ""
            Text2 = ""
            Text1.SetFocus
        Else
            MsgBox "Welcome to the System - " & adoRS!empname
            Load frmMain
            Unload Me
            frmMain.Show
        End If
    End If
End Sub

Private Sub Form_Load()
    OpenDatabase
    OpenEmployee
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseDatabase
End Sub
