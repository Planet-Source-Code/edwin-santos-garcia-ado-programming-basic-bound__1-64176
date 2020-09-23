VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "SPA and SALON - Customer Tracking System"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":224C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2568
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   1535
      ButtonWidth     =   1402
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Customer"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Services"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Post"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Receipt"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Backup"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Restore"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log Off"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCustomer 
         Caption         =   "&Customer Maintenance"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuServices 
         Caption         =   "&Services Maintenance"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit System"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuPostService 
         Caption         =   "&Post Services"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintReceipt 
         Caption         =   "Print &Receipt"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSearchCustomer 
         Caption         =   "S&earch Customer"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "T&ools"
      Begin VB.Menu mnuSystemUsers 
         Caption         =   "System &Users"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore &Database"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Rep&orts"
      Begin VB.Menu mnuListCustomer 
         Caption         =   "List of Customers"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuListServices 
         Caption         =   "List of Services"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuListEmployee 
         Caption         =   "List of Employees"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalesIncomeReport 
         Caption         =   "Sales/Income Report"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRptQuery 
         Caption         =   "Report Query"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Co&ntents"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show 1
End Sub

Private Sub mnuBackup_Click()
    Load frmBackup
    frmBackup.Show 1
End Sub

Private Sub mnuContents_Click()
    Load frmContents
    frmContents.Show 1
End Sub

Private Sub mnuCustomer_Click()
    Load frmCustomer
    frmCustomer.Show
End Sub

Private Sub mnuExit_Click()
    ans = MsgBox("Exit Customer Tracking System?", vbQuestion + vbYesNo, "Confirm")
    If ans = vbYes Then
        Load frmShutdown
        Unload Me
        frmShutdown.Show 1
    End If
End Sub

Private Sub mnuListCustomer_Click()
    Load drCustomers
    drCustomers.Show 1
    Unload drCustomers
End Sub

Private Sub mnuListEmployee_Click()
    Load drEmployee
    drEmployee.Show 1
    Unload drEmployee
End Sub

Private Sub mnuListServices_Click()
    Load drServices
    drServices.Show 1
    Unload drServices
End Sub

Private Sub mnuLogOff_Click()
    ans = MsgBox(varName & " - You will be Logged-off?", vbQuestion + vbYesNo, "Confirm")
    If ans = vbYes Then
        frmMain.Hide
        Load frmLogin
        frmLogin.Show
        Unload Me
    End If
End Sub

Private Sub mnuPostService_Click()
    Load frmPostServ
    frmPostServ.Show
End Sub

Private Sub mnuPrintReceipt_Click()
    MsgBox "Browse all the receipt and print selected page!"
    Load drTranDetails
    drTranDetails.Show 1
    Unload drTranDetails
End Sub

Private Sub mnuRestore_Click()
    Load frmRestore
    frmRestore.Show 1
End Sub

Private Sub mnuRptQuery_Click()
    Load frmRptQry
    frmRptQry.Show 1
End Sub

Private Sub mnuSalesIncomeReport_Click()
    Load frmReports
    frmReports.Show 1
End Sub

Private Sub mnuSearchCustomer_Click()
    Load frmCustomer
    frmCustomer.Show
End Sub

Private Sub mnuServices_Click()
    Load frmServices
    frmServices.Show
End Sub

Private Sub mnuSystemUsers_Click()
    Load frmSystemUser
    frmSystemUser.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnuCustomer_Click
        Case 2: mnuServices_Click

        Case 4: mnuPostService_Click
        Case 5: mnuPrintReceipt_Click
        Case 6: mnuSearchCustomer_Click

        Case 8: mnuSystemUsers_Click
        Case 9: mnuBackup_Click
        Case 10: mnuRestore_Click

        Case 12: mnuSalesIncomeReport_Click
        Case 13: mnuContents_Click

        Case 15: mnuLogOff_Click
        Case 16: mnuExit_Click
    End Select
End Sub
