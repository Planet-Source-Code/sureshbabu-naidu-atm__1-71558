VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "V.G.S. ATM"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2700
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/9/04"
            Object.ToolTipText     =   "System Date"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "7:31 PM"
            Object.ToolTipText     =   "System Time"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   10663
            Text            =   "                                                Welcome To   V.G.S.   ATM                                        "
            TextSave        =   "                                                Welcome To   V.G.S.   ATM                                        "
            Object.ToolTipText     =   " Welcome To   V.G.S.   ATM   "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock"
         EndProperty
      EndProperty
   End
   Begin VB.Menu m 
      Caption         =   "&Menu"
      Begin VB.Menu mi 
         Caption         =   "ID Form"
         Shortcut        =   ^I
      End
      Begin VB.Menu mu 
         Caption         =   "Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu mtst 
         Caption         =   "User's Status"
         Shortcut        =   ^S
      End
      Begin VB.Menu mts 
         Caption         =   "User's Detail"
         Shortcut        =   ^D
      End
      Begin VB.Menu mww 
         Caption         =   "Withdraw"
         Shortcut        =   ^W
      End
      Begin VB.Menu mds 
         Caption         =   "Deposit"
         Shortcut        =   ^O
      End
      Begin VB.Menu mng 
         Caption         =   "Navigation"
         Shortcut        =   ^N
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu admin 
         Caption         =   "Administrator"
         Shortcut        =   ^A
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ma 
      Caption         =   "&About"
   End
   Begin VB.Menu me 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub admin_Click()
frmadminlogin.Show
Unload frmlogin
Unload frmuserinfo
Unload frmdep
Unload frmdetail
Unload frmwithdraw
Unload frmnav
Unload frmwel
End Sub

Private Sub ext_Click()
End
End Sub

Private Sub ma_Click()
frmabout.Show
End Sub

Private Sub MDIForm_Load()
'admin.Enabled = False
mi.Enabled = False
mu.Enabled = False
mtst.Enabled = False
mts.Enabled = False
mww.Enabled = False
mds.Enabled = False
mng.Enabled = False
End Sub

Private Sub mds_Click()
frmdep.Show
Unload frmlogin
Unload frmuserinfo
Unload frmwel
Unload frmdetail
Unload frmwithdraw
Unload frmnav
Unload frmadminlogin
End Sub

Private Sub me_Click()
i = MsgBox(" Do you to Exit ", vbYesNo, " V.G.S ATM ")
If i = vbYes Then
Unload Me
Else
If i = vbNo Then
frmwel.Show
End If
End If
End Sub

Private Sub mi_Click()
MDIForm1.mi.Enabled = False
MDIForm1.mu.Enabled = False
MDIForm1.mtst.Enabled = False
MDIForm1.mts.Enabled = False
MDIForm1.mww.Enabled = False
MDIForm1.mds.Enabled = False
MDIForm1.mng.Enabled = False

frmwel.Show
Unload frmlogin
Unload frmuserinfo
Unload frmdep
Unload frmdetail
Unload frmwithdraw
Unload frmnav
Unload frmadminlogin
End Sub

Private Sub mng_Click()
frmnav.Show
Unload frmlogin
Unload frmuserinfo
Unload frmdep
Unload frmdetail
Unload frmwithdraw
Unload frmwel
Unload frmadminlogin
End Sub

Private Sub mts_Click()
frmdetail.Show
Unload frmlogin
Unload frmuserinfo
Unload frmdep
Unload frmwel
Unload frmwithdraw
Unload frmnav
Unload frmadminlogin
End Sub

Private Sub mtst_Click()
frmuserinfo.Show
Unload frmlogin
Unload frmwel
Unload frmdep
Unload frmdetail
Unload frmwithdraw
Unload frmnav
Unload frmadminlogin
End Sub

Private Sub mu_Click()
frmlogin.Show
Unload frmwel
Unload frmuserinfo
Unload frmdep
Unload frmdetail
Unload frmwithdraw
Unload frmnav
Unload frmadminlogin
End Sub

Private Sub mww_Click()
frmwithdraw.Show
Unload frmlogin
Unload frmuserinfo
Unload frmdep
Unload frmdetail
Unload frmwel
Unload frmnav
Unload frmadminlogin
End Sub

Private Sub mnunpad_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("c:\windows\notepad.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Notepad Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub

Private Sub mnucal_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("c:\windows\calc.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub

