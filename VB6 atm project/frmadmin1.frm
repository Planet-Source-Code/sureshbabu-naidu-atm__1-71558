VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmadmin1 
   Caption         =   "Administrator"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmadmin1.frx":0000
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1111
      ButtonWidth     =   1905
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New "
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back Up"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Status Report"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit "
         EndProperty
      EndProperty
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11040
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10200
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmadmin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_load()
TMaxAni1.FileName = App.Path & "\atm.gif"
TMaxAni1.ShowGif
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
        Case 1
                frmaddnew.Show
                frmaddnew.Text1.SetFocus
        
        Case 2
                frmbackup.Show
                                
        Case 3
                frmreport.Show
                frmreport.SetFocus
        Case 4
                i = MsgBox(" Do you to Exit ", vbYesNo, " V.G.S ATM Administrator")
                If i = vbYes Then
                Unload frmadmin1
                frmwel.Show
                Else
                If i = vbNo Then
                frmreport.SetFocus
                End If
                End If
        End Select
End Sub

