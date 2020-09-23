VERSION 5.00
Begin VB.Form frmreport 
   Caption         =   "Report"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form12"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnushow 
         Caption         =   "Show Report"
         Begin VB.Menu dreport 
            Caption         =   "Detailed Report"
            Shortcut        =   {F1}
         End
         Begin VB.Menu areport 
            Caption         =   "Account Report"
            Shortcut        =   {F2}
         End
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print Report"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub areport_Click()
DataReport2.Show
End Sub

Private Sub dreport_Click()
DataReport1.Show
End Sub

Private Sub mnuexit_Click()
i = MsgBox("Do you want to exit from User Report ", vbYesNo, "User Report")
If i = vbYes Then
Unload Me
Unload DataReport1
Else
If i = vbNo Then
frmreport.Show
End If
End If
End Sub

Private Sub mnuPrint_Click()
    DataReport1.PrintReport
    DataReport2.PrintReport
End Sub

