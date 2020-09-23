VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmwithdraw 
   Caption         =   "Withdraw"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmwithdraw.frx":0000
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<<  &Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      MaskColor       =   &H8000000B&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "To go back to Navigation Menu"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      MaskColor       =   &H8000000B&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "to cancel withdrawing"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "To complete your withdrawing process"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -1560
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount To Be Withdrawn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      TabIndex        =   5
      Top             =   2880
      Width           =   3060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.G.S. Bank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   5040
      TabIndex        =   4
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "frmwithdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim s, d As Double
If Text1.Text = "" Then
Text1.SetFocus
MsgBox "Please Enter a value", vbInformation, "Warning"
Exit Sub
End If
Data1.Recordset.MoveFirst
Data1.Recordset.FindFirst "user_id='" & frmwel.Text1.Text & "'"
s = Val(Data1.Recordset("balance"))
If s <= 5000 Then
MsgBox ("You can't Withdrw" & vbCrLf & "'Cause Your Min Balance is < 5000")
Else
d = s - Val(Text1.Text)
If d >= 5000 Then
Data1.Recordset.Edit
Data1.Recordset("balance") = Str(d)
Data1.Recordset.Update
MsgBox ("Your withdraw is Completed")
Text1.Text = ""
Else
MsgBox ("Sorry User" & vbCrLf & "'Cause after withdraw will be <5000")
End If
End If
End Sub
Private Sub Command2_Click()
Text1.Text = ""
End Sub


Private Sub Command3_Click()
frmnav.Show
Unload Me
End Sub

Private Sub Form_load()
Data1.DatabaseName = App.Path & "\atm.mdb"
Data1.RecordSource = "select * from user_account"
Data1.Refresh
 Data1.Recordset.FindFirst "user_id='" & frmwel.Text1.Text & "'"
TMaxAni1.FileName = App.Path & "\fallingmoney.gif"
TMaxAni1.ShowGif
End Sub


