VERSION 5.00
Begin VB.Form frmtransaction 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   4800
      TabIndex        =   11
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   4920
      TabIndex        =   10
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   4440
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Date and time "
      Height          =   855
      Left            =   720
      TabIndex        =   8
      Top             =   6840
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "Amount deposited"
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Amount withdrawn"
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Balance"
      Height          =   855
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Account no"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "name"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmtransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset

Private Sub Command1_Click()
MsgBox "Transaction proceeding"
frmnav.Show
End Sub

Private Sub Form_load()
Load frmwel
Load frmdep
Load frmwithdraw
Set db = OpenDatabase(App.Path & "\Atm.mdb")

Set rs = db.OpenRecordset("user_identi")
Set rs1 = db.OpenRecordset("user_account")

If rs.RecordCount = 0 Then
MsgBox "Record not found", vbCritical, "V.G.S. ATM"
frmwel.Show
frmtransaction.Hide
Else
rs.MoveFirst
rs1.MoveFirst
End If
Do While Not rs.EOF

If (frmwel.Text1.Text = rs.Fields(0) And frmwel.Text1.Text = rs1.Fields(0)) Then
Text1.Text = rs.Fields(1)
Text2.Text = rs.Fields(0)
Text3.Text = rs1.Fields(1)
Text4.Text = frmwithdraw.Text1.Text
Text5.Text = frmdep.Text2.Text
Text6.Text = Date & Time
Else
rs.MoveNext
rs1.MoveNext
 End If
 Loop
End Sub



