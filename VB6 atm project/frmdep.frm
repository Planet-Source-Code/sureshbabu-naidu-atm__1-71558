VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmdep 
   Caption         =   "Deposit"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmdep.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   1200
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Go back to Navigation Menu"
      Top             =   4080
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
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
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   4200
      TabIndex        =   8
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -1200
      Width           =   2055
   End
   Begin VB.TextBox Text2 
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
      Left            =   6720
      TabIndex        =   0
      Top             =   2880
      Width           =   2055
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
      Left            =   2160
      TabIndex        =   4
      Top             =   -3440
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount To Be Deposited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   7
      Top             =   2880
      Width           =   3030
   End
   Begin VB.Label Label2 
      Caption         =   "Check No"
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
      Left            =   240
      TabIndex        =   6
      Top             =   -3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.G.S. Bank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1110
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   5325
   End
End
Attribute VB_Name = "frmdep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt As Connection
Dim rs As Recordset
Private Sub Command1_Click()
If Text2.Text = "" Then
Text2.SetFocus
MsgBox "Please Enter a Correct Amount", vbInformation, "Warning"
Exit Sub
ElseIf Text2.Text <= 1 Then
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Text2.SetFocus
MsgBox "Please Enter a Correct Amount", vbInformation, "Warning"
Exit Sub
End If

 Data1.DatabaseName = App.Path & "\atm.mdb"
 Data1.RecordSource = "select * from user_account"
 Data1.Refresh
  Data1.Recordset.FindFirst "user_id='" & frmwel.Text1.Text & "'"
    If Text2.Text = "" Then
    Text2.SetFocus
    MsgBox ("Please Enter your Amount")
    Exit Sub
    End If
Data1.Recordset.Edit
Data1.Recordset("balance") = Str(Val(Data1.Recordset("balance")) + Val(Text2.Text))
Data1.Recordset.Update
MsgBox ("Your Balance has been Updated")
Text1.Text = ""
Text2.Text = ""
'frmnav.Show
'Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Command3_Click()
frmnav.Show
Unload Me
End Sub

Private Sub Form_load()
    Text1.Text = ""
    Text2.Text = ""
    TMaxAni1.FileName = App.Path & "\wallet.gif"
    TMaxAni1.ShowGif
    End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 9 _
Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
Debug.Print KeyAscii
End Sub



