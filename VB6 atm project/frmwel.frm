VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmwel 
   Caption         =   "Welome To V.G.S.  ATM"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmwel.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "To Exit the Program"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "To cansel the input"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "To login"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   5160
      Width           =   4815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      ToolTipText     =   "Enter your Account ID"
      Top             =   3000
      Width           =   1695
   End
   Begin TAni.TMaxAni TMaxAni2 
      Height          =   1095
      Left            =   8760
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
   End
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ac/No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME To  V.G.S. ATM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   4590
   End
End
Attribute VB_Name = "frmwel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sd As Boolean
Dim db As Database


Private Sub Command1_Click()
'On Error GoTo abc
Data1.DatabaseName = App.Path & "\atm.mdb"
Data1.RecordSource = "select * from user_identi where user_id='" & Text1.Text & "'"
Data1.Refresh
sd = True
   
        Dim s As String
    s = Text1.Text
     
If Data1.Recordset.EOF Then
MsgBox "Entered ID was not found", vbCritical, "Warning"
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Exit Sub
Else
    frmlogin.Label1.Caption = Data1.Recordset("user_name")
    frmlogin.Show
    Me.Hide
    SendKeys "{Home}+{End}"

End If
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
i = MsgBox(" Do you want to EXIT ", vbYesNo, " V.G.S. ATM ")
If i = vbYes Then
MsgBox " Thankyou for using our program "
End
Else
If i = vbNo Then
Text1.SetFocus
End If
End If
End Sub

Private Sub Form_GotFocus()
MDIForm1.mi.Enabled = True
MDIForm1.mu.Enabled = True
MDIForm1.mtst.Enabled = True
MDIForm1.mts.Enabled = True
MDIForm1.mww.Enabled = True
MDIForm1.mds.Enabled = True
MDIForm1.mng.Enabled = True
End Sub

Private Sub Form_load()
Set db = OpenDatabase(App.Path & "\atm.mdb")
Text1.Text = ""

TMaxAni1.FileName = App.Path & "\atm.gif"
TMaxAni1.ShowGif
TMaxAni2.FileName = App.Path & "\atm.gif"
TMaxAni2.ShowGif


End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Left(Text1.Text, 1)) & Mid(Text1.Text, 2)
ad = Text1.Text
End Sub



