VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmadminlogin 
   Caption         =   "Administrator"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmadminlogin.frx":0000
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
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
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -940
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -100
      Visible         =   0   'False
      Width           =   1140
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "+"
      TabIndex        =   0
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "frmadminlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Command1_Click()
Data1.DatabaseName = App.Path & "/atm.mdb"
Data1.RecordSource = "select * from admin"
Data1.Refresh
If Text1.Text <> Data1.Recordset("pass") Then
    a = a + 1
If a = 1 Then
    b = MsgBox("Please Enter a Corect PIN", , "Login......")
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
ElseIf a = 2 Then
     b = MsgBox("Please Enter a Corect PIN", , "Login")
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
ElseIf a = 3 Then
a = MsgBox("Hey user" & vbCrLf & "Don't be Over Smart", , "Login")
End
End If
ElseIf Data1.Recordset("pass") = Text1.Text Then
    Text1.Text = ""
    frmadmin1.Show
    Unload Me
End If
End Sub

Private Sub Command2_Click()
frmwel.Show
Unload Me
End Sub

Private Sub Form_load()
TMaxAni1.FileName = App.Path & "\moneyworld.gif"
TMaxAni1.ShowGif
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub


