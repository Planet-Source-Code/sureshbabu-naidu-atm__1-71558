VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Login"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   4080
      Picture         =   "frmlogin.frx":08CA
      ScaleHeight     =   1035
      ScaleWidth      =   3435
      TabIndex        =   13
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Change PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   3600
      TabIndex        =   12
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -2640
      Width           =   1065
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hello"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   4200
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   6600
      TabIndex        =   10
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm PIN"
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
      Height          =   360
      Left            =   3840
      TabIndex        =   9
      Top             =   3360
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New PIN"
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
      Height          =   360
      Left            =   3840
      TabIndex        =   8
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter PIN"
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
      Height          =   360
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   1350
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim db As Database
    Dim res As Recordset
Dim rd, unf As Recordset
Dim da, con As Connection
Dim a As Integer

Private Sub Command1_Click()
   If Text1.Text = "" Then
   Text1.SetFocus
   MsgBox "Please Enter a PIN", vbInformation, "Warning"
   Exit Sub
   ElseIf Text2.Text = "" Then
   Text2.SetFocus
   MsgBox "Please Enter a PIN", vbInformation, "Warning"
   Exit Sub
   ElseIf Text3.Text = "" Then
   Text3.SetFocus
   MsgBox "Please Enter a PIN", vbInformation, "Warning"
   Exit Sub
   End If
    Dim s As String
    s = frmwel.Text1.Text
        Data1.DatabaseName = App.Path & "\atm.mdb"
        Data1.RecordSource = "select * from user_identi where user_ID='" & s & "'"
        Data1.Refresh
        If Data1.Recordset.NoMatch Then
        MsgBox "Doesn't Exit ", vbInformation, "Warning"
        Exit Sub
        End If
If Data1.Recordset(4) <> Text1.Text Then
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
ElseIf Data1.Recordset(4) = Text1.Text Then
    Text1.Text = ""
    frmnav.Show
MDIForm1.mi.Enabled = True
MDIForm1.mu.Enabled = True
MDIForm1.mtst.Enabled = True
MDIForm1.mts.Enabled = True
MDIForm1.mww.Enabled = True
MDIForm1.mds.Enabled = True
MDIForm1.mng.Enabled = True
      Unload Me
End If
End Sub

Private Sub Command2_Click()
Text1.SetFocus
Command3.Visible = True
Label3.Visible = True
Label4.Visible = True
Text2.Visible = True
Text3.Visible = True
Text1.Text = ""
Text2.Text = ""
Text2.PasswordChar = "*"
Text3.Text = ""
Text3.PasswordChar = "*"
End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
Beep
MsgBox ("First Enter your PIN")
Text1.SetFocus
Exit Sub
End If
    If Text2.Text <> Text3.Text Then
    MsgBox ("Your both PINs doesn't match" & vbCrLf & "Please Correct it")
    Text3.SetFocus
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Exit Sub
    Else
        Data1.DatabaseName = App.Path & "\atm.mdb"
        Data1.RecordSource = "select * from user_identi where user_ID='" & frmwel.Text1.Text & "'"
        Data1.Refresh
        Data1.Recordset.Edit
        Data1.Recordset("user_passwd") = Text3.Text
        Data1.Recordset.Update
    MsgBox ("Your PIN has been Changed")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text2.Visible = False
Text3.Visible = False
Command3.Visible = False
Label3.Visible = False
Label4.Visible = False
End If
frmnav.Show
Me.Hide
End Sub

Private Sub Command4_Click()
frmwel.Show
Unload Me
frmwel.Text1.Text = ""
frmwel.Text1.SetFocus
End Sub

Private Sub Form_load()
Data1.DatabaseName = App.Path & "\atm.mdb"
Data1.RecordSource = "select * from user_identi"
Data1.Refresh
Text1.Text = ""
Text1.PasswordChar = "*"
Command3.Visible = False
Label3.Visible = False
Label4.Visible = False
Text2.Visible = False
Text3.Visible = False
End Sub


Private Sub Form_Paint()
Text1.SetFocus

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
Private Sub Text1_LostFocus()
Text1.Text = LCase(Text1.Text)
End Sub
Private Sub Text2_LostFocus()
Text2.Text = LCase(Text2.Text)
End Sub
Private Sub Text3_LostFocus()
Text3.Text = LCase(Text3.Text)
End Sub



