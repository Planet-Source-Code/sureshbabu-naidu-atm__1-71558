VERSION 5.00
Begin VB.Form frmaddnew 
   Caption         =   "Add New"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmaddnew.frx":0000
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmaddnew.frx":08CA
      Left            =   6120
      List            =   "frmaddnew.frx":08D7
      TabIndex        =   9
      Text            =   "Select "
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< &Back"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&New Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   3240
      TabIndex        =   25
      Top             =   6000
      Width           =   6255
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -2280
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -1440
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   24
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User  ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User  Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "User  Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User  Ph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "E - Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Ac Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.G.S.  Administrator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Left            =   3360
      TabIndex        =   13
      Top             =   -120
      Width           =   7215
   End
End
Attribute VB_Name = "frmaddnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim da As Connection
Dim rs, ss As Recordset

Private Sub Command1_Click()
    If Text2.Text = "" Then
    Text2.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Text3.Text = "" Then
    Text3.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Text4.Text = "" Then
    Text4.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Text5.Text = "" Then
    Text5.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Text6.Text = "" Then
    Text6.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Text7.Text = "" Then
    Text7.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Combo1.Text = "" Then
    Text8.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Text9.Text = "" Then
    Text9.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    ElseIf Text10.Text = "" Then
    Text10.SetFocus
    MsgBox "Value Required", vbInformation, "Warning"
    Exit Sub
    End If
If Text9.Text <> Text10.Text Then
Text10.SetFocus
Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)
MsgBox ("Your Confirm PIN doesn't match")
Exit Sub
End If
    Data1.DatabaseName = App.Path & "\atm.mdb"
    Data2.DatabaseName = App.Path & "\atm.mdb"
    Data1.RecordSource = "select * from user_identi"
    Data1.Refresh
    Data2.RecordSource = "select * from user_account"
    Data2.Refresh
Data1.Recordset.AddNew
Data1.Recordset!user_id = Text1.Text
Data1.Recordset!user_name = Text2.Text
Data1.Recordset!user_passwd = Text10.Text
Data1.Recordset!user_add = Text3.Text
Data1.Recordset!user_ph = Val(Text4.Text)
Data1.Recordset!e_mail = Text5.Text
Data1.Recordset.Update
    Data2.Recordset.AddNew
    Data2.Recordset!user_id = Text1.Text
    Data2.Recordset!balance = Text6.Text
    Data2.Recordset!min_balance = Text7.Text
    Data2.Recordset!ac_type = Combo1.Text
    Data2.Recordset.Update
    MsgBox ("New Entery is Completed")
    Command1.Enabled = Fale
End Sub

Private Sub Command2_Click()
Call Form_load
Command1.Enabled = True
Text2.SetFocus
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub

Private Sub Command3_Click()
frmwel.Show
Unload Me
End Sub

Private Sub Command4_Click()
frmadmin1.Show
Unload Me
End Sub

Private Sub Form_load()
Dim uid As String * 20
Data1.DatabaseName = App.Path & "\atm.mdb"
Data1.RecordSource = "select user_id from user_identi"
Data1.Refresh
Data1.Recordset.MoveLast
    uid = Data1.Recordset("user_id")
    str_uid = Left(uid, 1)
    num_uid = Mid(uid, 2)
    new_id = CInt(num_uid) + 1
    If new_id < 9 Then
    Text1.Text = str_uid & "00" & new_id
    ElseIf (new_id < 99 And new_id >= 10) Then
    Text1.Text = str_uid & "0" & new_id
    ElseIf (new_id < 999 And new_id >= 100) Then
    Text1.Text = str_uid & new_id
    End If
    
    End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 _
Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 45 Or KeyAscii = 9 _
Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 45 Or KeyAscii = 9 _
Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 45 Or KeyAscii = 9 _
Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Then
Else
KeyAscii = 0
End If
End Sub



