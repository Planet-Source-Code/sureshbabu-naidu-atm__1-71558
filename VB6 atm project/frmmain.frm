VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmmain 
   BackColor       =   &H80000012&
   Caption         =   "V.G.S.  ATM"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Skip"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   4680
      Top             =   -2760
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5280
      Top             =   -2760
   End
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "uresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   8760
      TabIndex        =   9
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1935
      Left            =   7440
      TabIndex        =   8
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "iripal "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   5400
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Software..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1200
      Left            =   3240
      TabIndex        =   5
      Top             =   5760
      Width           =   9360
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004-2005 V.G.S."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4080
      TabIndex        =   4
      Top             =   7680
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A T M"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   150.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3750
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   7845
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   9240
      X2              =   11400
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   3360
      X2              =   9480
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   2520
      X2              =   3360
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   2520
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "aibhav "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1080
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ld As Integer
Dim s As Integer


Private Sub Command1_Click()
Unload Me
frmwel.Show
End Sub

Private Sub Form_load()
TMaxAni1.FileName = App.Path & "\atm.gif"
TMaxAni1.ShowGif
End Sub

Private Sub Timer1_Timer()
Dim r As Byte
Dim g As Byte
Dim b As Byte
r = Rnd() * 255
g = Rnd() * 255
b = Rnd() * 255
Line1.BorderColor = RGB(r, g, b)
Line2.BorderColor = RGB(r, g, b)
Line3.BorderColor = RGB(r, g, b)
Line4.BorderColor = RGB(r, g, b)
Label1.ForeColor = RGB(b, r, g)
Label3.ForeColor = RGB(r, g, b)
Label5.ForeColor = RGB(g, b, r)
Label6.ForeColor = RGB(b, r, g)
Label8.ForeColor = RGB(b, r, g)
End Sub

Private Sub Timer2_Timer()
Dim r As Byte
Dim g As Byte
Dim b As Byte
r = Rnd() * 255
g = Rnd() * 255
b = Rnd() * 255
Label2.ForeColor = RGB(g, b, r)
Label4.ForeColor = RGB(g, b, r)
Label7.ForeColor = RGB(g, b, r)
'*********
If ld >= 2 Then
frmwel.Show
Unload Me
Exit Sub
End If
ld = ld + 1
End Sub


