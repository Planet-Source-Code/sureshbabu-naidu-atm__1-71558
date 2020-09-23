VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmabout 
   Caption         =   "About"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Close Me"
      Default         =   -1  'True
      Height          =   345
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   4800
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   7770
      Top             =   6135
   End
   Begin TAni.TMaxAni TMaxAni2 
      Height          =   615
      Left            =   8880
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
   End
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Authors : Vaibhav Cheturvedi"
      Height          =   195
      Left            =   4560
      TabIndex        =   8
      ToolTipText     =   ".   Abdul Rafay Mansoor   ."
      Top             =   2640
      Width           =   2070
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Authors : Giripal Rawat"
      Height          =   195
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   ".   Abdul Rafay Mansoor   ."
      Top             =   2400
      Width           =   1620
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   975
      Left            =   2640
      Top             =   5040
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmabout.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Suresh Giripal Vaibhav"
      Top             =   5040
      Width           =   5865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V G S =  Value Growth Success"
      Height          =   195
      Left            =   4425
      TabIndex        =   5
      ToolTipText     =   ".   windows_me@rediffmail.com   ."
      Top             =   2880
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Authors : H.SureshBabu.Naidu "
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   ".   Abdul Rafay Mansoor   ."
      Top             =   2160
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   2820
      X2              =   8490
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmabout.frx":0998
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   " Just Kidding !!!!!!!  u r free to distibute this program but give us some mention n Credit"
      Top             =   3240
      Width           =   5655
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_load()
TMaxAni1.FileName = App.Path & "\you.gif"
TMaxAni1.ShowGif
TMaxAni2.FileName = App.Path & "\atm.gif"
TMaxAni2.ShowGif
Shape1.BorderColor = vbBlack
End Sub

Private Sub Timer1_Timer()
If Shape1.BorderColor = vbBlack Then
Shape1.BorderColor = vbRed
Else
Shape1.BorderColor = vbBlack
If Shape1.BorderColor = vbRed Then
Shape1.BorderColor = vbBlue
Else
Shape1.BorderColor = vbBlack
End If
End If
End Sub

