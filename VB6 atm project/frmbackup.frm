VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form frmbackup 
   Caption         =   "Administrator"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmbackup.frx":0000
   LinkTopic       =   "Form13"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&BackUp"
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Frame frameCurrBackUp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choose Path for BackUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   2520
      TabIndex        =   11
      Top             =   4800
      Width           =   6375
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   3015
      End
      Begin VB.FileListBox File1 
         Height          =   675
         Left            =   0
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame frameBackup 
      Caption         =   "Last BackUp Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   6375
      Begin VB.Label lblLastPath 
         BackStyle       =   0  'Transparent
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   615
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Time 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblLastDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblLastTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note : Microsoft Scripting Runtime Library is Referenced
'       For Making Object of File System Object


Dim Fsys As New FileSystemObject
Dim bckupFile As File
'Backup Cmd Btn
Private Sub Command1_Click()
On Error Resume Next
    cmdSave.Enabled = False
    Label1.Caption = "Please Wait, Backup in Progress..."
    Label1.BackColor = vbGreen
    Label1.ForeColor = vbYellow
    Dim destination As String
    Dim source As String
    Dim currDate, currTime As String
    currDate = Format$(Now, "mm - dd - yy")
    currTime = Format$(Now, "hh:mm:ss AM/PM")
    destination = File1.Path & "\" & "Atm.mdb"
    source = App.Path & "\Atm.mdb"
    'MsgBox "Source : " & source
    'MsgBox "Destination : " & destination
    Set bckupFile = Fsys.GetFile(finalpath)
    bckupFile.Attributes = Compressed
    Fsys.CopyFile source, destination, True
    'Saving Current Backup Details
    SaveSetting App.Title, "Settings", "BackupPath", destination
    SaveSetting App.Title, "Settings", "BackupDate", currDate
    SaveSetting App.Title, "Settings", "BackupTime", currTime
    Label1.Caption = " V.G.S. ATM "
    Label1.BackColor = vbYellow
    Label1.ForeColor = vbBlue
    cmdSave.Enabled = True
    MsgBox "BackUp Process Over", vbInformation, "Backup"
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

'Reading Previously Backup Details
Private Sub Form_load()
    Dim lastPath As String
    Dim lastDate As String
    Dim lastTime As String
    
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.Title, "Settings", "BackupPath")
    lastDate = GetSetting(App.Title, "Settings", "BackupDate")
    lastTime = GetSetting(App.Title, "Settings", "BackupTime")
    
    If lastPath = "" Then
        lblLastPath.Caption = "No Backup made previously"
        lblLastDate.Caption = " "
        lblLastTime.Caption = " "
    Else
        lblLastPath.Caption = lastPath
        lblLastDate.Caption = lastDate & "  (mm-dd-yy)"
        lblLastTime.Caption = lastTime
    End If
TMaxAni1.FileName = App.Path & "\save.gif"
TMaxAni1.ShowGif
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub






