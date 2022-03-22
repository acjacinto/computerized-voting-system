VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmloading 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   6435
   ClientTop       =   3270
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   Picture         =   "frmloading.frx":0000
   ScaleHeight     =   5100
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   7320
      Top             =   480
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
      Max             =   105
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label lblstatus1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   3360
      Width           =   2535
   End
End
Attribute VB_Name = "frmloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
  If ProgressBar1.Value = 10 Then
     lblstatus.Caption = "Loading."
  ElseIf ProgressBar1.Value = 20 Then
     lblstatus.Caption = "Loading.."
  ElseIf ProgressBar1.Value = 30 Then
     lblstatus.Caption = "Loading..."
  ElseIf ProgressBar1.Value = 40 Then
     lblstatus.Caption = "Initializing."
  ElseIf ProgressBar1.Value = 50 Then
     lblstatus.Caption = "Initializing.."
  ElseIf ProgressBar1.Value = 60 Then
     lblstatus.Caption = "Initializing..."
  ElseIf ProgressBar1.Value = 70 Then
     lblstatus.Caption = "Please Wait."
  ElseIf ProgressBar1.Value = 80 Then
     lblstatus.Caption = "Please Wait.."
  ElseIf ProgressBar1.Value = 90 Then
     lblstatus.Caption = "Please Wait..."
  ElseIf ProgressBar1.Value = 100 Then
     lblstatus.Caption = "Loading Successful"
  ElseIf ProgressBar1.Value = 105 Then
     lblstatus.Caption = ""
     Timer1.Enabled = False
     frmlogin.Show
  End If
lblstatus1.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
Unload Me

End If
End Sub
