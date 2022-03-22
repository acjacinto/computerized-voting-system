VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   7320
   ClientLeft      =   5820
   ClientTop       =   1920
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSettings.frx":0000
   ScaleHeight     =   7320
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adosample 
      Height          =   375
      Left            =   1920
      Top             =   5760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmSettings.frx":16A0A
      OLEDBString     =   $"frmSettings.frx":16A9A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from sampletable"
      Caption         =   "adosample"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fmeaccesscode 
      BackColor       =   &H00000080&
      Caption         =   "Input Access Code"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   4800
      TabIndex        =   6
      Top             =   4200
      Width           =   4215
      Begin VB.CommandButton cmdenable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Functions"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtaccesscode 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdback 
      Height          =   735
      Left            =   6720
      Picture         =   "frmSettings.frx":16B2A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdRall 
      Height          =   855
      Left            =   4920
      Picture         =   "frmSettings.frx":18476
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdRvoters 
      Height          =   855
      Left            =   4920
      Picture         =   "frmSettings.frx":19FF8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   3975
   End
   Begin VB.CommandButton cmdRpartylist 
      Height          =   855
      Left            =   480
      Picture         =   "frmSettings.frx":1BE37
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdRcandidates 
      Height          =   855
      Left            =   480
      Picture         =   "frmSettings.frx":1DF45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdRvotes 
      Height          =   855
      Left            =   480
      Picture         =   "frmSettings.frx":202DB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   3975
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmSettings.Hide
frmAdminmenu.Show
End Sub

Private Sub cmdenable_Click()

If txtaccesscode.Text = "080955089" Then
  cmdRvotes.Enabled = True
  cmdRcandidates.Enabled = True
  cmdRpartylist.Enabled = True
  cmdRvoters.Enabled = True
  cmdRall.Enabled = True
  txtaccesscode.Text = ""
ElseIf txtaccesscode.Text = "1098765432" Then
  cmdRvotes.Enabled = False
  cmdRcandidates.Enabled = False
  cmdRpartylist.Enabled = False
  cmdRvoters.Enabled = False
  cmdRall.Enabled = False
  txtaccesscode.Text = ""
Else
  MsgBox "Invalid Access Code", vbExclamation, "Verification"
  txtaccesscode.Text = ""
End If
End Sub

Private Sub cmdRall_Click()
MsgBox("Are you sure to reset all?", vbYesNo + vbQuestion, "Verification") = vbYes
End Sub

Private Sub cmdRcandidates_Click()
MsgBox("Are you sure to reset all candidates?", vbYesNo + vbQuestion, "Verification") = vbYes
End Sub

Private Sub cmdRpartylist_Click()
MsgBox("Are you sure to reset party list?", vbYesNo + vbQuestion, "Verification") = vbYes
End Sub

Private Sub cmdRvoters_Click()
If MsgBox("Are you sure to reset all votes?", vbYesNo + vbQuestion, "Verification") = vbYes Then
   Do Until adosample.Recordset.EOF
   adosample.Recordset.Fields("LastName") = "0"
   Loop
End If
End Sub

Private Sub cmdRvotes_Click()
MsgBox("Are you sure to reset all voters?", vbYesNo + vbQuestion, "Verification") = vbYes
End Sub

Private Sub Form_Load()
cmdRvotes.Enabled = False
cmdRcandidates.Enabled = False
cmdRpartylist.Enabled = False
cmdRvoters.Enabled = False
cmdRall.Enabled = False
End Sub
