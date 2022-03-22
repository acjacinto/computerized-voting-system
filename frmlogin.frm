VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmlogin 
   BorderStyle     =   0  'None
   Caption         =   "Log In"
   ClientHeight    =   8415
   ClientLeft      =   3510
   ClientTop       =   1530
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdadmin 
      BackColor       =   &H0080FFFF&
      Height          =   235
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7920
      Width           =   5910
   End
   Begin MSAdodcLib.Adodc adoadmin 
      Height          =   330
      Left            =   3720
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      Connect         =   $"frmlogin.frx":0000
      OLEDBString     =   $"frmlogin.frx":008F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from admintable"
      Caption         =   "adoadmin"
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
   Begin VB.CommandButton cmdclose 
      Height          =   495
      Left            =   13080
      Picture         =   "frmlogin.frx":011E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame fmeinstructions 
      BackColor       =   &H000000C0&
      Caption         =   "Instructions:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   240
      TabIndex        =   15
      Top             =   5280
      Width           =   5415
      Begin VB.PictureBox Picture4 
         Height          =   2295
         Left            =   120
         Picture         =   "frmlogin.frx":0679
         ScaleHeight     =   2235
         ScaleWidth      =   5115
         TabIndex        =   16
         Top             =   240
         Width           =   5175
         Begin VB.Label lbl4 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmlogin.frx":3F81
            DataField       =   "          "
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Left            =   720
            TabIndex        =   20
            Top             =   1440
            Width           =   4335
         End
         Begin VB.Label lbl3 
            BackStyle       =   0  'Transparent
            Caption         =   "3. You can vote all candidates on a Selected              Party list by using Vote Straight function."
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   19
            Top             =   840
            Width           =   4335
         End
         Begin VB.Label lbl2 
            BackStyle       =   0  'Transparent
            Caption         =   "2. You will be redirected to Voting Module."
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   18
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label lbl1 
            BackStyle       =   0  'Transparent
            Caption         =   "1. Use valid Student Number and Password."
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   17
            Top             =   240
            Width           =   4095
         End
      End
   End
   Begin VB.PictureBox p3 
      Height          =   6135
      Left            =   5880
      Picture         =   "frmlogin.frx":401F
      ScaleHeight     =   6075
      ScaleWidth      =   7635
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.TextBox txtstudnum 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   3480
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc adologin 
      Height          =   375
      Left            =   1320
      Top             =   8040
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Connect         =   $"frmlogin.frx":1395F
      OLEDBString     =   $"frmlogin.frx":139EE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from RegisterTable"
      Caption         =   "adologin"
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
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00FF8080&
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "frmlogin.frx":13A7D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdregister 
      BackColor       =   &H000080FF&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Picture         =   "frmlogin.frx":13EBB
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -480
      Picture         =   "frmlogin.frx":1441D
      ScaleHeight     =   615
      ScaleWidth      =   14295
      TabIndex        =   2
      Top             =   1440
      Width           =   14295
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   6360
         X2              =   6360
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Label lblhome 
         BackStyle       =   0  'Transparent
         Caption         =   "CVS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   480
         X2              =   14280
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Image Image2 
         Height          =   7185
         Left            =   6360
         Picture         =   "frmlogin.frx":1554B
         Top             =   600
         Width           =   7185
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      Picture         =   "frmlogin.frx":2236C
      ScaleHeight     =   135
      ScaleWidth      =   13815
      TabIndex        =   1
      Top             =   1200
      Width           =   13815
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   13680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   13800
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "frmlogin.frx":2284F
      ScaleHeight     =   255
      ScaleWidth      =   13815
      TabIndex        =   0
      Top             =   8160
      Width           =   13815
   End
   Begin VB.Timer Timer3 
      Interval        =   3500
      Left            =   0
      Top             =   7680
   End
   Begin VB.Label lblpass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblstudnum 
      BackStyle       =   0  'Transparent
      Caption         =   "Student No:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      Height          =   6375
      Left            =   0
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Label lbllogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblpassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   13920
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lbldate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lbltime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   5880
      X2              =   5880
      Y1              =   1800
      Y2              =   8160
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   0  'Transparent
      Height          =   6495
      Left            =   13560
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   13560
      X2              =   13560
      Y1              =   1800
      Y2              =   8160
   End
   Begin VB.Line Line9 
      X1              =   13560
      X2              =   13560
      Y1              =   1800
      Y2              =   8160
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   13800
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   0
      Picture         =   "frmlogin.frx":22D32
      Top             =   0
      Width           =   13665
   End
   Begin VB.Image p1 
      Height          =   6165
      Left            =   5880
      Picture         =   "frmlogin.frx":285CF
      Top             =   2040
      Width           =   7725
   End
   Begin VB.Image p2 
      Height          =   6105
      Left            =   5880
      Picture         =   "frmlogin.frx":356BF
      Top             =   2040
      Visible         =   0   'False
      Width           =   7680
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadmin_Click()
Dim admin As String
Dim password As String
Dim msg As String

adoadmin.Refresh
admin = txtstudnum.Text
password = txtpass.Text

Do Until adoadmin.Recordset.EOF
If adoadmin.Recordset.Fields("Username").Value = admin And adoadmin.Recordset.Fields("Password").Value = password Then
frmlogin.Hide
frmAdminmenu.Show
Exit Sub

Else
adoadmin.Recordset.MoveNext
End If

Loop
msg = MsgBox("Invalid username or password, try again!", vbOKCancel)
If (msg = 1) Then
frmlogin.Show
txtstudnum.Text = ""
txtpass.Text = ""
txtstudnum.SetFocus

Else
End
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdlogin_Click()
Dim studnum As String
Dim pass As String
Dim msg As String

adologin.Refresh
studnum = txtstudnum.Text
pass = txtpass.Text

Do Until adologin.Recordset.EOF
If adologin.Recordset.Fields("StudentNumber").Value = studnum And adologin.Recordset.Fields("Password").Value = pass Then
frmlogin.Hide
frmVotingModule.Show
Exit Sub

Else
adologin.Recordset.MoveNext
End If

Loop

msg = MsgBox("Invalid student number or password, try again!", vbOKCancel)
If (msg = 1) Then
frmlogin.Show
txtstudnum.Text = ""
txtpass.Text = ""
txtstudnum.SetFocus

Else
End
End If
End Sub

Private Sub cmdregister_Click()
frmregister.Show
frmlogin.Hide
End Sub


Private Sub Timer3_Timer()
If p1.Visible = True Then
p3.Visible = True
p1.Visible = False
ElseIf p2.Visible = True Then
p2.Visible = False
p1.Visible = True
ElseIf p3.Visible = True Then
p3.Visible = False
p2.Visible = True
End If
End Sub

