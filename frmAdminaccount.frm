VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAdminaccount 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   4515
   ClientTop       =   2310
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAdminaccount.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtpass 
      DataField       =   "Password"
      DataSource      =   "adminado"
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
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtname 
      DataField       =   "Name"
      DataSource      =   "adminado"
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
      Left            =   7800
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtuser 
      DataField       =   "Username"
      DataSource      =   "adminado"
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
      Left            =   7800
      TabIndex        =   9
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmdsave 
      Height          =   350
      Left            =   4560
      Picture         =   "frmAdminaccount.frx":1D198
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   500
   End
   Begin VB.CommandButton cmdback 
      Height          =   735
      Left            =   8400
      Picture         =   "frmAdminaccount.frx":1D5D3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CommandButton cmddelete 
      Height          =   350
      Left            =   4080
      Picture         =   "frmAdminaccount.frx":1EF1F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   500
   End
   Begin VB.CommandButton cmdadd 
      Height          =   350
      Left            =   3600
      Picture         =   "frmAdminaccount.frx":1F36C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   500
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   1080
      Picture         =   "frmAdminaccount.frx":1F75D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   600
   End
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   2640
      Picture         =   "frmAdminaccount.frx":1FBA3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   600
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   2160
      Picture         =   "frmAdminaccount.frx":20023
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   500
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   1680
      Picture         =   "frmAdminaccount.frx":203CC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   500
   End
   Begin MSAdodcLib.Adodc adminado 
      Height          =   375
      Left            =   2640
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmAdminaccount.frx":207AE
      OLEDBString     =   $"frmAdminaccount.frx":2083D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "admintable"
      Caption         =   "adminado"
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
   Begin MSDataGridLib.DataGrid admingrid 
      Bindings        =   "frmAdminaccount.frx":208CC
      Height          =   3975
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   11206655
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   24
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   0
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblpass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      TabIndex        =   14
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lbluser 
      BackStyle       =   0  'Transparent
      Caption         =   "Usename:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblname 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdminaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdadd_Click()
On Error GoTo errormsg
adminado.Recordset.AddNew
MsgBox "Sucessfully Added", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Adding ", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub cmdback_Click()
frmAdminmenu.Show
frmAdminaccount.Hide
End Sub

Private Sub cmddelete_Click()
adminado.Recordset.Delete
End Sub

Private Sub cmdfirst_Click()
adminado.Recordset.MoveFirst
End Sub

Private Sub cmdlast_Click()
adminado.Recordset.MoveLast
End Sub

Private Sub cmdnext_Click()
adminado.Recordset.MoveNext
End Sub

Private Sub cmdprev_Click()
adminado.Recordset.MovePrevious
End Sub

Private Sub cmdsave_Click()
On Error GoTo errormsg
adminado.Recordset.Fields("Username") = txtuser.Text
adminado.Recordset.Fields("Name") = txtname.Text
adminado.Recordset.Fields("Password") = txtpass.Text
adminado.Recordset.Update
MsgBox "Sucessfully Updated", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Updating Information", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub Form_Load()
adminado.Visible = False
adminado.Recordset.AddNew
End Sub


