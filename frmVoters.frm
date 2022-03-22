VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVoters 
   BorderStyle     =   0  'None
   Caption         =   "Voters"
   ClientHeight    =   7305
   ClientLeft      =   4695
   ClientTop       =   1920
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVoters.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsearch 
      Height          =   585
      Left            =   10320
      Picture         =   "frmVoters.frx":1BE02
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   4440
      Picture         =   "frmVoters.frx":1C532
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   4920
      Picture         =   "frmVoters.frx":1C914
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   5400
      Picture         =   "frmVoters.frx":1CCBD
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   600
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   3840
      Picture         =   "frmVoters.frx":1D13D
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3480
      Width           =   600
   End
   Begin VB.CommandButton cmdadd 
      Height          =   350
      Left            =   6360
      Picture         =   "frmVoters.frx":1D583
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmddelete 
      Height          =   350
      Left            =   6840
      Picture         =   "frmVoters.frx":1D974
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdsave 
      Height          =   350
      Left            =   7320
      Picture         =   "frmVoters.frx":1DDC1
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   500
   End
   Begin VB.CommandButton cmdback 
      Height          =   735
      Left            =   8640
      Picture         =   "frmVoters.frx":1E1FC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   2655
   End
   Begin VB.ComboBox cmbcourse 
      DataField       =   "Course"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmVoters.frx":1FB48
      Left            =   7080
      List            =   "frmVoters.frx":1FB73
      TabIndex        =   7
      Text            =   "Choose Course"
      Top             =   3960
      Width           =   2895
   End
   Begin VB.ComboBox cmbyear 
      DataField       =   "CYear"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmVoters.frx":1FC04
      Left            =   7080
      List            =   "frmVoters.frx":1FC14
      TabIndex        =   6
      Text            =   "Choose Year"
      Top             =   4440
      Width           =   2895
   End
   Begin VB.ComboBox cmbgender 
      DataField       =   "Gender"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmVoters.frx":1FC24
      Left            =   2400
      List            =   "frmVoters.frx":1FC2E
      TabIndex        =   5
      Text            =   "Choose Gender"
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox txtfname 
      DataField       =   "FirstName"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtlname 
      DataField       =   "LastName"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      DataField       =   "Password"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   7080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox txtstudnum 
      DataField       =   "StudentNumber"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3960
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc votersado 
      Height          =   330
      Left            =   720
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Connect         =   $"frmVoters.frx":1FC40
      OLEDBString     =   $"frmVoters.frx":1FCCF
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "RegisterTable"
      Caption         =   "votersado"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmVoters.frx":1FD5E
      Height          =   2295
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   11468799
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Status"
      DataSource      =   "votersado"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7320
      TabIndex        =   16
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblstat 
      BackColor       =   &H00000000&
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblfname 
      BackColor       =   &H00000000&
      Caption         =   "First Name:"
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
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label lbllname 
      BackColor       =   &H00000000&
      Caption         =   "Last Name:"
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
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label lblcourse 
      BackColor       =   &H00000000&
      Caption         =   "Course:"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Gender:"
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
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblyear 
      BackColor       =   &H00000000&
      Caption         =   "Year:"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblpassword 
      BackColor       =   &H00000000&
      Caption         =   "Password:"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblstudentno 
      BackColor       =   &H00000000&
      Caption         =   "Student Number:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
End
Attribute VB_Name = "frmVoters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
frmVoters.Hide
frmAdminmenu.Show
End Sub
Private Sub cmdadd_Click()
On Error GoTo errormsg
votersado.Recordset.AddNew
MsgBox "Sucessfully Added", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Adding ", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub cmddelete_Click()
votersado.Recordset.Delete
End Sub

Private Sub cmdfirst_Click()
votersado.Recordset.MoveFirst
End Sub

Private Sub cmdlast_Click()
votersado.Recordset.MoveLast
End Sub

Private Sub cmdnext_Click()
votersado.Recordset.MoveNext
End Sub

Private Sub cmdprev_Click()
votersado.Recordset.MovePrevious
End Sub

Private Sub cmdsave_Click()
On Error GoTo errormsg
votersado.Recordset.Fields("Studentnumber") = txtstudnum.Text
votersado.Recordset.Fields("FirstName") = txtfname.Text
votersado.Recordset.Fields("LastName") = txtlname.Text
votersado.Recordset.Fields("Gender") = cmbgender.Text
votersado.Recordset.Fields("Course") = cmbcourse.Text
votersado.Recordset.Fields("CYear") = cmbyear.Text
votersado.Recordset.Fields("Password") = txtpass.Text
votersado.Recordset.Update
MsgBox "Sucessfully Updated", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Updating Information", vbExclamation + vbOKOnly, "Error"
End Sub

Private Sub cmdsearch_Click()
frmSearch.Show
End Sub

