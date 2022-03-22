VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   4890
   ClientTop       =   2115
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
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
      Height          =   390
      Left            =   840
      TabIndex        =   11
      Top             =   1425
      Width           =   2895
   End
   Begin VB.CommandButton cmdrefresh 
      Height          =   350
      Left            =   8760
      Picture         =   "frmSearch.frx":1C2BA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   500
   End
   Begin VB.CommandButton cmdadd 
      Height          =   350
      Left            =   9240
      Picture         =   "frmSearch.frx":1C716
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   500
   End
   Begin VB.CommandButton cmddelete 
      Height          =   350
      Left            =   9720
      Picture         =   "frmSearch.frx":1CB07
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   500
   End
   Begin VB.CommandButton cmdsave 
      Height          =   350
      Left            =   10200
      Picture         =   "frmSearch.frx":1CF54
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   500
   End
   Begin MSAdodcLib.Adodc adosearch 
      Height          =   330
      Left            =   960
      Top             =   6120
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
      Connect         =   $"frmSearch.frx":1D38F
      OLEDBString     =   $"frmSearch.frx":1D41E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from RegisterTable"
      Caption         =   "adosearch"
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
   Begin MSDataGridLib.DataGrid gridsearch 
      Bindings        =   "frmSearch.frx":1D4AD
      Height          =   3615
      Left            =   600
      TabIndex        =   5
      Top             =   2280
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   9961471
      HeadLines       =   1
      RowHeight       =   21
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin VB.CommandButton cmdsearch 
      Height          =   530
      Left            =   10800
      Picture         =   "frmSearch.frx":1D4C5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   530
   End
   Begin VB.ComboBox cmbstatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmSearch.frx":1D9D6
      Left            =   8520
      List            =   "frmSearch.frx":1D9E0
      TabIndex        =   3
      Text            =   "Choose Status"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox cmbyear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmSearch.frx":1D9FA
      Left            =   6480
      List            =   "frmSearch.frx":1DA0A
      TabIndex        =   2
      Text            =   "Choose Year"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbcourse 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmSearch.frx":1DA1A
      Left            =   3960
      List            =   "frmSearch.frx":1DA45
      TabIndex        =   1
      Text            =   "Choose Course"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdclose 
      Height          =   495
      Left            =   10920
      Picture         =   "frmSearch.frx":1DAD6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblstudnum 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Student Number"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1800
      Width           =   2895
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbcourse_Click()
If cmbcourse.Text <> "" Then
adosearch.RecordSource = " select * from RegisterTable where Course= '" & cmbcourse & "'"
adosearch.Refresh
gridsearch.Refresh
End If
End Sub

Private Sub cmbstatus_Click()
If cmbyear.Text <> "" Then
adosearch.RecordSource = " select * from RegisterTable where Status= '" & cmbstatus & "'"
adosearch.Refresh
gridsearch.Refresh
End If
End Sub

Private Sub cmbyear_Click()
If cmbyear.Text <> "" Then
adosearch.RecordSource = " select * from RegisterTable where CYear= '" & cmbyear & "'"
adosearch.Refresh
gridsearch.Refresh
End If
End Sub

Private Sub cmdadd_Click()
adosearch.Recordset.AddNew
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
adosearch.Recordset.Delete
End Sub

Private Sub cmdrefresh_Click()
adosearch.RecordSource = "select * from RegisterTable"
adosearch.Refresh
txtstudnum.Text = ""
cmbcourse.Text = "Choose Course"
cmbyear.Text = "Choose Year"
cmbstatus.Text = "Choose Status"
End Sub

Private Sub cmdsave_Click()
adosearch.Recordset.Update
End Sub

Private Sub cmdsearch_Click()
If txtstudnum.Text <> "" Then
adosearch.RecordSource = " select * from RegisterTable where StudentNumber= '" + txtstudnum + "'"
adosearch.Refresh
gridsearch.Refresh
End If
End Sub


Private Sub Form_Load()
adosearch.Refresh
With adosearch.Recordset
Do Until .EOF
.MoveNext
Loop
End With
End Sub


