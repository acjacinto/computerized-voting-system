VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCandidates 
   BorderStyle     =   0  'None
   Caption         =   "Candidates"
   ClientHeight    =   7305
   ClientLeft      =   2820
   ClientTop       =   1725
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCandidates.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdposition 
      Height          =   585
      Left            =   9000
      Picture         =   "frmCandidates.frx":1D543
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5160
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adoposition 
      Height          =   330
      Left            =   5640
      Top             =   480
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
      Connect         =   $"frmCandidates.frx":1DA85
      OLEDBString     =   $"frmCandidates.frx":1DB14
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "positiontable"
      Caption         =   "adoposition"
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
   Begin VB.TextBox txtparty 
      DataField       =   "Party"
      DataSource      =   "adocandidates"
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
      Left            =   6120
      TabIndex        =   25
      Top             =   5160
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10680
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      Picture         =   "frmCandidates.frx":1DBA3
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4200
      Width           =   2535
   End
   Begin VB.PictureBox piccandidate 
      BackColor       =   &H00C0C0C0&
      Height          =   2935
      Left            =   7800
      ScaleHeight     =   2880
      ScaleWidth      =   3090
      TabIndex        =   22
      Top             =   1200
      Width           =   3150
      Begin VB.Image imgcandidate 
         Height          =   2655
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.TextBox txtinfo 
      DataField       =   "Info"
      DataSource      =   "adocandidates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   20
      Top             =   6120
      Width           =   2775
   End
   Begin VB.ComboBox cmbposition 
      DataField       =   "Position"
      DataSource      =   "adocandidates"
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
      ItemData        =   "frmCandidates.frx":1E587
      Left            =   6120
      List            =   "frmCandidates.frx":1E59D
      TabIndex        =   17
      Text            =   "Choose Position"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.ComboBox cmbyear 
      DataField       =   "Year"
      DataSource      =   "adocandidates"
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
      ItemData        =   "frmCandidates.frx":1E5E0
      Left            =   1680
      List            =   "frmCandidates.frx":1E5F0
      TabIndex        =   14
      Text            =   "Choose Year"
      Top             =   6600
      Width           =   2775
   End
   Begin VB.ComboBox cmbcourse 
      DataField       =   "Course"
      DataSource      =   "adocandidates"
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
      ItemData        =   "frmCandidates.frx":1E600
      Left            =   1680
      List            =   "frmCandidates.frx":1E62B
      TabIndex        =   13
      Text            =   "Choose Course"
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox txtfullname 
      DataField       =   "FullName"
      DataSource      =   "adocandidates"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox txtID 
      DataField       =   "CandidateID"
      DataSource      =   "adocandidates"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   5145
      Width           =   2775
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   2160
      Picture         =   "frmCandidates.frx":1E6BC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   2640
      Picture         =   "frmCandidates.frx":1EA9E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   3120
      Picture         =   "frmCandidates.frx":1EE47
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   600
   End
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   1560
      Picture         =   "frmCandidates.frx":1F2C7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   600
   End
   Begin VB.CommandButton cmdadd 
      Height          =   350
      Left            =   3960
      Picture         =   "frmCandidates.frx":1F70D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton cmddelete 
      Height          =   350
      Left            =   4440
      Picture         =   "frmCandidates.frx":1FAFE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   500
   End
   Begin VB.CommandButton cmdsave 
      Height          =   350
      Left            =   4920
      Picture         =   "frmCandidates.frx":1FF4B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   500
   End
   Begin MSAdodcLib.Adodc adocandidates 
      Height          =   330
      Left            =   3000
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   $"frmCandidates.frx":20386
      OLEDBString     =   $"frmCandidates.frx":20415
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "candidatestable"
      Caption         =   "adocandidates"
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
   Begin VB.CommandButton cmdback 
      Height          =   735
      Left            =   9120
      Picture         =   "frmCandidates.frx":204A4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid candidategrid 
      Bindings        =   "frmCandidates.frx":21DF0
      Height          =   3495
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6165
      _Version        =   393216
      BackColor       =   9699327
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
   Begin VB.Label lblpath 
      DataField       =   "Picture"
      DataSource      =   "adocandidates"
      Height          =   255
      Left            =   7800
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblinfo 
      BackColor       =   &H00000000&
      Caption         =   "Information:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Position:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Party:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblyear 
      BackColor       =   &H00000000&
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblcourse 
      BackColor       =   &H00000000&
      Caption         =   "Course:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Full Name:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label lbluser 
      BackColor       =   &H00000000&
      Caption         =   "Candidate ID:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5160
      Width           =   1455
   End
End
Attribute VB_Name = "frmCandidates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadd_Click()
adocandidates.Recordset.AddNew
MsgBox "Sucessfully Added", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Adding ", vbExclamation + vbOKOnly, "Error"
End Sub



Private Sub cmdback_Click()
frmAdminmenu.Show
frmCandidates.Hide
End Sub

Private Sub cmddelete_Click()
adocandidates.Recordset.Delete
End Sub

Private Sub cmdfirst_Click()
sadocandidates.Recordset.MoveFirst
imgcandidate.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdlast_Click()
adocandidates.Recordset.MoveLast
imgcandidate.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdload_Click()
CommonDialog1.FileName = ""
CommonDialog1.Filter = "JPEG files|*.jpg|GIF Files|*.gif|All Files|*.*"
CommonDialog1.ShowOpen
lblpath = CommonDialog1.FileName

   If Len(Trim(lblpath)) < 1 Then
      Exit Sub
   End If
imgcandidate.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdnext_Click()
adocandidates.Recordset.MoveNext
imgcandidate.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdposition_Click()
frmposition.Show
End Sub

Private Sub cmdprev_Click()
adocandidates.Recordset.MovePrevious
imgcandidate.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdsave_Click()

On Error GoTo errormsg
adocandidates.Recordset.Fields("CandidateID") = txtID.Text
adocandidates.Recordset.Fields("FullName") = txtfullname.Text
adocandidates.Recordset.Fields("Course") = cmbcourse.Text
adocandidates.Recordset.Fields("Year") = cmbyear.Text
adocandidates.Recordset.Fields("Party") = txtparty.Text
adocandidates.Recordset.Fields("Position") = cmbposition.Text
adocandidates.Recordset.Fields("Info") = txtinfo.Text
adocandidates.Recordset.Fields("Picture") = "" & lblpath.Caption
adocandidates.Recordset.Update
MsgBox "Sucessfully Updated", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Updating Information", vbExclamation + vbOKOnly, "Error"
End Sub



