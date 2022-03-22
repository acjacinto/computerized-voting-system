VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPartyList 
   BorderStyle     =   0  'None
   Caption         =   "Party List"
   ClientHeight    =   7185
   ClientLeft      =   4695
   ClientTop       =   1725
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPartyList.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdfirst 
      Height          =   350
      Left            =   3240
      Picture         =   "frmPartyList.frx":1AA7A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2640
      Width           =   600
   End
   Begin VB.CommandButton cmdlast 
      Height          =   350
      Left            =   4800
      Picture         =   "frmPartyList.frx":1AEC0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Width           =   600
   End
   Begin VB.CommandButton cmdnext 
      Height          =   350
      Left            =   4320
      Picture         =   "frmPartyList.frx":1B340
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   500
   End
   Begin VB.CommandButton cmdprev 
      Height          =   350
      Left            =   3840
      Picture         =   "frmPartyList.frx":1B6E9
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   500
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox piclogo 
      Height          =   2700
      Left            =   7320
      ScaleHeight     =   2640
      ScaleWidth      =   2715
      TabIndex        =   14
      Top             =   3120
      Width           =   2780
      Begin VB.Image imglogo 
         Height          =   2415
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSAdodcLib.Adodc partyado 
      Height          =   330
      Left            =   1920
      Top             =   6120
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Connect         =   $"frmPartyList.frx":1BACB
      OLEDBString     =   $"frmPartyList.frx":1BB5A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "partylisttable"
      Caption         =   "partyado"
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
   Begin MSDataGridLib.DataGrid partygrid 
      Bindings        =   "frmPartyList.frx":1BBE9
      Height          =   1695
      Left            =   1440
      TabIndex        =   12
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   10485759
      HeadLines       =   1
      RowHeight       =   21
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
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
   Begin VB.CommandButton cmdadd 
      Height          =   350
      Left            =   5640
      Picture         =   "frmPartyList.frx":1BC00
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   500
   End
   Begin VB.CommandButton cmddelete 
      Height          =   350
      Left            =   6120
      Picture         =   "frmPartyList.frx":1BFF1
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   500
   End
   Begin VB.CommandButton cmdsave 
      Height          =   350
      Left            =   6600
      Picture         =   "frmPartyList.frx":1C43E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   500
   End
   Begin VB.CommandButton cmdload 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10200
      Picture         =   "frmPartyList.frx":1C879
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtparty 
      DataField       =   "Party"
      DataSource      =   "partyado"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox txtinfo 
      DataField       =   "Information"
      DataSource      =   "partyado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox txtdescript 
      DataField       =   "Description"
      DataSource      =   "partyado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   3960
      Width           =   3255
   End
   Begin VB.CommandButton cmdback 
      Height          =   735
      Left            =   8520
      Picture         =   "frmPartyList.frx":1D11C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label lblpath 
      DataField       =   "LOGO"
      DataSource      =   "partyado"
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblpartylogo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PARTY LOGO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Information:"
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
      Left            =   960
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblparty 
      BackStyle       =   0  'Transparent
      Caption         =   "PARTY:"
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
      Left            =   960
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lbldescript 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "frmPartyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadd_Click()
partyado.Recordset.AddNew
End Sub

Private Sub cmdback_Click()
frmAdminmenu.Show
frmPartyList.Hide
End Sub

Private Sub cmddelete_Click()
partyado.Recordset.Delete
End Sub

Private Sub cmdfirst_Click()
partyado.Recordset.MoveFirst
imglogo.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdlast_Click()
partyado.Recordset.MoveLast
imglogo.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdload_Click()
CommonDialog1.FileName = ""
CommonDialog1.Filter = "JPEG files|*.jpg|GIF Files|*.gif|All Files|*.*"
CommonDialog1.ShowOpen
lblpath = CommonDialog1.FileName

   If Len(Trim(lblpath)) < 1 Then
      Exit Sub
   End If
imglogo.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdnext_Click()
partyado.Recordset.MoveNext
imglogo.Picture = LoadPicture(lblpath)
End Sub

Private Sub cmdprev_Click()
partyado.Recordset.MovePrevious
imglogo.Picture = LoadPicture(lblpath)
End Sub


Private Sub cmdsave_Click()
On Error GoTo errormsg
partyado.Recordset.Fields("Party") = txtparty.Text
partyado.Recordset.Fields("Description") = txtdescript.Text
partyado.Recordset.Fields("Information") = txtinfo.Text
partyado.Recordset.Fields("LOGO") = "" & lblpath.Caption
partyado.Recordset.Update
MsgBox "Sucessfully Updated", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Updating Information", vbExclamation + vbOKOnly, "Error"
End Sub



