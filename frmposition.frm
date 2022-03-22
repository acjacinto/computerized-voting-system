VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmposition 
   BorderStyle     =   0  'None
   Caption         =   "frmposition"
   ClientHeight    =   7290
   ClientLeft      =   10530
   ClientTop       =   1725
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmposition.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsave 
      Height          =   350
      Left            =   7440
      Picture         =   "frmposition.frx":15E27
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   500
   End
   Begin VB.CommandButton cmddelete 
      Height          =   350
      Left            =   6840
      Picture         =   "frmposition.frx":16262
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   500
   End
   Begin VB.CommandButton cmdadd 
      Height          =   350
      Left            =   6240
      Picture         =   "frmposition.frx":166AF
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   500
   End
   Begin VB.TextBox txtpro 
      DataField       =   "PRO"
      DataSource      =   "adoposition"
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
      Left            =   5760
      TabIndex        =   12
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox txtauditor 
      DataField       =   "Auditor"
      DataSource      =   "adoposition"
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
      Left            =   5760
      TabIndex        =   10
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txttreasurer 
      DataField       =   "Treasurer"
      DataSource      =   "adoposition"
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
      Left            =   5760
      TabIndex        =   8
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txtsec 
      DataField       =   "Secretary"
      DataSource      =   "adoposition"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox txtvpres 
      DataField       =   "Vice President"
      DataSource      =   "adoposition"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txtpres 
      DataField       =   "President"
      DataSource      =   "adoposition"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton cmdclose 
      Height          =   495
      Left            =   7920
      Picture         =   "frmposition.frx":16AA0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adoposition 
      Height          =   375
      Left            =   5880
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   $"frmposition.frx":16FFB
      OLEDBString     =   $"frmposition.frx":1708A
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
   Begin MSDataGridLib.DataGrid gridposition 
      Bindings        =   "frmposition.frx":17119
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   10485759
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
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   " PRO:"
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
      Left            =   4560
      TabIndex        =   13
      Top             =   5535
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   " Auditor:"
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
      Left            =   4560
      TabIndex        =   11
      Top             =   4935
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   " Treasurer:"
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
      Left            =   4560
      TabIndex        =   9
      Top             =   4335
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   " Secretary:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   5535
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " VicePresident:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   4935
      Width           =   1695
   End
   Begin VB.Label lbluser 
      BackColor       =   &H00000000&
      Caption         =   " President:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   4335
      Width           =   1695
   End
End
Attribute VB_Name = "frmposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
adoposition.Recordset.AddNew
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
adoposition.Recordset.Delete
End Sub

Private Sub cmdsave_Click()
On Error GoTo errormsg
adoposition.Recordset.Fields("President") = txtpres.Text
adoposition.Recordset.Fields("Vice President") = txtvpres.Text
adoposition.Recordset.Fields("Secretary") = txtsec.Text
adoposition.Recordset.Fields("Treasurer") = txttreasurer.Text
adoposition.Recordset.Fields("Auditor") = txtauditor.Text
adoposition.Recordset.Fields("PRO") = txtpro.Text
adoposition.Recordset.Update
MsgBox "Sucessfully Updated", vbInformation + vbOKOnly, "Verification"
Exit Sub
errormsg:
MsgBox "Error in Updating Information", vbExclamation + vbOKOnly, "Error"
End Sub
