VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVotingModule 
   BorderStyle     =   0  'None
   ClientHeight    =   9240
   ClientLeft      =   2445
   ClientTop       =   1155
   ClientWidth     =   15555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVotingModule.frx":0000
   ScaleHeight     =   9240
   ScaleWidth      =   15555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmeVotersInfo 
      BackColor       =   &H000000C0&
      Caption         =   "Voter's Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   240
      TabIndex        =   37
      Top             =   1440
      Width           =   6615
      Begin VB.TextBox txtlname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   44
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   43
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtstudnum 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   42
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdgenerate 
         Height          =   855
         Left            =   4560
         Picture         =   "frmVotingModule.frx":1C233
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   40
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdvc6 
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
      Height          =   350
      Left            =   12840
      Picture         =   "frmVotingModule.frx":1D316
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdvc5 
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
      Height          =   350
      Left            =   10080
      Picture         =   "frmVotingModule.frx":1DD85
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdvc4 
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
      Height          =   350
      Left            =   7320
      Picture         =   "frmVotingModule.frx":1E7F4
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdvc3 
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
      Height          =   350
      Left            =   12840
      Picture         =   "frmVotingModule.frx":1F263
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdvc2 
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
      Height          =   350
      Left            =   10080
      Picture         =   "frmVotingModule.frx":1FCD2
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdvc1 
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
      Height          =   350
      Left            =   7320
      Picture         =   "frmVotingModule.frx":20741
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc adoposition 
      Height          =   495
      Left            =   10320
      Top             =   240
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      Connect         =   $"frmVotingModule.frx":211B0
      OLEDBString     =   $"frmVotingModule.frx":2123F
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
   Begin VB.Frame Frame6 
      BackColor       =   &H000000C0&
      Caption         =   "P.R.O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   12600
      TabIndex        =   21
      Top             =   5160
      Width           =   2655
      Begin VB.ComboBox cmbpro 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVotingModule.frx":212CE
         Left            =   240
         List            =   "frmVotingModule.frx":212D0
         TabIndex        =   23
         Text            =   "Choose P.R.O"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.PictureBox Picture6 
         Height          =   1980
         Left            =   360
         ScaleHeight     =   1920
         ScaleWidth      =   1935
         TabIndex        =   22
         Top             =   360
         Width           =   1995
         Begin VB.Image imgpro 
            Height          =   1935
            Left            =   0
            Picture         =   "frmVotingModule.frx":212D2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H000000C0&
      Caption         =   "Auditor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   9840
      TabIndex        =   18
      Top             =   5160
      Width           =   2655
      Begin VB.ComboBox cmbauditor 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVotingModule.frx":26956
         Left            =   240
         List            =   "frmVotingModule.frx":26958
         TabIndex        =   20
         Text            =   "Choose Auditor"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.PictureBox Picture5 
         Height          =   1980
         Left            =   360
         ScaleHeight     =   1920
         ScaleWidth      =   1935
         TabIndex        =   19
         Top             =   360
         Width           =   1995
         Begin VB.Image imgauditor 
            Height          =   1935
            Left            =   0
            Picture         =   "frmVotingModule.frx":2695A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H000000C0&
      Caption         =   "Treasurer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   7080
      TabIndex        =   15
      Top             =   5160
      Width           =   2655
      Begin VB.ComboBox cmbtreasurer 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVotingModule.frx":2BFDE
         Left            =   240
         List            =   "frmVotingModule.frx":2BFE0
         TabIndex        =   17
         Text            =   "Choose Treasurer"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.PictureBox Picture4 
         Height          =   1980
         Left            =   360
         ScaleHeight     =   1920
         ScaleWidth      =   1935
         TabIndex        =   16
         Top             =   360
         Width           =   1995
         Begin VB.Image imgtreasurer 
            Height          =   1935
            Left            =   0
            Picture         =   "frmVotingModule.frx":2BFE2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000C0&
      Caption         =   "Secretary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   12600
      TabIndex        =   12
      Top             =   1440
      Width           =   2655
      Begin VB.ComboBox cmbsec 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVotingModule.frx":31666
         Left            =   240
         List            =   "frmVotingModule.frx":31668
         TabIndex        =   14
         Text            =   "Choose Secretary"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         Height          =   1980
         Left            =   360
         ScaleHeight     =   1920
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   360
         Width           =   1995
         Begin VB.Image imgsec 
            Height          =   1935
            Left            =   0
            Picture         =   "frmVotingModule.frx":3166A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "Vice President"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   9840
      TabIndex        =   9
      Top             =   1440
      Width           =   2655
      Begin VB.ComboBox cmbvpres 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVotingModule.frx":36CEE
         Left            =   240
         List            =   "frmVotingModule.frx":36CF0
         TabIndex        =   11
         Text            =   "Choose Vice President"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         Height          =   1980
         Left            =   360
         ScaleHeight     =   1920
         ScaleWidth      =   1935
         TabIndex        =   10
         Top             =   360
         Width           =   1995
         Begin VB.Image imgvpres 
            Height          =   1935
            Left            =   0
            Picture         =   "frmVotingModule.frx":36CF2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000000C0&
      Caption         =   "President"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   7080
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
      Begin VB.PictureBox Picture3 
         Height          =   1980
         Left            =   360
         ScaleHeight     =   1920
         ScaleWidth      =   1935
         TabIndex        =   7
         Top             =   360
         Width           =   1995
         Begin VB.Image imgpres 
            Height          =   1935
            Left            =   0
            Picture         =   "frmVotingModule.frx":3C376
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbpres 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmVotingModule.frx":419FA
         Left            =   240
         List            =   "frmVotingModule.frx":419FC
         TabIndex        =   6
         Text            =   "Choose President"
         Top             =   2520
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid gridcandidates 
      Bindings        =   "frmVotingModule.frx":419FE
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1296
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc adocandidates 
      Height          =   495
      Left            =   7320
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
      Connect         =   $"frmVotingModule.frx":41A1A
      OLEDBString     =   $"frmVotingModule.frx":41AA9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from candidatestable"
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
   Begin VB.Frame fmeBallot 
      BackColor       =   &H000000C0&
      Caption         =   "Voting Ballot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   6615
      Begin VB.ListBox lstballot 
         Columns         =   1
         DataField       =   "FullName"
         DataSource      =   "adocandid"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         ItemData        =   "frmVotingModule.frx":41B38
         Left            =   2640
         List            =   "frmVotingModule.frx":41B3A
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton cmdfinish 
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
         Height          =   975
         Left            =   1440
         Picture         =   "frmVotingModule.frx":41B3C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3960
         Width           =   3975
      End
      Begin VB.Label lblcandid 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CANDIDATES"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   495
         Left            =   2640
         TabIndex        =   30
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblpro 
         BackStyle       =   0  'Transparent
         Caption         =   "P.R.O:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   29
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblaudit 
         BackStyle       =   0  'Transparent
         Caption         =   "Auditor:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lbltrea 
         BackStyle       =   0  'Transparent
         Caption         =   "Treasurer:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblsec 
         BackStyle       =   0  'Transparent
         Caption         =   "Secretary:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   26
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblvpres 
         BackStyle       =   0  'Transparent
         Caption         =   "Vice President:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblpres 
         BackStyle       =   0  'Transparent
         Caption         =   "President:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdclose 
      Height          =   495
      Left            =   14880
      Picture         =   "frmVotingModule.frx":43ACE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adologin 
      Height          =   375
      Left            =   10320
      Top             =   840
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
      Connect         =   $"frmVotingModule.frx":44029
      OLEDBString     =   $"frmVotingModule.frx":440B8
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
   Begin VB.Label lblpath 
      DataField       =   "Picture"
      DataSource      =   "adocandidates"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmVotingModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbparty_Click()
If cmbparty.Text <> "" Then
adopartylist.RecordSource = " select * from partylisttable where Description= '" & cmbparty & "'"
adopartylist.Refresh
gridparty.Refresh
imglogo.Picture = LoadPicture(lblpath)
End If
End Sub

Private Sub cmbauditor_Click()
If cmbauditor.Text <> "" Then
adocandidates.RecordSource = " select * from candidatestable where FullName= '" & cmbauditor & "'"
adocandidates.Refresh
gridcandidates.Refresh
imgauditor.Picture = LoadPicture(lblpath)
End If
If cmbauditor.Text = adocandidates.Recordset.Fields("FullName") Then
   adocandidates.Recordset.Fields("Votes") = adocandidates.Recordset.Fields("Votes") + 1
End If
End Sub

Private Sub cmbpres_Click()
If cmbpres.Text <> "" Then
adocandidates.RecordSource = " select * from candidatestable where FullName= '" & cmbpres & "'"
adocandidates.Refresh
gridcandidates.Refresh
imgpres.Picture = LoadPicture(lblpath)
End If

If cmbpres.Text = adocandidates.Recordset.Fields("FullName") Then
   adocandidates.Recordset.Fields("Votes") = adocandidates.Recordset.Fields("Votes") + 1
   cmbpres.Enabled = True
   cmbvpres.Enabled = False
   cmbsec.Enabled = False
   cmbtreasurer.Enabled = False
   cmbauditor.Enabled = False
   cmbpro.Enabled = False
End If
End Sub

Private Sub cmbpro_Click()
If cmbpro.Text <> "" Then
adocandidates.RecordSource = " select * from candidatestable where FullName= '" & cmbpro & "'"
adocandidates.Refresh
gridcandidates.Refresh
imgpro.Picture = LoadPicture(lblpath)
End If
If cmbpro.Text = adocandidates.Recordset.Fields("FullName") Then
   adocandidates.Recordset.Fields("Votes") = adocandidates.Recordset.Fields("Votes") + 1
End If
End Sub

Private Sub cmbsec_Click()
If cmbsec.Text <> "" Then
adocandidates.RecordSource = " select * from candidatestable where FullName= '" & cmbsec & "'"
adocandidates.Refresh
gridcandidates.Refresh
imgsec.Picture = LoadPicture(lblpath)
End If
If cmbsec.Text = adocandidates.Recordset.Fields("FullName") Then
    adocandidates.Recordset.Fields("Votes") = adocandidates.Recordset.Fields("Votes") + 1
End If
End Sub

Private Sub cmbtreasurer_Click()
If cmbtreasurer.Text <> "" Then
adocandidates.RecordSource = " select * from candidatestable where FullName= '" & cmbtreasurer & "'"
adocandidates.Refresh
gridcandidates.Refresh
imgtreasurer.Picture = LoadPicture(lblpath)
End If
If cmbtreasurer.Text = adocandidates.Recordset.Fields("FullName") Then
   adocandidates.Recordset.Fields("Votes") = adocandidates.Recordset.Fields("Votes") + 1
End If
End Sub

Private Sub cmbvpres_Click()
If cmbvpres.Text <> "" Then
adocandidates.RecordSource = " select * from candidatestable where FullName= '" & cmbvpres & "'"
adocandidates.Refresh
gridcandidates.Refresh
imgvpres.Picture = LoadPicture(lblpath)
End If
If cmbvpres.Text = adocandidates.Recordset.Fields("FullName") Then
   adocandidates.Recordset.Fields("Votes") = adocandidates.Recordset.Fields("Votes") + 1
End If
End Sub




Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdfinish_Click()
cmbpro.Enabled = False
cmdvc1.Enabled = False
cmdvc2.Enabled = False
cmdvc3.Enabled = False
cmdvc4.Enabled = False
cmdvc5.Enabled = False
cmdvc6.Enabled = False
adologin.Recordset.Fields("Status") = "Finished"
adologin.Recordset.Update
MsgBox "Voting Successful"
txtstudnum.Text = ""
txtname.Text = ""
txtlname.Text = ""
txtstudnum.SetFocus
End Sub

Private Sub cmdtransfer_Click()






End Sub

Private Sub cmdgenerate_Click()
Dim studnum As String
Dim pass As String
Dim msg As String

adologin.Refresh
studnum = txtstudnum.Text
fname = txtname.Text
lname = txtlname.Text


Do Until adologin.Recordset.EOF
If adologin.Recordset.Fields("StudentNumber").Value = studnum And adologin.Recordset.Fields("FirstName").Value = fname _
   And adologin.Recordset.Fields("LastName").Value = lname And adologin.Recordset.Fields("Status").Value = "Unfinished" Then
   cmdvc1.Visible = True
   cmdvc2.Visible = True
   cmdvc3.Visible = True
   cmdvc4.Visible = True
   cmdvc5.Visible = True
   cmdvc6.Visible = True
Exit Sub

Else
adologin.Recordset.MoveNext
End If

Loop

msg = MsgBox("       You've already Voted!", vbOKCancel, "Verification")
If (msg = 1) Then
frmlogin.Show
txtstudnum.Text = ""
txtname.Text = ""
txtlname.Text = ""
txtstudnum.SetFocus

Else
End
End If
End Sub

Private Sub cmdvc1_Click()
cmbpres.Enabled = False
cmbvpres.Enabled = True
lstballot.AddItem cmbpres.Text
adocandidates.Recordset.Update
End Sub

Private Sub cmdvc2_Click()
cmbvpres.Enabled = False
cmbsec.Enabled = True
lstballot.AddItem cmbvpres.Text
adocandidates.Recordset.Update
End Sub

Private Sub cmdvc3_Click()
cmbsec.Enabled = False
cmbtreasurer.Enabled = True
lstballot.AddItem cmbsec.Text
adocandidates.Recordset.Update
End Sub

Private Sub cmdvc4_Click()
cmbtreasurer.Enabled = False
cmbauditor.Enabled = True
lstballot.AddItem cmbtreasurer.Text
adocandidates.Recordset.Update
End Sub

Private Sub cmdvc5_Click()
cmbauditor.Enabled = False
cmbpro.Enabled = True
lstballot.AddItem cmbauditor.Text
adocandidates.Recordset.Update
End Sub

Private Sub cmdvc6_Click()
lstballot.AddItem cmbpro.Text
adocandidates.Recordset.Update
End Sub


Private Sub Form_Load()
adoposition.Refresh
With adoposition.Recordset
Do Until .EOF
cmbpres.AddItem ![President]
cmbvpres.AddItem ![Vice President]
cmbsec.AddItem ![Secretary]
cmbtreasurer.AddItem ![Treasurer]
cmbauditor.AddItem ![Auditor]
cmbpro.AddItem ![PRO]
.MoveNext
Loop
End With

End Sub

Private Sub lblpath_Click()
imglogo.Picture = LoadPicture(lblpath)
End Sub


