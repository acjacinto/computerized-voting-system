VERSION 5.00
Begin VB.Form frmAdminmenu 
   BorderStyle     =   0  'None
   Caption         =   "Home"
   ClientHeight    =   6975
   ClientLeft      =   4695
   ClientTop       =   1530
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmhome.frx":0000
   ScaleHeight     =   6975
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdadminaccount 
      Height          =   855
      Left            =   1320
      Picture         =   "frmhome.frx":1C79B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton cmdparty 
      Height          =   855
      Left            =   1320
      Picture         =   "frmhome.frx":1E7CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   3975
   End
   Begin VB.CommandButton cmdcandidates 
      Height          =   855
      Left            =   1320
      Picture         =   "frmhome.frx":2033B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton cmdvoters 
      Height          =   855
      Left            =   6120
      Picture         =   "frmhome.frx":220C7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   3975
   End
   Begin VB.CommandButton cmdreports 
      Height          =   855
      Left            =   6120
      Picture         =   "frmhome.frx":23A98
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   3975
   End
   Begin VB.CommandButton cmdsettings 
      Height          =   855
      Left            =   6120
      Picture         =   "frmhome.frx":2555E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   3975
   End
   Begin VB.CommandButton cmdexit 
      Height          =   735
      Left            =   8400
      Picture         =   "frmhome.frx":274B8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAdminmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadminaccount_Click()
frmAdminaccount.Show
frmAdminmenu.Hide
End Sub

Private Sub cmdcandidates_Click()
frmCandidates.Show
frmAdminmenu.Hide
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdparty_Click()
frmPartyList.Show
frmAdminmenu.Hide
End Sub

Private Sub cmdreports_Click()
frmReports.Show
frmAdminmenu.Hide
End Sub

Private Sub cmdsettings_Click()
frmSettings.Show
frmAdminmenu.Hide
End Sub

Private Sub cmdvoters_Click()
frmVoters.Show
frmAdminmenu.Hide
End Sub
