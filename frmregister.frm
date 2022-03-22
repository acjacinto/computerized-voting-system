VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmregister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   7200
   ClientLeft      =   7380
   ClientTop       =   2295
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmregister.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   6390
   Begin VB.TextBox txtstudnum 
      DataField       =   "StudentNumber"
      DataSource      =   "adoregister"
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
      Left            =   2640
      TabIndex        =   18
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      DataField       =   "Password"
      DataSource      =   "adoregister"
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
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox txtlname 
      DataField       =   "LastName"
      DataSource      =   "adoregister"
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
      Left            =   2640
      TabIndex        =   16
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txtfname 
      DataField       =   "FirstName"
      DataSource      =   "adoregister"
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
      Left            =   2640
      TabIndex        =   15
      Top             =   2040
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc adoregister 
      Height          =   330
      Left            =   3840
      Top             =   6840
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
      Connect         =   $"frmregister.frx":A0EC
      OLEDBString     =   $"frmregister.frx":A17B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "RegisterTable"
      Caption         =   "adoregister"
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
   Begin VB.ComboBox cmbgender 
      DataField       =   "Gender"
      DataSource      =   "adoregister"
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
      ItemData        =   "frmregister.frx":A20A
      Left            =   2640
      List            =   "frmregister.frx":A214
      TabIndex        =   5
      Text            =   "Choose Gender"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ComboBox cmbyear 
      DataField       =   "CYear"
      DataSource      =   "adoregister"
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
      ItemData        =   "frmregister.frx":A226
      Left            =   2640
      List            =   "frmregister.frx":A236
      TabIndex        =   4
      Text            =   "Choose Year"
      Top             =   4440
      Width           =   2895
   End
   Begin VB.ComboBox cmbcourse 
      DataField       =   "Course"
      DataSource      =   "adoregister"
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
      ItemData        =   "frmregister.frx":A246
      Left            =   2640
      List            =   "frmregister.frx":A271
      TabIndex        =   3
      Text            =   "Choose Course"
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton cmdregister 
      BackColor       =   &H0080FF80&
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Picture         =   "frmregister.frx":A302
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdreset 
      BackColor       =   &H0080FFFF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      Picture         =   "frmregister.frx":A77C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student Registration Form"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label lbltitle2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter yout Valid Details for Registration"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label lblstudentno 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblpassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblyear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblcourse 
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lbllname 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblfname 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
frmlogin.Show
frmregister.Hide
End Sub

Private Sub cmdregister_Click()
adoregister.Recordset.Fields("StudentNumber") = txtstudnum.Text
adoregister.Recordset.Fields("FirstName") = txtfname.Text
adoregister.Recordset.Fields("LastName") = txtlname.Text
adoregister.Recordset.Fields("Gender") = cmbgender.Text
adoregister.Recordset.Fields("Course") = cmbcourse.Text
adoregister.Recordset.Fields("CYear") = cmbyear.Text
adoregister.Recordset.Fields("Password") = txtpass.Text
adoregister.Recordset.Fields("Status") = "Unfinished"
adoregister.Recordset.Update
MsgBox "Registration Successful", vbInformation + vbOKOnly, "Verification"
frmlogin.Show
frmregister.Hide
End Sub

Private Sub cmdreset_Click()
txtstudnum.Text = ""
txtfname.Text = ""
txtlname.Text = ""
cmbgender.Text = "Choose Gender"
cmbcourse.Text = "Choose Course"
cmbyear.Text = "Choose Year"
txtpass.Text = ""
End Sub

Private Sub Form_Load()
adoregister.Recordset.AddNew
cmbgender.Text = "Choose Gender"
cmbcourse.Text = "Choose Course"
cmbyear.Text = "Choose Year"
End Sub

Private Sub txtfname_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 64 And KeyAscii <> 46 Then
KeyAscii = 0
End If
End Sub

Private Sub txtlname_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 64 And KeyAscii <> 46 Then
KeyAscii = 0
End If
End Sub

Private Sub txtstudnum_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 44 Or KeyAscii >= 46 _
And KeyAscii <= 37 Or KeyAscii >= 65 And KeyAscii <= 122 Then
KeyAscii = 0
End If
End Sub
