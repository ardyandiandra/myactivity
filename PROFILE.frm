VERSION 5.00
Object = "{19883799-9381-41CA-8342-B4546CA216BF}#1.0#0"; "vistaButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PROFILE 
   Caption         =   "PROFILE"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   13800
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt_Id_Semester 
      DataField       =   "ID_SEMESTER"
      DataSource      =   "Adodb_Profile"
      Height          =   285
      Left            =   6000
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Cmb_Semester 
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Txt_Password 
      DataField       =   "PASSWORD"
      DataSource      =   "Adodb_Profile"
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   4560
      Width           =   4335
   End
   Begin VB.TextBox Txt_Username 
      DataField       =   "USERNAME"
      DataSource      =   "Adodb_Profile"
      Height          =   405
      Left            =   3120
      TabIndex        =   5
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox Txt_Universitas 
      DataField       =   "UNIVERSITAS"
      DataSource      =   "Adodb_Profile"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3360
      Width           =   4335
   End
   Begin VB.TextBox Txt_Jurusan 
      DataField       =   "JURUSAN"
      DataSource      =   "Adodb_Profile"
      Height          =   405
      Left            =   3120
      TabIndex        =   3
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox Txt_Angkatan 
      DataField       =   "ANGKATAN"
      DataSource      =   "Adodb_Profile"
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox Txt_Nama 
      DataField       =   "NAMA"
      DataSource      =   "Adodb_Profile"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox Txt_Nim 
      DataField       =   "NIM"
      DataSource      =   "Adodb_Profile"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
   Begin vistaButton.ButtonVista Btnkeluar 
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonShape     =   1
      ButtonStyle     =   7
      BackColor       =   65280
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "KELUAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vistaButton.ButtonVista Btnsimpan 
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonShape     =   1
      ButtonStyle     =   7
      BackColor       =   65280
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "SIMPAN"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodb_Profile 
      Height          =   375
      Left            =   9360
      Top             =   4920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=activity.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=activity.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "PROFILE"
      Caption         =   ""
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   19
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   18
      Top             =   1560
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANGKATAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   17
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JURUSAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   16
      Top             =   2760
      Width           =   1650
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIVERSITAS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   15
      Top             =   3360
      Width           =   2400
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   14
      Top             =   3960
      Width           =   2130
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   13
      Top             =   4560
      Width           =   2010
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEMESTER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   600
      TabIndex        =   12
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROFIL MAHASISWA"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   1500
      TabIndex        =   11
      Top             =   0
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   9600
      Picture         =   "PROFILE.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2760
   End
End
Attribute VB_Name = "PROFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As ADODB.Connection  'this is the connection object
Private rs As ADODB.Recordset   'this is the recordset object

Private Sub Btnkeluar_Click()
    Unload Me
End Sub
Private Sub Form_Resize()
On Error Resume Next
With Image1
.Left = Me.ScaleTop
.Top = Me.ScaleLeft
.Width = Me.ScaleWidth
.Height = Me.ScaleHeight
End With
Err.Clear
End Sub

Private Sub Btnsimpan_Click()
    'On Error GoTo Err:
    Adodb_Profile.Recordset.Fields("NIM") = Txt_Nim.Text
    Adodb_Profile.Recordset.Fields("NAMA") = Txt_Nama.Text
    Adodb_Profile.Recordset.Fields("ANGKATAN") = Txt_Angkatan.Text
    Adodb_Profile.Recordset.Fields("JURUSAN") = Txt_Jurusan.Text
    Adodb_Profile.Recordset.Fields("UNIVERSITAS") = Txt_Universitas.Text
    Adodb_Profile.Recordset.Fields("USERNAME") = Txt_Username.Text
    Adodb_Profile.Recordset.Fields("PASSWORD") = Txt_Password.Text
    Adodb_Profile.Recordset.Fields("ID_SEMESTER") = Txt_Id_Semester.Text
    Adodb_Profile.Recordset.Update
    'MsgBox Cmb_Semester.List(Cmb_Semester.ListIndex)
    MsgBox "Data profile '" & Txt_Nama.Text & "' berhasil disimpan!"
    Exit Sub
Err:
    MsgBox "Terjadi kesalahan penyimpanan"
End Sub

Private Sub Cmb_Semester_Click()
    Txt_Id_Semester.Text = Cmb_Semester.Text
End Sub

Private Sub Form_Load()
    'turn MousePointer to HourGlass to show that we are busy processing
    Me.MousePointer = vbHourglass
    
    'instantiate the connection object
    Set cn = New ADODB.Connection
    'specify the connectionstring
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                          "Data Source=" & App.Path & "\activity.mdb"
    'open the connection
    cn.Open
    
    'instantiate the recordset object
    Set rs = New ADODB.Recordset
    'open the recordset
    With rs
        .Open "SEMESTER", cn, adOpenKeyset, adLockPessimistic, adCmdTable
           
        'loop through the records until reaching the end or last record
        Do While Not .EOF
            Cmb_Semester.AddItem rs.Fields("SEMESTER")
            Cmb_Semester.ItemData(Cmb_Semester.NewIndex) = rs.Fields("ID_SEMESTER")
            rs.MoveNext 'moves next record
        Loop
        
        If Not (.EOF And .BOF) Then
            rs.MoveFirst    'go to the first record if there are existing records
        End If
        
    End With
    
    Me.MousePointer = vbNormal 'sets the mouse pointer to the normal arrow
    
    Cmb_Semester.SelText = Txt_Id_Semester.Text
End Sub
