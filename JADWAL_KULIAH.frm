VERSION 5.00
Object = "{19883799-9381-41CA-8342-B4546CA216BF}#1.0#0"; "vistaButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form JADWAL_KULIAH 
   Caption         =   "JADWAL KULIAH"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12720
   BeginProperty Font 
      Name            =   "Algerian"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   12720
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      Height          =   3135
      Left            =   6240
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txt_mata_kuliah 
         DataField       =   "MATA_KULIAH"
         DataSource      =   "Adodb_Jadwal"
         Height          =   390
         Left            =   1320
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_hari 
         DataField       =   "HARI"
         DataSource      =   "Adodb_Jadwal"
         Height          =   390
         Left            =   360
         TabIndex        =   14
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt_id_mata_kuliah 
         DataField       =   "ID_MATA_KULIAH"
         DataSource      =   "Adodb_Jadwal"
         Height          =   390
         Left            =   360
         TabIndex        =   13
         Top             =   2280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt_jam 
         DataField       =   "JAM"
         DataSource      =   "Adodb_Jadwal"
         Height          =   390
         Left            =   2640
         TabIndex        =   12
         Top             =   1680
         Width           =   3375
      End
      Begin VB.ComboBox cmb_mata_kuliah 
         Height          =   390
         Left            =   2640
         TabIndex        =   8
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox cmb_hari 
         Height          =   390
         Left            =   2640
         TabIndex        =   7
         Top             =   1080
         Width           =   3375
      End
      Begin vistaButton.ButtonVista Btnkeluar 
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   2280
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
         Caption         =   "BATAL"
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
         Left            =   2640
         TabIndex        =   5
         Top             =   2280
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JAM"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HARI"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MATA KULIAH"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2370
      End
   End
   Begin vistaButton.ButtonVista Btnhapus 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   4200
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
      Caption         =   "HAPUS"
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
   Begin vistaButton.ButtonVista Btnedit 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   4200
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
      Caption         =   "EDIT"
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
   Begin vistaButton.ButtonVista Btntambah 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4200
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
      Caption         =   "TAMBAH"
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
   Begin MSAdodcLib.Adodc Adodb_Jadwal 
      Height          =   375
      Left            =   5520
      Top             =   4680
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=activity.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=activity.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"JADWAL_KULIAH.frx":0000
      Caption         =   "Adodc1"
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
      Bindings        =   "JADWAL_KULIAH.frx":0092
      Height          =   3015
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "MATA_KULIAH"
         Caption         =   "MATA_KULIAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ID_JADWAL_KULIAH"
         Caption         =   "ID_JADWAL_KULIAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ID_MATA_KULIAH"
         Caption         =   "ID_MATA_KULIAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "HARI"
         Caption         =   "HARI"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "JAM"
         Caption         =   "JAM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   4680
      Left            =   8760
      Picture         =   "JADWAL_KULIAH.frx":00AD
      Stretch         =   -1  'True
      Top             =   960
      Width           =   3240
   End
End
Attribute VB_Name = "JADWAL_KULIAH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btnedit_Click()
    Frame1.Visible = True
    cmb_mata_kuliah.SelText = txt_mata_kuliah
    cmb_hari.SelText = txt_hari.Text
End Sub

Private Sub Btnhapus_Click()
    On Error GoTo err:
    Tanya = MsgBox("Yakin hapus data terpilih?" _
             , vbYesNo + vbQuestion, "Hapus")
            
    If Tanya = vbYes Then
        Adodb_Jadwal.Recordset.Delete
    End If
err:
End Sub

Private Sub Btnsimpan_Click()
    'Validasi
    If cmb_mata_kuliah.Text = "" Then
        MsgBox "Mata Kuliah tidak boleh kosong!"
    ElseIf cmb_hari.Text = "" Then
        MsgBox "Hari tidak boleh kosong!"
    ElseIf txt_jam.Text = "" Then
        MsgBox "Jam tidak boleh kosong!"
        txt_jam.SetFocus
    'Simpan data
    Else
        On Error GoTo err:
        Adodb_Jadwal.Recordset.Fields("ID_MATA_KULIAH") = txt_id_mata_kuliah.Text
        Adodb_Jadwal.Recordset.Fields("HARI") = txt_hari.Text
        Adodb_Jadwal.Recordset.Fields("JAM") = txt_jam.Text
        Adodb_Jadwal.Recordset.Update
        Frame1.Visible = False
        Adodb_Jadwal.Refresh
    End If
err:
End Sub

Private Sub Btntambah_Click()
    Adodb_Jadwal.Recordset.AddNew
    Frame1.Visible = True
End Sub

Private Sub cmb_hari_Click()
    txt_hari.Text = cmb_hari.Text
End Sub

Private Sub cmb_mata_kuliah_Click()
    txt_id_mata_kuliah.Text = cmb_mata_kuliah.ItemData(cmb_mata_kuliah.ListIndex)
End Sub

Private Sub Form_Resize()
On Error Resume Next
With Image1
.Left = Me.ScaleTop
.Top = Me.ScaleLeft
.Width = Me.ScaleWidth
.Height = Me.ScaleHeight
End With
err.Clear
End Sub
Private Sub Btnkeluar_Click()
    Frame1.Visible = False
    Adodb_Jadwal.Recordset.CancelBatch
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
        .Open "MATA_KULIAH", cn, adOpenKeyset, adLockPessimistic, adCmdTable
           
        'loop through the records until reaching the end or last record
        Do While Not .EOF
            cmb_mata_kuliah.AddItem rs.Fields("MATA_KULIAH")
            cmb_mata_kuliah.ItemData(cmb_mata_kuliah.NewIndex) = rs.Fields("ID_MATA_KULIAH")
            rs.MoveNext 'moves next record
        Loop
        
        If Not (.EOF And .BOF) Then
            rs.MoveFirst    'go to the first record if there are existing records
        End If
        
    End With
    
    Me.MousePointer = vbNormal 'sets the mouse pointer to the normal arrow
    
    'Hari
    cmb_hari.AddItem ("SENIN")
    cmb_hari.AddItem ("SELASA")
    cmb_hari.AddItem ("RABU")
    cmb_hari.AddItem ("KAMIS")
    cmb_hari.AddItem ("JUM'AT")
    cmb_hari.AddItem ("SABTU")
    cmb_hari.AddItem ("MINGGU")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Adodb_Jadwal.Recordset.CancelBatch
End Sub
