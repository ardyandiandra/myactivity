VERSION 5.00
Object = "{19883799-9381-41CA-8342-B4546CA216BF}#1.0#0"; "vistaButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form MATA_KULIAH 
   Caption         =   "MATA KULIAH"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      Height          =   3495
      Left            =   6840
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txt_dosen 
         DataField       =   "NAMA"
         DataSource      =   "Adodb_mata_kuliah"
         Height          =   285
         Left            =   5160
         TabIndex        =   19
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txt_semester 
         DataField       =   "SEMESTER"
         DataSource      =   "Adodb_mata_kuliah"
         Height          =   285
         Left            =   4680
         TabIndex        =   18
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txt_id_dosen 
         DataField       =   "ID_DOSEN"
         DataSource      =   "Adodb_mata_kuliah"
         Height          =   285
         Left            =   4200
         TabIndex        =   17
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txt_id_semester 
         DataField       =   "ID_SEMESTER"
         DataSource      =   "Adodb_mata_kuliah"
         Height          =   285
         Left            =   3720
         TabIndex        =   16
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txt_mata_kuliah 
         DataField       =   "MATA_KULIAH"
         DataSource      =   "Adodb_mata_kuliah"
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txt_sks 
         DataField       =   "SKS"
         DataSource      =   "Adodb_mata_kuliah"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   2040
         Width           =   3495
      End
      Begin VB.ComboBox cmb_semester 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   1080
         Width           =   3495
      End
      Begin VB.ComboBox cmb_dosen 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1560
         Width           =   3495
      End
      Begin vistaButton.ButtonVista Btnkeluar 
         Height          =   495
         Left            =   1800
         TabIndex        =   14
         Top             =   2760
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
         Left            =   360
         TabIndex        =   15
         Top             =   2760
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SKS"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOSEN"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEMESTER"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Height          =   390
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2370
      End
   End
   Begin vistaButton.ButtonVista Btnhapus 
      Height          =   495
      Left            =   3120
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
      Left            =   1800
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
      Left            =   480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MATA_KULIAH.frx":0000
      Height          =   2655
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
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
      ColumnCount     =   7
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "ID_SEMESTER"
         Caption         =   "ID_SEMESTER"
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
         DataField       =   "ID_DOSEN"
         Caption         =   "ID_DOSEN"
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
         DataField       =   "SKS"
         Caption         =   "SKS"
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
      BeginProperty Column05 
         DataField       =   "NAMA"
         Caption         =   "NAMA DOSEN"
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
      BeginProperty Column06 
         DataField       =   "SEMESTER"
         Caption         =   "SEMESTER"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodb_mata_kuliah 
      Height          =   375
      Left            =   5400
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
      RecordSource    =   $"MATA_KULIAH.frx":0020
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR MATA KULIAH"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   10275
   End
   Begin VB.Image Image1 
      Height          =   3240
      Left            =   11160
      Picture         =   "MATA_KULIAH.frx":00DE
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2880
   End
End
Attribute VB_Name = "MATA_KULIAH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btnedit_Click()
    Frame1.Visible = True
    cmb_semester.SelText = txt_semester.Text
    cmb_dosen.SelText = txt_dosen
End Sub

Private Sub Btnhapus_Click()
    Tanya = MsgBox("Yakin hapus data terpilih?" _
             , vbYesNo + vbQuestion, "Hapus")
            
    If Tanya = vbYes Then
        Adodb_mata_kuliah.Recordset.Delete
    End If
End Sub

Private Sub Btnkeluar_Click()
    Frame1.Visible = False
    Adodb_mata_kuliah.Recordset.CancelBatch
End Sub

Private Sub Btnsimpan_Click()
    'Validasi
    If txt_mata_kuliah.Text = "" Then
        MsgBox "Mata Kuliah tidak boleh kosong!"
        txt_mata_kuliah.SetFocus
    ElseIf cmb_semester.Text = "" Then
        MsgBox "Semester tidak boleh kosong!"
    ElseIf cmb_dosen.Text = "" Then
        MsgBox "Dosen tidak boleh kosong!"
    ElseIf txt_sks.Text = "" Then
        MsgBox "SKS tidak boleh kosong!"
        txt_sks.SetFocus
    ElseIf IsNumeric(txt_sks.Text) = False Then
        MsgBox "SKS harus berupa angka!"
        txt_sks.SetFocus
    'Simpan data
    Else
        Adodb_mata_kuliah.Recordset.Fields("MATA_KULIAH") = txt_mata_kuliah.Text
        Adodb_mata_kuliah.Recordset.Fields("ID_SEMESTER") = txt_id_semester.Text
        Adodb_mata_kuliah.Recordset.Fields("ID_DOSEN") = txt_id_dosen.Text
        Adodb_mata_kuliah.Recordset.Fields("SKS") = txt_sks.Text
        Adodb_mata_kuliah.Recordset.Update
        Adodb_mata_kuliah.Refresh
        Frame1.Visible = False
    End If
End Sub

Private Sub Btntambah_Click()
    Adodb_mata_kuliah.Recordset.AddNew
    Frame1.Visible = True
End Sub

Private Sub cmb_dosen_Click()
    txt_id_dosen.Text = cmb_dosen.ItemData(cmb_dosen.ListIndex)
End Sub

Private Sub cmb_semester_Click()
    txt_id_semester.Text = cmb_semester.List(cmb_semester.ListIndex)
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
    'open the recordset SEMESTER
    With rs
        .Open "SEMESTER", cn, adOpenKeyset, adLockPessimistic, adCmdTable
           
        'loop through the records until reaching the end or last record
        Do While Not .EOF
            cmb_semester.AddItem rs.Fields("SEMESTER")
            cmb_semester.ItemData(cmb_semester.NewIndex) = rs.Fields("ID_SEMESTER")
            rs.MoveNext 'moves next record
        Loop
        
        If Not (.EOF And .BOF) Then
            rs.MoveFirst    'go to the first record if there are existing records
        End If
        
    End With
    
    Set rs2 = New ADODB.Recordset
    'open the recordset DOSEN
    With rs2
        .Open "DOSEN", cn, adOpenKeyset, adLockPessimistic, adCmdTable
           
        'loop through the records until reaching the end or last record
        Do While Not .EOF
            cmb_dosen.AddItem rs2.Fields("NIP") & " - " & rs2.Fields("NAMA")
            cmb_dosen.ItemData(cmb_dosen.NewIndex) = rs2.Fields("ID_DOSEN")
            rs2.MoveNext 'moves next record
        Loop
        
        If Not (.EOF And .BOF) Then
            rs2.MoveFirst    'go to the first record if there are existing records
        End If
        
    End With
    
    Me.MousePointer = vbNormal 'sets the mouse pointer to the normal arrow
    
    
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

Private Sub Form_Unload(Cancel As Integer)
    Adodb_mata_kuliah.Recordset.CancelBatch
End Sub
