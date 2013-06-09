VERSION 5.00
Object = "{19883799-9381-41CA-8342-B4546CA216BF}#1.0#0"; "vistaButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DOSEN 
   Caption         =   "DOSEN"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   15645
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      Height          =   4935
      Left            =   8640
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox Txt_Email 
         DataField       =   "EMAIL"
         DataSource      =   "Adodb_Dosen"
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox Txt_Keterangan 
         DataField       =   "KETERANGAN"
         DataSource      =   "Adodb_Dosen"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   3360
         Width           =   3975
      End
      Begin VB.TextBox Txt_Telepon 
         DataField       =   "TELP"
         DataSource      =   "Adodb_Dosen"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox Txt_Alamat 
         DataField       =   "ALAMAT"
         DataSource      =   "Adodb_Dosen"
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox Txt_Nama 
         DataField       =   "NAMA"
         DataSource      =   "Adodb_Dosen"
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox Txt_Nip 
         DataField       =   "NIP"
         DataSource      =   "Adodb_Dosen"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   3975
      End
      Begin vistaButton.ButtonVista Btnkeluar 
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   4080
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
         Left            =   2160
         TabIndex        =   8
         Top             =   4080
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "nama"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "alamat"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "email"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "keterangan"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   360
         TabIndex        =   10
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIP"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   360
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "DOSEN.frx":0000
      Height          =   3735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Algerian"
         Size            =   12
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
      BeginProperty Column01 
         DataField       =   "NIP"
         Caption         =   "NIP"
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
         DataField       =   "NAMA"
         Caption         =   "NAMA"
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
         DataField       =   "TELP"
         Caption         =   "TELP"
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
         DataField       =   "EMAIL"
         Caption         =   "EMAIL"
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
         DataField       =   "ALAMAT"
         Caption         =   "ALAMAT"
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
         DataField       =   "KETERANGAN"
         Caption         =   "KETERANGAN"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
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
   Begin MSAdodcLib.Adodc Adodb_Dosen 
      Height          =   375
      Left            =   6600
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "DOSEN"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin vistaButton.ButtonVista Btnhapus 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   4440
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
      Left            =   1920
      TabIndex        =   2
      Top             =   4440
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
      TabIndex        =   3
      Top             =   4440
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
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   120
      Picture         =   "DOSEN.frx":001A
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2880
   End
End
Attribute VB_Name = "DOSEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub Btnedit_Click()
     Frame1.Visible = True
End Sub

Private Sub Btnhapus_Click()
    Tanya = MsgBox("Yakin hapus data terpilih?" _
             , vbYesNo + vbQuestion, "Hapus")
            
    If Tanya = vbYes Then
        Adodb_Dosen.Recordset.Delete
    End If
End Sub
Private Sub Btnkeluar_Click()
    Frame1.Visible = False
    Adodb_Dosen.Recordset.CancelBatch
End Sub
Private Sub Btnsimpan_Click()
    'Validasi
    If Txt_Nip.Text = "" Then
        MsgBox "NIP tidak boleh kosong!"
        Txt_Nip.SetFocus
    ElseIf Txt_Nama = "" Then
        MsgBox "Nama tidak boleh kosong!"
        Txt_Nama.SetFocus
    'Simpan Data
    Else
        Adodb_Dosen.Recordset.Fields("NIP") = Txt_Nip.Text
        Adodb_Dosen.Recordset.Fields("NAMA") = Txt_Nama.Text
        Adodb_Dosen.Recordset.Fields("TELP") = Txt_Telepon.Text
        Adodb_Dosen.Recordset.Fields("ALAMAT") = Txt_Alamat.Text
        Adodb_Dosen.Recordset.Fields("EMAIL") = Txt_Email.Text
        Adodb_Dosen.Recordset.Fields("KETERANGAN") = Txt_Keterangan.Text
        Adodb_Dosen.Recordset.Update
        Frame1.Visible = False
    End If
End Sub

Private Sub Btntambah_Click()
    Adodb_Dosen.Recordset.AddNew
    Frame1.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Adodb_Dosen.Recordset.CancelBatch
End Sub

