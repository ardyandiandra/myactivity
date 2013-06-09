VERSION 5.00
Object = "{19883799-9381-41CA-8342-B4546CA216BF}#1.0#0"; "vistaButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SEMESTER 
   Caption         =   "SEMESTER"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   14325
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      Height          =   3015
      Left            =   7560
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox Txt_Semester 
         DataField       =   "SEMESTER"
         DataSource      =   "Adodb_Semester"
         Height          =   405
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox Txt_Periode 
         DataField       =   "PERIODE"
         DataSource      =   "Adodb_Semester"
         Height          =   405
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
      End
      Begin vistaButton.ButtonVista Btnkeluar 
         Height          =   495
         Left            =   3600
         TabIndex        =   7
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
         Left            =   2160
         TabIndex        =   8
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
      Begin VB.Label Label2 
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
         ForeColor       =   &H0000FF00&
         Height          =   390
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODE"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
   End
   Begin vistaButton.ButtonVista Btnhapus 
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   4560
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
      Left            =   2040
      TabIndex        =   1
      Top             =   4560
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
      Left            =   720
      TabIndex        =   2
      Top             =   4560
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
      Bindings        =   "SEMESTER.frx":0000
      Height          =   2655
      Left            =   720
      TabIndex        =   3
      Top             =   1560
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
      ColumnCount     =   3
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "PERIODE"
         Caption         =   "PERIODE"
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
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodb_Semester 
      Height          =   375
      Left            =   7080
      Top             =   5880
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
      RecordSource    =   "SEMESTER"
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
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA SEMESTER"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1065
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   7605
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   360
      Picture         =   "SEMESTER.frx":001D
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   4080
   End
End
Attribute VB_Name = "SEMESTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btnedit_Click()
    Frame1.Visible = True
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

Private Sub Btnhapus_Click()
    Tanya = MsgBox("Yakin hapus data terpilih?" _
             , vbYesNo + vbQuestion, "Hapus")
            
    If Tanya = vbYes Then
        Adodb_Semester.Recordset.Delete
    End If
End Sub

Private Sub Btnkeluar_Click()
    Frame1.Visible = False
    Adodb_Semester.Recordset.CancelBatch
End Sub

Private Sub Btnsimpan_Click()
    'Validasi
    If Txt_Semester.Text = "" Then
        MsgBox "Semester tidak boleh kosong!"
        Txt_Semester.SetFocus
    ElseIf IsNumeric(Txt_Semester.Text) = False Then
        MsgBox "Semester harus berupa angka!"
        Txt_Semester.SetFocus
    ElseIf Txt_Periode.Text = "" Then
        MsgBox "Periode tidak boleh kosong!"
        Txt_Periode.SetFocus
    'Simpan data
    Else
        Adodb_Semester.Recordset.Fields("SEMESTER") = Txt_Semester.Text
        Adodb_Semester.Recordset.Fields("PERIODE") = Txt_Periode.Text
        Adodb_Semester.Recordset.Update
        Frame1.Visible = False
    End If
End Sub

Private Sub Btntambah_Click()
    Adodb_Semester.Recordset.AddNew
    Frame1.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Adodb_Semester.Recordset.CancelBatch
End Sub
