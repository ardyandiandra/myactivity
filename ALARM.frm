VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{19883799-9381-41CA-8342-B4546CA216BF}#1.0#0"; "vistaButton.ocx"
Begin VB.Form ALARM 
   BackColor       =   &H80000004&
   Caption         =   "ALARM"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleMode       =   0  'User
   ScaleWidth      =   14535
   WindowState     =   2  'Maximized
   Begin vistaButton.ButtonVista Btnexit 
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2400
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
      Caption         =   "EXIT"
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
   Begin vistaButton.ButtonVista Btnreset 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2400
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
      Caption         =   "RESET"
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
   Begin vistaButton.ButtonVista Btnstart 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2400
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
      Caption         =   "START"
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   3495
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   6165
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2012
      Month           =   1
      Day             =   15
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox TextWAKTU 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox TextPESAN 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESAN"
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
      Left            =   600
      TabIndex        =   7
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WAKTU"
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
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   120
      Picture         =   "ALARM.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   2880
   End
End
Attribute VB_Name = "ALARM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PESAN As String
Private WAKTU As Date
Public MAXTIME As Long
Public STARTTIME As Date
Public CURRENTTIME As Long

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

Private Sub Btnexit_Click()
Unload Me
End Sub

Private Sub Btnreset_Click()
TextPESAN.Text = ""
ALARM.Show
End Sub

Private Sub Btnstart_Click()
If TextPESAN.Text = "" Then
    MsgBox "Tulis Pesan Alarm Anda"
    TextPESAN.SetFocus
ElseIf TextWAKTU.Text = "" Then
    MsgBox "ISI WAKTU ANDA"
    TextWAKTU.SetFocus
Else
    PESAN = TextPESAN.Text
    WAKTU = TextWAKTU.Text
    TextPESAN.Enabled = False
    TextWAKTU.Enabled = False
    Timer1.Enabled = True
    ALARM.Hide
End If
End Sub

Private Sub Calendar1_DblClick()
With Calendar1
.SetFocus
.Refresh
.Today
End With
End Sub

Private Sub Form_Load()
TextWAKTU.Text = Format(Time, "hh:mm")
End Sub


Private Sub Timer1_Timer()
    STARTTIME = Format(Time, "hh:mm")
    If STARTTIME = WAKTU Then
        MsgBox "" & PESAN, vbInformation, "PERHATIAN"
        Timer1.Enabled = False
    End If
End Sub
