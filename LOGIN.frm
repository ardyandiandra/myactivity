VERSION 5.00
Object = "{19883799-9381-41CA-8342-B4546CA216BF}#1.0#0"; "vistaButton.ocx"
Begin VB.Form LOGIN 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "LOGIN"
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Algerian"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LOGIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vistaButton.ButtonVista ButtonVista3 
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonShape     =   1
      ButtonStyle     =   7
      BackColor       =   49152
      BackColorPressed=   15715986
      BackColorHover  =   65280
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      ForeColorHover  =   0
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
   Begin vistaButton.ButtonVista ButtonVista2 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonShape     =   1
      ButtonStyle     =   7
      BackColor       =   49152
      BackColorPressed=   15715986
      BackColorHover  =   65280
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      ForeColorHover  =   0
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
   Begin vistaButton.ButtonVista ButtonVista1 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonShape     =   1
      ButtonStyle     =   7
      BackColor       =   49152
      BackColorPressed=   15715986
      BackColorHover  =   65280
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      ForeColorHover  =   0
      Caption         =   "MASUK"
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
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "COLLEGE"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MY"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2520
      TabIndex        =   8
      Top             =   480
      Width           =   750
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ACTIVITY"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   -480
      Picture         =   "LOGIN.frx":9632
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim MoveScreen As Boolean, color As Long, flag As Byte
Dim MousX, MousY, CurrX, CurrY As Integer

Function Login_User(user As String, pass As String)
    Dim oRS As ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim oConn As ADODB.Connection
       
    Set oConn = New ADODB.Connection
   
    oConn.Provider = "Microsoft.Jet.OLEDB.4.0"
    oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                          "Data Source=" & App.Path & "\activity.mdb"
    oConn.Open
       
    sSQL = "SELECT USERNAME, PASSWORD FROM PROFILE WHERE USERNAME = '" & user & "'"

    Set oRS = New ADODB.Recordset
   
    'Open our recordset
    oRS.Open sSQL, oConn
   
    If oRS.EOF Then
        MsgBox ("User tidak ditemukan!")
    Else
        If oRS.Fields.Item("PASSWORD") <> pass Then
            MsgBox ("Password salah!")
        Else
            MDIMenu.Show
            Unload Me
        End If
    End If

    oRS.Close
    oConn.Close
End Function

Private Sub ButtonVista1_Click()
   Call Login_User(Text1.Text, Text2.Text)
End Sub

Private Sub ButtonVista2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub ButtonVista3_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorRtn
color = RGB(0, 0, 255): flag = 0
flag = flag Or LWA_COLORKEY: LOGIN.Show
SetTranslucent LOGIN.hwnd, color, 0, flag
Exit Sub
ErrorRtn: MsgBox Err.Description & " Source : " & Err.Source
End Sub

Private Sub image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MoveScreen = True: MousX = x: MousY = Y
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If MoveScreen Then
CurrX = Me.Left - MousX + x: CurrY = Me.Top - MousY + Y
Me.Move CurrX, CurrY
End If
End Sub

Private Sub image1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
MoveScreen = False
End Sub


Private Sub BATAL_Click()
Unload Me
End Sub


Private Sub MASUK_Click()
If Text1.Text = "NAMA" Then
MDIMenu.Show
ElseIf Text2.Text = "NAMA" Then
MDIMenu.Show
Unload Me
Else
MsgBox "masukan user name dan password"
End If
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub RESET_Click()
Text1.Text = ""
Text2.Text = ""
End Sub


