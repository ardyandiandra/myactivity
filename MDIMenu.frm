VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "MY ACTIVITY"
   ClientHeight    =   10650
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14490
   Icon            =   "MDIMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIMenu.frx":9632
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   84
      ImageHeight     =   84
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":264B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":2FAF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":39137
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":42779
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":4BDBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":553FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":5EA3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":68081
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":716C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":7AD05
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14490
      _ExtentX        =   25559
      _ExtentY        =   2910
      ButtonWidth     =   2408
      ButtonHeight    =   2752
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DOSEN"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PROFILE"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SEMESTER"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MATA KULIAH"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "JADWAL KULIAH"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "KEGIATAN"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MODUL"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ALARM"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "HISTORY"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "KELUAR"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menupengaturan 
      Caption         =   "PENGATURAN"
      Index           =   1
      Begin VB.Menu menudosen 
         Caption         =   "DOSEN"
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu menuprofile 
         Caption         =   "PROFILE"
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu menusemester 
         Caption         =   "SEMESTER"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu menuexit 
         Caption         =   "EXIT"
         Index           =   6
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menukuliah 
      Caption         =   "KULIAH"
      Index           =   2
      Begin VB.Menu menukegiatan 
         Caption         =   "KEGIATAN "
         Index           =   1
         Shortcut        =   ^K
      End
      Begin VB.Menu menumodul 
         Caption         =   "MODUL"
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu menumatkul 
         Caption         =   "MATA KULIAH"
         Index           =   3
      End
      Begin VB.Menu menujadkul 
         Caption         =   "JADWAL KULIAH"
         Index           =   4
      End
   End
   Begin VB.Menu menualarm 
      Caption         =   "ALARM"
      Index           =   3
   End
   Begin VB.Menu menulogout 
      Caption         =   "LOG OUT"
      Index           =   4
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub alarm_Click(Index As Integer)
ALARM.Show
End Sub


Private Sub keluar_Click(Index As Integer)
End
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo err:
err:
End Sub

Private Sub menualarm_Click(Index As Integer)
ALARM.Show
Unload JADWAL_KULIAH
Unload KEGIATAN
Unload mata_kuliah
Unload MODUL
Unload PROFILE
Unload SEMESTER
Unload DOSEN
End Sub

Private Sub menudosen_Click(Index As Integer)
DOSEN.Show
Unload JADWAL_KULIAH
Unload KEGIATAN
Unload mata_kuliah
Unload MODUL
Unload PROFILE
Unload SEMESTER
Unload ALARM
End Sub

Private Sub menuexit_Click(Index As Integer)
End
End Sub

Private Sub menujadkul_Click(Index As Integer)
JADWAL_KULIAH.Show
Unload DOSEN
Unload KEGIATAN
Unload mata_kuliah
Unload MODUL
Unload PROFILE
Unload SEMESTER
Unload ALARM
End Sub

Private Sub menukegiatan_Click(Index As Integer)
KEGIATAN.Show
Unload JADWAL_KULIAH
Unload DOSEN
Unload mata_kuliah
Unload MODUL
Unload PROFILE
Unload SEMESTER
Unload ALARM
End Sub


Private Sub menumasuk_Click(Index As Integer)
LOGIN.Show
End Sub


Private Sub menulogout_Click(Index As Integer)
Unload Me
LOGIN.Show
End Sub

Private Sub menumatkul_Click(Index As Integer)
mata_kuliah.Show
Unload JADWAL_KULIAH
Unload KEGIATAN
Unload DOSEN
Unload MODUL
Unload PROFILE
Unload SEMESTER
Unload ALARM
End Sub

Private Sub menumodul_Click(Index As Integer)
MODUL.Show
Unload JADWAL_KULIAH
Unload KEGIATAN
Unload mata_kuliah
Unload DOSEN
Unload PROFILE
Unload SEMESTER
Unload ALARM
End Sub

Private Sub menuprofile_Click(Index As Integer)
PROFILE.Show
Unload JADWAL_KULIAH
Unload KEGIATAN
Unload mata_kuliah
Unload MODUL
Unload DOSEN
Unload SEMESTER
Unload ALARM
End Sub

Private Sub menureset_Click(Index As Integer)

End Sub

Private Sub menusemester_Click(Index As Integer)
SEMESTER.Show
Unload JADWAL_KULIAH
Unload KEGIATAN
Unload mata_kuliah
Unload MODUL
Unload PROFILE
Unload DOSEN
Unload ALARM
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    DOSEN.Show
    Unload JADWAL_KULIAH
    Unload KEGIATAN
    Unload mata_kuliah
    Unload MODUL
    Unload PROFILE
    Unload SEMESTER
    Unload ALARM
Case 2
    PROFILE.Show
    Unload JADWAL_KULIAH
    Unload KEGIATAN
    Unload mata_kuliah
    Unload MODUL
    Unload DOSEN
    Unload SEMESTER
    Unload ALARM
Case 3
    SEMESTER.Show
    Unload JADWAL_KULIAH
    Unload KEGIATAN
    Unload mata_kuliah
    Unload MODUL
    Unload PROFILE
    Unload DOSEN
    Unload ALARM
Case 4
    mata_kuliah.Show
    Unload JADWAL_KULIAH
    Unload KEGIATAN
    Unload DOSEN
    Unload MODUL
    Unload PROFILE
    Unload SEMESTER
    Unload ALARM
Case 5
    JADWAL_KULIAH.Show
    Unload DOSEN
    Unload KEGIATAN
    Unload mata_kuliah
    Unload MODUL
    Unload PROFILE
    Unload SEMESTER
    Unload ALARM
Case 6
    KEGIATAN.Show
    Unload JADWAL_KULIAH
    Unload DOSEN
    Unload mata_kuliah
    Unload MODUL
    Unload PROFILE
    Unload SEMESTER
    Unload ALARM
Case 7
    MODUL.Show
    Unload JADWAL_KULIAH
    Unload KEGIATAN
    Unload mata_kuliah
    Unload DOSEN
    Unload PROFILE
    Unload SEMESTER
    Unload ALARM
Case 8
    ALARM.Show
    Unload JADWAL_KULIAH
    Unload KEGIATAN
    Unload mata_kuliah
    Unload DOSEN
    Unload PROFILE
    Unload SEMESTER
    Unload MODUL
Case 9
Case 10
    End
End Select
End Sub
