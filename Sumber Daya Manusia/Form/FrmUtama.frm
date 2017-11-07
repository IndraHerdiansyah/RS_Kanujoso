VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmNotifikasi2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notifications"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "FrmUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4695
   Begin VB.CommandButton Command2 
      Caption         =   "List Notifikasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingatkan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "C:\Users\Public\Music\Sample Music\Maid with the Flaxen Hair.mp3"
      MaxFileSize     =   8000
   End
   Begin VB.Timer tmrFlash 
      Left            =   2160
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   2640
   End
   Begin VB.PictureBox XPFrame1 
      Height          =   1335
      Left            =   2520
      ScaleHeight     =   1275
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      MouseIcon       =   "FrmUtama.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmNotifikasi2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Dim Naik As Boolean

Private Sub Command1_Click()
    Select Case Combo1.Text
        Case "1 Jam lagi"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~1"
        Case "5 Jam lagi"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~5"
        Case "10 Jam lagi"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~10"
        Case "Lupakan"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~0"
    End Select
    Unload Me
End Sub

Private Sub Command2_Click()
    frmNotifikasi.Show
    MDIUtama.Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Naik = False
    Timer1.Enabled = True
End Sub

Private Sub Label1_Click()
    MDIUtama.Timer1.Enabled = False
    Unload Me
'    frmDaftarPemesananBarangdariRuangan.Show
End Sub

Private Sub tmrFlash_Timer()

  Dim ret As Long

  ret = FlashWindow(hWnd, CLng(True))
End Sub
Private Sub Form_Load()
    Top = ((GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY)
    Left = (GetSystemMetrics(16) * Screen.TwipsPerPixelX) - Width
    Naik = True
Dim Reg, s
s = Replace(App.path & "\" & App.EXEName & ".exe", "\\", "\")
Set Reg = CreateObject("WScript.Shell")
Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, s
  tmrFlash.Interval = 1000 / 5
Dim i As Integer
  tmrFlash.Enabled = True
'  For i = 1 To Timer2.Interval
'  Beep
'    Next i

Combo1.AddItem "1 Jam lagi"
Combo1.AddItem "5 Jam lagi"
Combo1.AddItem "10 Jam lagi"
Combo1.AddItem "Lupakan"
Combo1.Text = ""
End Sub

Private Sub Timer1_Timer()
    Const s = 80 'kecepatan gerak / slide
    Dim v As Single
    v = (GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY
    
    If Naik = True Then
        If Top - s <= v - Height Then
            Top = Top - (Top - (v - Height))
            Timer1.Enabled = False
        Else
            Top = Top - s
        End If
        
    Else
        Top = Top + s
        If Top >= v Then Unload Me
    End If
End Sub

