VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLaporanRiwayatPangkat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Laporan Riwayat Pangkat"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanRiwayatPangkat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5565
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   2025
   End
   Begin VB.CommandButton cmdtutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2280
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   5475
      Begin VB.Frame Frame3 
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   5055
         Begin VB.OptionButton Option2 
            Caption         =   "Tahun"
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Bulan"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPickerAwal 
            Height          =   375
            Left            =   2640
            TabIndex        =   3
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMMM yyyy"
            Format          =   478347267
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4440
      Picture         =   "frmLaporanRiwayatPangkat.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1155
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanRiwayatPangkat.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanRiwayatPangkat.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmLaporanRiwayatPangkat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    On Error GoTo errLoad
    Dim TempThn As String
    Dim pesan As VbMsgBoxResult
    TempThn = Format(DTPickerAwal.Value, "yyyy")

    If Option1.Value = True Then

        strSQL = "select * from V_LaporanRiwayatPangkat " & _
        " Where month(TglSK) = '" & Format(DTPickerAwal, "MM") & "' and year(TglSK) = '" & TempThn & "' "
        Call msubRecFO(dbRst, strSQL)

    ElseIf Option2.Value = True Then
        strSQL = "select * from V_LaporanRiwayatPangkat " & _
        " Where Year(TglSK) ='" & TempThn & "'"
        Call msubRecFO(dbRst, strSQL)

    End If
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    frm_cetak_LaporanPangkat.Show
    Exit Sub
errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub DTPickerAwal_Change()
    DTPickerAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Option2.Value = True
    DTPickerAwal.Value = Format(Now, "MMMM yyyy")
    DTPickerAwal.CustomFormat = ("yyyy")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Option1_Click()
    DTPickerAwal.CustomFormat = ("MMMM yyyy")
End Sub

Private Sub Option2_Click()
    DTPickerAwal.CustomFormat = ("yyyy")
End Sub
