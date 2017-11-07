VERSION 5.00
Begin VB.Form frmFilterUsulan 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3630
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7005
   Begin VB.TextBox txtIdPegawai 
      Height          =   375
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   6495
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdUsulan 
         Caption         =   "&Usulan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Riwayat Usulan Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.OptionButton option5 
         Caption         =   "Pengangkatan Pegawai TPHL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton option6 
         Caption         =   "Pengangkatan Pegawai Negeri Sipil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
      Begin VB.OptionButton option7 
         Caption         =   "Pemberhentian Pegawai TPHL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   8
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "TAPERUM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Pensiun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Kenaikan Gaji"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Kenaikan Pangkat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmFilterUsulan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdUsulan_Click()
    On Error GoTo hell
    If frmDataPegawaiNew.txtidpegawai.Text = "" Then Exit Sub
    mstrIdPegawai = frmDataPegawaiNew.txtidpegawai.Text

    '' usulan pengangkatan pegawai TPHL
    If option5.Value = True Then
        frmRiwayatUsulanPengangkatanTPHL.Show
        With frmRiwayatUsulanPengangkatanTPHL
            .txtidpegawai.Text = mstrIdPegawai
            .txtnamapegawai.Text = frmDataPegawaiNew.txtNama.Text
            .txtTempatlahir.Text = frmDataPegawaiNew.txtTptLhr.Text
            .meTglLahir.Text = frmDataPegawaiNew.meTglLahir.Text
            .txtPendidikan.Text = frmDataPegawaiNew.dcJurusan.Text
        End With
    End If

    '' usulan pengangkatan PNS
    If option6.Value = True Then
        strSQL = "SELECT dbo.DetailKategoryPegawai_M.KdKategoryPegawai" & _
        " FROM dbo.DataCurrentPegawai INNER JOIN " & _
        "dbo.DetailKategoryPegawai_M ON dbo.DataCurrentPegawai.KdDetailKategoryPegawai = dbo.DetailKategoryPegawai_M.KdDetailKategoryPegawai where dbo.DataCurrentPegawai.IdPegawai='" & mstrIdPegawai & "' "
        Call msubRecFO(rsb, strSQL)
        If rsb.EOF = True Then MsgBox "Lengkapi kategori pegawai ", vbCritical, "Validasi": Exit Sub
        If rsb(0).Value = "2" Then
            frmRiwayatUsulanPengangkatanPNS.Show
            With frmRiwayatUsulanPengangkatanPNS
                .txtidpegawai.Text = mstrIdPegawai
                .txtnamapegawai.Text = frmDataPegawaiNew.txtNama.Text
                .txtTempatlahir.Text = frmDataPegawaiNew.txtTptLhr.Text
                .meTglLahir.Text = frmDataPegawaiNew.meTglLahir.Text
                .txtPendidikan.Text = frmDataPegawaiNew.dcJurusan.Text
            End With
        Else
            MsgBox "Pegawai yang dipilih bukan PNS, atau kategory pegawai kosong", vbCritical, "Validasi"
            Exit Sub
        End If
    End If

    '' usulan kenaikan gaji
    If Option2.Value = True Then
        frmRiwayatUsulanKenaikanGaji.Show
        With frmRiwayatUsulanKenaikanGaji
            .txtidpegawai.Text = mstrIdPegawai
            .txtnamapegawai.Text = frmDataPegawaiNew.txtNama.Text
            .txtTempatlahir.Text = frmDataPegawaiNew.txtTptLhr.Text
            .meTglLahir.Text = frmDataPegawaiNew.meTglLahir.Text
            .txtPendidikan.Text = frmDataPegawaiNew.dcJurusan.Text
        End With
    End If

    '' usulan kenaikan pangkat
    If Option1.Value = True Then
        frmRiwayatUsulanKenaikanPangkat.Show
        With frmRiwayatUsulanKenaikanPangkat
            .txtidpegawai.Text = mstrIdPegawai
            .txtnamapegawai.Text = frmDataPegawaiNew.txtNama.Text
            .txtTempatlahir.Text = frmDataPegawaiNew.txtTptLhr.Text
            .meTglLahir.Text = frmDataPegawaiNew.meTglLahir.Text
            .txtPendidikan.Text = frmDataPegawaiNew.dcJurusan.Text
        End With
    End If

    '' usulan pensiun
    If Option3.Value = True Then
        frmRiwayatUsulanPesiun.Show
        With frmRiwayatUsulanPesiun
            .txtidpegawai.Text = mstrIdPegawai
            .txtnamapegawai.Text = frmDataPegawaiNew.txtNama.Text
            .txtTempatlahir.Text = frmDataPegawaiNew.txtTptLhr.Text
            .meTglLahir.Text = frmDataPegawaiNew.meTglLahir.Text
            .txtPendidikan.Text = frmDataPegawaiNew.dcJurusan.Text
        End With
    End If

    ''riwayat usulan taperum
    If Option4.Value = True Then
        frmRiwayatUsulanTaperum.Show
        With frmRiwayatUsulanTaperum
            .txtidpegawai.Text = mstrIdPegawai
            .txtnamapegawai.Text = frmDataPegawaiNew.txtNama.Text
            .txtTempatlahir.Text = frmDataPegawaiNew.txtTptLhr.Text
            .meTglLahir.Text = frmDataPegawaiNew.meTglLahir.Text
            .txtPendidikan.Text = frmDataPegawaiNew.dcJurusan.Text
        End With
    End If

    ' usulan TPHL berhenti
    If option7.Value = True Then
        strSQL = "SELECT dbo.DetailKategoryPegawai_M.KdKategoryPegawai" & _
        " FROM dbo.DataCurrentPegawai INNER JOIN " & _
        "dbo.DetailKategoryPegawai_M ON dbo.DataCurrentPegawai.KdDetailKategoryPegawai = dbo.DetailKategoryPegawai_M.KdDetailKategoryPegawai where dbo.DataCurrentPegawai.IdPegawai='" & mstrIdPegawai & "' "
        Call msubRecFO(rsb, strSQL)
        If rsb.EOF = True Then MsgBox "Lengkapi kategori pegawai ", vbCritical, "Validasi": Exit Sub
        If rsb(0).Value = "3" Then
            frmRiwayatUsulanTPHLBerhenti.Show
            With frmRiwayatUsulanTPHLBerhenti
                .txtidpegawai.Text = mstrIdPegawai
                .txtnamapegawai.Text = frmDataPegawaiNew.txtNama.Text
                .txtTempatlahir.Text = frmDataPegawaiNew.txtTptLhr.Text
                .meTglLahir.Text = frmDataPegawaiNew.meTglLahir.Text
                .txtPendidikan.Text = frmDataPegawaiNew.dcJurusan.Text
            End With
        Else
            MsgBox "Pegawai yang dipilih bukan TPHL, atau kategory pegawai kosong", vbCritical, "Validasi"
            Exit Sub
        End If
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
End Sub
