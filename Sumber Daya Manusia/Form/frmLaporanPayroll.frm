VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLaporanPayroll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Human Resources Department"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   Icon            =   "frmLaporanPayroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9915
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Left            =   3480
      TabIndex        =   18
      Top             =   1440
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   6600
      Width           =   9885
      Begin VB.CommandButton cmdBaru 
         Caption         =   "Baru"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7920
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   7
         Top             =   240
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpBulanPenghitungan 
         Height          =   360
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM, yyyy"
         Format          =   59441155
         UpDown          =   -1  'True
         CurrentDate     =   38231
      End
      Begin VB.Label lblBulanPenghitungan 
         AutoSize        =   -1  'True
         Caption         =   "Bulan Penghitungan"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kriteria Laporan"
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
      TabIndex        =   2
      Top             =   960
      Width           =   9915
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   50
         TabIndex        =   12
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9600
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtIdAkhir 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "0000000001"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtIdAwal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "0000000001"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   6600
         TabIndex        =   19
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Unit Kerja"
         Height          =   195
         Left            =   6600
         TabIndex        =   17
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan"
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ID Pegawai Akhir"
         Height          =   195
         Left            =   11880
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID Pegawai Awal"
         Height          =   195
         Left            =   9720
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   11400
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
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
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   4935
      Left            =   0
      TabIndex        =   20
      Top             =   1920
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8705
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanPayroll.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8040
      Picture         =   "frmLaporanPayroll.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanPayroll.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmLaporanPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilter As String

Private Sub cmdBaru_Click()
    Select Case strMenu
        Case "Daftar Pegawai"
            Label2.Visible = True
            Label3.Visible = True
            Label1.Caption = "ID Pegawai Awal"
            txtIdAkhir.Visible = True
            lblBulanPenghitungan.Visible = False
            dtpBulanPenghitungan.Visible = False
            Text1.Text = ""
            DataCombo2.Text = ""
            DataCombo1.Text = ""
        Case "CV Pegawai"
            Label2.Visible = False
            Label3.Visible = False
            Label1.Caption = "ID Pegawai"
            txtIdAkhir.Visible = False
            lblBulanPenghitungan.Visible = False
            dtpBulanPenghitungan.Visible = False
        Case "Index Pegawai"
            Label2.Visible = False
            Label3.Visible = False
            Label1.Caption = "ID Pegawai"
            Text1.Text = ""
            DataCombo2.Text = ""
            DataCombo1.Text = ""
            txtIdAkhir.Visible = False
            lblBulanPenghitungan.Visible = True
            dtpBulanPenghitungan.Visible = True
            dtpBulanPenghitungan.Value = Format(Now, "dd/mmmm/yyyy")
    End Select
    On Error GoTo errLoad
    strSQL = "SELECT * FROM V_S_DataPegawai"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgData.DataSource = rs
    subSetGrid
    txtIdAwal.Text = "0000000001"
    txtIdAkhir.Text = "0000000001"
    dtpBulanPenghitungan.Value = Format(Now, "dd/mmmm/yyyy")
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub cmdCari_Click()
    Select Case strMenu
        Case "Daftar Pegawai"
            strFilter = " IdPegawai BETWEEN '" & txtIdAwal.Text & "' AND '" & txtIdAkhir.Text & "' "
        Case "CV Pegawai", "Index Pegawai"
            strFilter = " NamaLengkap like '%" & Text1.Text & "%'"
    End Select
    subLoadData strFilter
End Sub

Private Sub cmdCetak_Click()
On Error GoTo hell
    cmdCetak.Enabled = False
    
    Select Case strMenu
        Case "Daftar Pegawai"
            strSQL = "SELECT * FROM V_S_DataPegawai WHERE NamaLengkap like '%" & Text1.Text & "%' and NamaJabatan like '%" & DataCombo1.Text & "%' and NamaRuangan like '%" & DataCombo2.Text & "%'"
        Case "CV Pegawai"
            strSQL = "SELECT * FROM V_S_DataPegawai WHERE IdPegawai = '" & txtIdAwal.Text & "'"
        Case "Index Pegawai"
            If DataCombo2.Text = "" Then
                strSQL = "SELECT * FROM V_HitungIndexPegawai WHERE IdPegawai = '" & dgData.Columns("Id Pegawai").Value & "' AND month(TglHitung) = '" & dtpBulanPenghitungan.Month & "' AND year(TglHitung) = '" & dtpBulanPenghitungan.Year & "'"
            Else
                strSQL = "SELECT * FROM V_HitungIndexPegawaiPerRuangan WHERE NamaRuangan = '" & DataCombo2.Text & "' AND month(TglHitung) = '" & dtpBulanPenghitungan.Month & "' AND year(TglHitung) = '" & dtpBulanPenghitungan.Year & "'"
            End If
    End Select
    
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    Select Case strMenu
        Case "Daftar Pegawai"
            Unload frmCetakDaftarPegawai
            frmCetakDaftarPegawai.Show
        Case "CV Pegawai"
            Unload frmCetakBukuCVPegawai
            frmCetakBukuCVPegawai.Show
        Case "Index Pegawai"
            If DataCombo2.Text = "" Then
                Unload frmCetakIndexPegawai
                frmCetakIndexPegawai.Show
            Else
                Unload frmCetakIndexPegawaiPerRuangan
                frmCetakIndexPegawaiPerRuangan.Show
            End If
    End Select
    cmdCetak.Enabled = True
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpBulanPenghitungan_Change()
    dtpBulanPenghitungan.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Select Case strMenu
        Case "Daftar Pegawai"
            Label2.Visible = True
            Label3.Visible = True
            Label1.Caption = "ID Pegawai Awal"
            txtIdAkhir.Visible = True
            lblBulanPenghitungan.Visible = False
            dtpBulanPenghitungan.Visible = False
            Call SetComboJabatan
            Call SetComboRuangan
        Case "CV Pegawai"
            Label2.Visible = False
            Label3.Visible = False
            Label1.Caption = "ID Pegawai"
            txtIdAkhir.Visible = False
            lblBulanPenghitungan.Visible = False
            dtpBulanPenghitungan.Visible = False
        Case "Index Pegawai"
            Label2.Visible = False
            Label3.Visible = False
            Label1.Caption = "ID Pegawai"
            txtIdAkhir.Visible = False
            lblBulanPenghitungan.Visible = True
            dtpBulanPenghitungan.Visible = True
            dtpBulanPenghitungan.Value = Format(Now, "dd/mmmm/yyyy")
            Call SetComboJabatan
            Call SetComboRuangan
    End Select
    On Error GoTo errLoad
    strSQL = "SELECT * FROM V_S_DataPegawai"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgData.DataSource = rs
    subSetGrid
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub subLoadData(strFilter As String)
    On Error GoTo errLoad
    strSQL = "SELECT * FROM V_S_DataPegawai WHERE " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgData.DataSource = rs
    subSetGrid
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub subSetGrid()
    With dgData
        .Columns(0).Caption = "ID Pegawai"
        .Columns(1).Caption = "Jenis Pegawai"
        .Columns(2).Caption = "Nama Lengkap"
        .Columns(3).Caption = "Sex"
        .Columns(4).Caption = "Tempat Lahir"
        .Columns(5).Caption = "Tgl Lahir"
        .Columns(6).Caption = "Pangkat"
        .Columns(7).Caption = "Golongan"
        .Columns(8).Caption = "Jabatan"
        .Columns(9).Caption = "Pendidikan"
        .Columns(10).Caption = "NIP"
        .Columns(11).Caption = "Aktif"
        .Columns(0).Width = 1500
        .Columns(2).Width = 2500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmKriteriaLaporan = Nothing
End Sub

Private Sub Text1_Change()
    strSQL = "SELECT * FROM V_S_DataPegawai WHERE NamaLengkap like '%" & Text1.Text & "%' and NamaJabatan like '%" & DataCombo1.Text & "%' and NamaRuangan like '%" & DataCombo2.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    subSetGrid
End Sub

'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDown Then dgData.SetFocus
'End Sub
Private Sub datacombo1_Change()
    strSQL = "SELECT * FROM V_S_DataPegawai WHERE NamaLengkap like '%" & Text1.Text & "%' and NamaJabatan like '%" & DataCombo1.Text & "%' and NamaRuangan like '%" & DataCombo2.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    subSetGrid
End Sub

Private Sub datacombo2_Change()
    strSQL = "SELECT * FROM V_S_DataPegawai WHERE NamaLengkap like '%" & Text1.Text & "%' and NamaJabatan like '%" & DataCombo1.Text & "%' and NamaRuangan like '%" & DataCombo2.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    subSetGrid
End Sub

Sub SetComboJabatan()
    Set rs = Nothing
    strSQL = "Select * from Jabatan"
    Call msubDcSource(DataCombo1, rs, strSQL)
End Sub

Sub SetComboRuangan()
    Set rs = Nothing
    strSQL = "Select * from Ruangan"
    Call msubDcSource(DataCombo2, rs, strSQL)
End Sub
