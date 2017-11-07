VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCekAbsensi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Monitoring Absensi"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "frmCekAbsensi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11580
   Begin VB.CommandButton cmdtutup 
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
      Height          =   375
      Left            =   8640
      TabIndex        =   17
      Top             =   7200
      Width           =   2895
   End
   Begin VB.TextBox txtParameter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   12
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   11535
      Begin VB.TextBox txtjk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtNamaPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txttempattugas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   8880
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtjabatan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   5640
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JK"
         Height          =   210
         Index           =   3
         Left            =   5160
         TabIndex        =   16
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   210
         Index           =   8
         Left            =   5640
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan"
         Height          =   210
         Index           =   4
         Left            =   8880
         TabIndex        =   8
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         Height          =   210
         Index           =   2
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID Pegawai"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1260
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
   Begin MSDataGridLib.DataGrid dgAbsensi 
      Height          =   4335
      Left            =   0
      TabIndex        =   10
      Top             =   2760
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpTgl 
      Height          =   405
      Left            =   9960
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   109510659
      UpDown          =   -1  'True
      CurrentDate     =   37760
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal Monitoring  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   14
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Cari Pegawai :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9720
      Picture         =   "frmCekAbsensi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmCekAbsensi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14175
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   0
      Picture         =   "frmCekAbsensi.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmCekAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgAbsensi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgAbsensi
    WheelHook.WheelHook dgAbsensi
End Sub

Private Sub dgAbsensi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgAbsensi
        If .Columns("IdPegawai").Value = "" Then
            txtidpegawai.Text = ""
        Else
            txtidpegawai.Text = .Columns("IdPegawai").Value
        End If
        If .Columns("NamaLengkap").Value = "" Then
            txtnamapegawai.Text = ""
        Else
            txtnamapegawai.Text = .Columns("NamaLengkap").Value
        End If
        'If .Columns("JenisKelamin").Value = "" Then
        If txtidpegawai.Text = "" Then
            txtJK.Text = ""
        Else
            Set rs = Nothing
            strSQL = "SELECT DISTINCT dbo.DataPegawai.JenisKelamin " & _
                     "FROM         dbo.v_AbsensiPegawai RIGHT OUTER JOIN " & _
                     "dbo.DataPegawai ON dbo.v_AbsensiPegawai.IdPegawai = dbo.DataPegawai.IdPegawai " & _
                     "WHERE     (dbo.v_AbsensiPegawai.IdPegawai = '" & txtidpegawai.Text & "')"
            Call msubRecFO(rs, strSQL)
            txtJK.Text = rs.Fields(0).Value
            'txtjk.Text = .Columns("JenisKelamin").Value
        End If
        If .Columns("NamaRuangan").Value = "" Then
            txtTempatTugas.Text = ""
        Else
            txtTempatTugas.Text = .Columns("NamaRuangan").Value
        End If
        If .Columns("NamaJabatan").Value = "" Then
            txtJabatan.Text = ""
        Else
            txtJabatan.Text = .Columns("NamaJabatan").Value
        End If
    End With

End Sub

Private Sub dtpTgl_Change()

    mdCekTgl = dtpTgl.Value
    Call loadgrid
    txtParameter.Text = ""

End Sub

Private Sub Form_Load()

    centerForm Me, MDIUtama
    Call PlayFlashMovie(Me)
    dtpTgl.Value = Date
    mdCekTgl = dtpTgl.Value
    Call loadgrid

End Sub

Sub loadgrid()
        On Error GoTo hell
    Set rs = Nothing
    '//yayang.agus 2014-08-07
    'strSQL = "SELECT * from v_AbsensiPegawai where NamaLengkap like '%" & txtParameter.Text & "%' and tanggal='" & Format(dtpTgl.Value, "yyyy-mm-dd") & "'"
'    strSQL = "SELECT distinct  [IdPegawai],[NamaLengkap],[EmployeId],[NIP],[KdTypePegawai],[TypePegawai],[KdJabatan] " & _
             ",[NamaJabatan],[KdRuanganKerja],[NamaRuangan],[kdShift],[NamaShift],[JamMasuk],[JamPulang],[JamIstirahatAwal] " & _
             ",[JamIstirahatAkhir],[Tanggal],left(cast([WaktuMasuk] as time),8) as WaktuMasuk,left(cast([WaktuKeluar] as time),8) as WaktuKeluar,[JmlJamMasuk] " & _
             "From [bethesda].[dbo].[v_AbsensiPegawai]" & _
             "where NamaLengkap like '%" & txtParameter.Text & "%' and tanggal='" & Format(dtpTgl.Value, "yyyy-mm-dd") & "'"
    'strSQL = "SELECT distinct  [IdPegawai],[NamaLengkap],[EmployeId],[NIP],[KdTypePegawai],[TypePegawai],[KdJabatan] ,[NamaJabatan],[KdRuanganKerja],[NamaRuangan],[kdShift],[NamaShift],[JamMasuk],[JamPulang],[JamIstirahatAwal] ,[JamIstirahatAkhir],[Tanggal],convert(varchar(8),[WaktuMasuk],108) as WaktuMasuk,convert(varchar(8),[WaktuKeluar],108) as WaktuKeluar,dbo.hitungjamdarimenit([JmlJamMasuk]) as JmlJamMasuk From [bethesda].[dbo].[v_AbsensiPegawai] " & _
           "where NamaLengkap like '%" & txtParameter.Text & "%' and tanggal='" & Format(dtpTgl.Value, "yyyy-mm-dd") & "'"
'    strSQL = "SELECT  [idPegawai],[NamaLengkap],[NIP],[KdTypePegawai],[TypePegawai],[KdJabatan],[NamaJabatan] , " & _
             "[KdRuangan],[NamaRuangan],[JadwalKerja],[KdShift],[NamaShift],[JamMasuk],[JamPulang],[JamIstirahatAwal] , " & _
             "[JamIstirahatAkhir],convert(varchar(8),[TglMasuk],108) as WaktuMasuk,convert(varchar(8),[TglPulang],108) as WaktuKeluar,dbo.selisih_tanggal(tglmasuk,tglpulang) as JmlJamMasuk " & _
             "From v_AbsensiPegawai2_1 " & _
             "where NamaLengkap like '%" & txtParameter.Text & "%' and jadwalkerja='" & Format(dtpTgl.Value, "yyyy-mm-dd") & "' and TglMasuk is not null and TglPulang is not null " & _
             "order by idpegawai"
    strSQL = "SELECT  [idPegawai],[NamaLengkap],[NIP],[KdTypePegawai],[TypePegawai],[KdJabatan],[NamaJabatan] , [KdRuangan],[NamaRuangan],[JadwalKerja],[KdShift],[NamaShift]," & _
             "[JamMasuk],[JamPulang],[JamIstirahatAwal] , [JamIstirahatAkhir],dbo.selisih_tanggal([JamIstirahatAwal],[JamIstirahatAkhir]) as JmlJamIstirahat , " & _
             "convert(varchar(8),[TglMasuk],108) as WaktuMasuk,convert(varchar(8),[TglPulang],108) as WaktuKeluar, " & _
             "dbo.selisih_tanggal(tglmasuk,tglpulang)as JmlJamMasuk ,dbo.JmlJamKerja(tglmasuk,tglpulang,[JamIstirahatAwal],[JamIstirahatAkhir]) as JmlJamKerja " & _
            "From v_AbsensiPegawai2_1 " & _
             "where NamaLengkap like '%" & txtParameter.Text & "%' and jadwalkerja='" & Format(dtpTgl.Value, "yyyy-mm-dd") & "' and TglMasuk is not null and TglPulang is not null " & _
             "order by idpegawai"
    '//
    'strSQL = "SELECT * from v_AbsensiPegawai where NamaLengkap like '%" & txtParameter.Text & "%'"

    Call msubRecFO(rs, strSQL)
    Set dgAbsensi.DataSource = rs
    Call setdgAbsensi
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtParameter_Change()
    Call loadgrid
End Sub
Sub setdgAbsensi()

    With dgAbsensi

        .Columns("idPegawai").Visible = 1500
        .Columns("NamaLengkap").Visible = 1500
'        .Columns("EmployeId").Visible = 1500
        .Columns("NIP").Visible = 1500
        .Columns("KdTypePegawai").Visible = False
        .Columns("TypePegawai").Width = 2000
        .Columns("KdJabatan").Width = False
        .Columns("NamaJabatan").Visible = 1500
        .Columns("KdRuangan").Visible = False
        .Columns("NamaRuangan").Width = 1500
'        .Columns("JadwalKerja").Width = 1500
        .Columns("kdShift").Visible = False
        .Columns("NamaShift").Visible = 1500
        .Columns("JamMasuk").Visible = 1500
        .Columns("JamPulang").Visible = 1500
        .Columns("JamIstirahatAwal").Visible = 1500
        .Columns("JamIstirahatAkhir").Visible = 1500
        .Columns("jadwalkerja").Width = 1500
'        .Columns("WaktuMasuk").Visible = 1500
'        .Columns("WaktuKeluar").Visible = 2000
'        .Columns("JmlJamMasuk").Width = 2000



'        .Columns("PINAbsensi").Visible = False
'        .Columns("NIP").Visible = False
'        .Columns("NamaLengkap").Visible = 2500
'        .Columns("KdJabatan").Visible = False
'        .Columns("NamaJabatan").Width = 2000
'        .Columns("KdJenisJabatan").Width = False
'        .Columns("KdTypePegawai").Visible = False
'        .Columns("TypePegawai").Visible = False
'        .Columns("KdRuanganKerja").Width = 1500
'        .Columns("NamaRuangan").Width = 1500
'        .Columns("JadwalKerja").Visible = 1500
'        .Columns("KdShift").Visible = False
'        .Columns("NamaShift").Visible = 1500
'        .Columns("TglMasuk").Visible = 1500
''        .Columns("JamPulang").Visible = 1500
''        .Columns("JamIstirahatAwal").Visible = 1500
''        .Columns("JamIstirahatAkhir").Width = 1500
''        .Columns("Tanggal").Visible = 1500
'        .Columns("JamAbsenMasuk").Visible = 2000
'        .Columns("JamAbsenPulang").Width = 2000
'        .Columns("ThnTglAbsen").Width = 2000
'        .Columns("BlnTglAbsen").Width = 2000
'        .Columns("MntTglAbsen").Visible = 2000
'        .Columns("DtkTglAbsen").Visible = 2000
'        .Columns("NIP").Visible = False
'        .Columns("Tanggal").Visible = False
'        .Columns("TglJadwalKerja").Visible = False
'        .Columns("NamaLengkap").Width = 2000
'        .Columns("TglAbsen").Width = 2000
'        .Columns("WaktuMasuk").Visible = False
'        .Columns("WaktuKeluar").Visible = False
'        .Columns("AbsenMasuk").Width = 1500
'        .Columns("AbsenKeluar").Width = 1500
'        .Columns("KdShift").Visible = False
'        .Columns("Terlambat").Visible = False
'        .Columns("PulangCepat").Visible = False
'        .Columns("JmlJamKerja").Visible = False
'        .Columns("TypePegawai").Visible = False
'        .Columns("NoRiwayat").Visible = False
'        .Columns("StatusAbsensi").Width = 1500
'        .Columns("JenisKelamin").Visible = False
'        .Columns("Ruangan").Width = 2000
'        .Columns("NamaJabatan").Visible = False
'        .Columns("KdJabatan").Visible = False
    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
