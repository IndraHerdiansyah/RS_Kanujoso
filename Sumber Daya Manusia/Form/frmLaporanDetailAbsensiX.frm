VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLaporanDetailAbsensiX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Cetak Abensi Pegawai"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11940
   Icon            =   "frmLaporanDetailAbsensiX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11940
   Begin VB.OptionButton optTahun 
      Caption         =   "per Tahun"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton optBulan 
      Caption         =   "per Bulan"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton optHari 
      Caption         =   "per Hari"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Frame fraTahun 
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
      Left            =   5520
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   5175
      Begin MSComCtl2.DTPicker dtpBulanTahun 
         Height          =   330
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy"
         Format          =   60686339
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "Cari"
      Height          =   375
      Left            =   10800
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame fraHari 
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
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   5175
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   60686339
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   60686339
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label1 
         Caption         =   "s\d"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   6840
      Width           =   11775
      Begin VB.CommandButton cmdcetakdetail 
         Caption         =   "Cetak &Detail Absensi"
         Height          =   495
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdCetakAbsensi 
         Caption         =   "Cetak &Absensi"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   9960
         TabIndex        =   2
         Top             =   240
         Width           =   1575
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
   Begin MSDataGridLib.DataGrid dgData 
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcRuangan 
      Height          =   390
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   12000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Ruangan"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLaporanDetailAbsensiX.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10080
      Picture         =   "frmLaporanDetailAbsensiX.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanDetailAbsensiX.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmLaporanDetailAbsensiX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCari_Click()
    Dim strFilter As String
    
    Screen.MousePointer = vbHourglass
    
    strFilter = ""
    strsql = ""
    
    If optHari.Value = True Then
        
        strFilter = "ThnTglMasuk BETWEEN '" & Format(dtpAwal, "yyyy") & "' AND '" & Format(dtpAkhir, "yyyy") & "' AND " & _
                    "BlnTglMasuk BETWEEN '" & Format(dtpAwal, "mm") & "' AND '" & Format(dtpAkhir, "mm") & "' AND " & _
                    "Tgl_TglMasuk BETWEEN '" & Format(dtpAwal, "dd") & "' AND '" & Format(dtpAkhir, "dd") & "') as A"

        strsql = "Select  B.NamaPegawai, B.NIP, B.Jabatan, B.SubRuangKerja, TglMasuk=A.TglMasuk, TglPulang=A.TglPulang, TotalAbsensi=A.Total " & _
                  "From " & _
                  "(select distinct NamaLengkap, NIP, Jabatan, TglMasuk, TglIstirahatAwal, TglIstirahatAkhir, TglPulang, Total, idPegawai, ThnTglMasuk, BlnTglMasuk, Tgl_TglMasuk " & _
                          "From v_CetakAbsensiX " & _
                  "WHERE " & strFilter & " " & _
                  "right outer join" & _
                  "(SELECT     dbo.DataPegawai.IdPegawai, ISNULL(dbo.DataCurrentPegawai.KdRuanganKerja, '---') " & _
                  "AS KdRuanganKerja, ISNULL(dbo.DataPegawai.NamaLengkap, '---') AS NamaPegawai, " & _
                  "ISNULL(dbo.DataCurrentPegawai.NIP, '---') AS NIP, ISNULL(dbo.Jabatan.NamaJabatan, '---') AS Jabatan, " & _
                  "dbo.SubRuangKerja.SubRuangKerja " & _
                  "FROM   dbo.Jabatan RIGHT OUTER JOIN dbo.DataCurrentPegawai LEFT OUTER JOIN dbo.RuangKerja INNER JOIN " & _
                  "dbo.SubRuangKerja ON dbo.RuangKerja.KdRuangKerja = dbo.SubRuangKerja.KdRuangKerja ON " & _
                  "dbo.DataCurrentPegawai.KdRuanganKerja = dbo.SubRuangKerja.KdSubRuangKerja RIGHT OUTER JOIN dbo.DataPegawai ON " & _
                  "dbo.DataCurrentPegawai.IdPegawai = dbo.DataPegawai.IdPegawai ON dbo.Jabatan.KdJabatan = dbo.DataCurrentPegawai.KdJabatan " & _
                  "Where DataCurrentPegawai.KdStatus='01') as B on A.idPegawai=B.idPegawai WHERE B.SubRuangKerja LIKE '%" & dcRuangan.Text & "%'"
          
    End If
    
    If optBulan.Value = True Then
    
        strFilter = "ThnTglMasuk BETWEEN '" & Format(dtpAwal, "yyyy") & "' AND '" & Format(dtpAkhir, "yyyy") & "' AND " & _
                    "BlnTglMasuk BETWEEN '" & Format(dtpAwal, "mm") & "' AND '" & Format(dtpAkhir, "mm") & "') as A"
        
        strsql = "Select  B.NamaPegawai, B.NIP, B.Jabatan, B.SubRuangKerja, A.Total as TotalAbsensi, A.BlnTglMasuk, A.ThnTglMasuk " & _
                    "From " & _
                    "(select distinct NamaLengkap, NIP, Jabatan, TglMasuk, TglIstirahatAwal, TglIstirahatAkhir, TglPulang, Total, idPegawai, ThnTglMasuk, BlnTglMasuk, Tgl_TglMasuk " & _
                            "From v_CetakAbsensiX " & _
                    "WHERE " & strFilter & " " & _
                    "right outer join" & _
                    "(SELECT     dbo.DataPegawai.IdPegawai, ISNULL(dbo.DataCurrentPegawai.KdRuanganKerja, '---') " & _
                    "AS KdRuanganKerja, ISNULL(dbo.DataPegawai.NamaLengkap, '---') AS NamaPegawai, " & _
                    "ISNULL(dbo.DataCurrentPegawai.NIP, '---') AS NIP, ISNULL(dbo.Jabatan.NamaJabatan, '---') AS Jabatan, " & _
                    "dbo.SubRuangKerja.SubRuangKerja " & _
                    "FROM   dbo.Jabatan RIGHT OUTER JOIN dbo.DataCurrentPegawai LEFT OUTER JOIN dbo.RuangKerja INNER JOIN " & _
                    "dbo.SubRuangKerja ON dbo.RuangKerja.KdRuangKerja = dbo.SubRuangKerja.KdRuangKerja ON " & _
                    "dbo.DataCurrentPegawai.KdRuanganKerja = dbo.SubRuangKerja.KdSubRuangKerja RIGHT OUTER JOIN dbo.DataPegawai ON " & _
                    "dbo.DataCurrentPegawai.IdPegawai = dbo.DataPegawai.IdPegawai ON dbo.Jabatan.KdJabatan = dbo.DataCurrentPegawai.KdJabatan " & _
                    "Where DataCurrentPegawai.KdStatus='01') as B on A.idPegawai=B.idPegawai WHERE B.SubRuangKerja LIKE '%" & dcRuangan.Text & "%' " & _
                    "GROUP BY B.NamaPegawai, B.NIP, B.Jabatan, B.SubRuangKerja, A.Total, A.BlnTglMasuk, A.ThnTglMasuk"

    End If
    
    If optTahun.Value = True Then
        
        strFilter = "ThnTglMasuk = '" & Format(dtpBulanTahun, "yyyy") & "') as A"
                
        strsql = "Select  B.NamaPegawai, B.NIP, B.Jabatan, B.SubRuangKerja, A.ThnTglMasuk " & _
                    "From " & _
                    "(select distinct NamaLengkap, NIP, Jabatan, TglMasuk, TglIstirahatAwal, TglIstirahatAkhir, TglPulang, idPegawai, ThnTglMasuk, BlnTglMasuk, Tgl_TglMasuk " & _
                            "From v_CetakAbsensiX " & _
                    "WHERE " & strFilter & " " & _
                    "right outer join" & _
                    "(SELECT     dbo.DataPegawai.IdPegawai, ISNULL(dbo.DataCurrentPegawai.KdRuanganKerja, '---') " & _
                    "AS KdRuanganKerja, ISNULL(dbo.DataPegawai.NamaLengkap, '---') AS NamaPegawai, " & _
                    "ISNULL(dbo.DataCurrentPegawai.NIP, '---') AS NIP, ISNULL(dbo.Jabatan.NamaJabatan, '---') AS Jabatan, " & _
                    "dbo.SubRuangKerja.SubRuangKerja " & _
                    "FROM   dbo.Jabatan RIGHT OUTER JOIN dbo.DataCurrentPegawai LEFT OUTER JOIN dbo.RuangKerja INNER JOIN " & _
                    "dbo.SubRuangKerja ON dbo.RuangKerja.KdRuangKerja = dbo.SubRuangKerja.KdRuangKerja ON " & _
                    "dbo.DataCurrentPegawai.KdRuanganKerja = dbo.SubRuangKerja.KdSubRuangKerja RIGHT OUTER JOIN dbo.DataPegawai ON " & _
                    "dbo.DataCurrentPegawai.IdPegawai = dbo.DataPegawai.IdPegawai ON dbo.Jabatan.KdJabatan = dbo.DataCurrentPegawai.KdJabatan " & _
                    "Where DataCurrentPegawai.KdStatus='01') as B on A.idPegawai=B.idPegawai WHERE B.SubRuangKerja LIKE '%" & dcRuangan.Text & "%' " & _
                    "GROUP BY B.NamaPegawai, B.NIP, B.Jabatan, B.SubRuangKerja, A.ThnTglMasuk"
    End If
    
'sum(isnull(A.StatusAbsensi,0)) as TotalAbsensi,
'StatusAbsensi,
    If strsql <> "" Then
        Set rs = Nothing
        rs.Open strsql, dbConn, adOpenForwardOnly, adLockReadOnly
        Set dgData.DataSource = rs
        
        If optHari.Value = True Then
            Call subSetGridHari
        End If
        
        If optBulan.Value = True Then
            Call subSetGridBulan
        End If
        
        If optTahun.Value = True Then
            Call subSetGridTahun
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub subSetGridTahun()
    With dgData
        .Columns(0).Caption = "Nama Lengkap"
        .Columns(1).Caption = "NIP"
        .Columns(2).Caption = "Jabatan"
        .Columns(3).Caption = "Nama Ruangan"
        .Columns(4).Caption = "Total Absensi"
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 2000
        .Columns(4).Width = 2000
    End With
End Sub

Private Sub subSetGridBulan()
    With dgData
        .Columns(0).Caption = "Nama Lengkap"
        .Columns(1).Caption = "NIP"
        .Columns(2).Caption = "Jabatan"
        .Columns(3).Caption = "Nama Ruangan"
        .Columns(4).Caption = "Total Absensi"
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 2000
        .Columns(4).Width = 2000
    End With
End Sub

Private Sub subSetGridHari()
    With dgData
        .Columns(0).Caption = "Nama Lengkap"
        .Columns(1).Caption = "NIP"
        .Columns(2).Caption = "Jabatan"
        .Columns(3).Caption = "Nama Ruangan"
        .Columns(4).Caption = "Tgl Masuk"
        .Columns(5).Caption = "Tgl Pulang"
        .Columns(6).Caption = "Status"
        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 2000
        .Columns(4).Width = 2000
        .Columns(5).Width = 2000
        .Columns(6).Width = 2000
    End With
End Sub

Private Sub cmdCetakAbsensi_Click()
    Select Case dtpAwal.Month
        Case 1, 3, 5, 7, 8, 10, 12
            subTanggalTerakhir = 31
        Case 4, 6, 9, 11
            subTanggalTerakhir = 30
        Case 2
            subTanggalTerakhir = 28
    End Select
    
    pubStrRuangan = UCase(dcRuangan.Text)
    pubStrPeriode = ""
    
    If optHari.Value = True Then
        glDateCetak = Format(dtpAwal, "dd/MM/yyyy")
        pubStrPeriode = UCase("PERIODE TANGGAL " & Format(dtpAwal, "dd MMMM yyyy") & " S/D " & Format(dtpAkhir, "dd MMMM yyyy"))
        frmCetakAbsensiPegawaiX.Show
        Exit Sub
    End If
    
    If optBulan.Value = True Then
        pubStrPeriode = UCase("PERIODE BULAN " & Format(dtpAwal, "MMMM yyyy") & " S/D " & Format(dtpAkhir, "MMMM yyyy"))
        frmCetakAbsensiBulan.Show
        Exit Sub
    End If
    
    If optTahun.Value = True Then
        pubStrPeriode = UCase("PERIODE TAHUN " & Format(dtpBulanTahun, "yyyy"))
        frmCetakAbsensiTahun.Show
        Exit Sub
    End If
End Sub

Private Sub cmdcetakdetail_Click()
    If optHari.Value = True Then
        strCetak = "Hari"
    ElseIf optBulan.Value = True Then
        strCetak = "Bulan"
    Else
        strCetak = ""
    End If
    
    mdTglAwal = dtpAwal.Value 'TglAwal
    mdTglAkhir = dtpAkhir.Value 'TglAkhir

    strIsiGroup = dcRuangan.Text
    strCetak2 = "CetakDetailAbsensi"
    frmCetakAbsensiPegawai.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpBulanTahun_Change()
    dtpBulanTahun.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call cmdCari_Click
    
    Set rs = Nothing
    strsql = "Select * from SubRuangKerja where StatusEnabled = 1 order by SubRuangKerja"
    Call msubDcSource(dcRuangan, rs, strsql)
    
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    dtpBulanTahun.Value = Now
    
End Sub

Private Sub optHari_Click()
    cmdcetakdetail.Visible = True
    fraTahun.Visible = False
    
    dtpAwal.CustomFormat = "dd MMMM yyyy"
    dtpAkhir.CustomFormat = "dd MMMM yyyy"
    
    Call cmdCari_Click
End Sub

Private Sub optBulan_Click()
    cmdcetakdetail.Visible = True
    fraTahun.Visible = False
    
    dtpAwal.CustomFormat = "MMMM yyyy"
    dtpAkhir.CustomFormat = "MMMM yyyy"
    
    Call cmdCari_Click
End Sub

Private Sub optTahun_Click()
    cmdcetakdetail.Visible = False
    fraTahun.Visible = True
    Call cmdCari_Click
End Sub
