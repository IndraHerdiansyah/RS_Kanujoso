VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarSchedule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medifirst2000 - Gap Kompetensi Pegawai"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12810
   Icon            =   "frmDaftarSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid fgLooping 
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2355
      _Version        =   393216
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   12735
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11430
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDetailKirim 
         Caption         =   "Edit Gap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblxx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid dgGapMapping 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   12735
      Begin VB.TextBox txtInstitusiCari 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4200
         Width           =   4755
      End
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
         Left            =   6840
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            Format          =   72351747
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            Format          =   72351747
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   9
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.Label lblInstansi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   4245
         Width           =   1185
      End
      Begin VB.Label lblKodeSupplier 
         BackStyle       =   0  'Transparent
         Caption         =   "List Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarSchedule.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarSchedule.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10920
      Picture         =   "frmDaftarSchedule.frx":4CE9
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "frmDaftarSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long

Private Sub cmdCari_Click()
Call LoadGrid
End Sub

Private Sub cmdDetailKirim_Click()
    If dgGapMapping.ApproxCount = 0 Then Exit Sub
    
    frmEditGap.Show
    frmEditGap.txtIDPegawai.Text = dgGapMapping.Columns("IdPegawai")
    
    'frmMappingGapKompetensiValidasi.TxtNoGap.Text = dgGapMapping.Columns(0)
    'frmMappingGapKompetensiValidasi.TxtNoGap_KeyPress (13)
    'frmMappingGapKompetensiValidasi.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
'___________________________________________________________________________________________________________________________________________________________________________________________________
'    strSQL = "select * from V_GapKompetensiDetail"
'    frm_cetak_GapKompetensi.Show
'___________________________________________________________________________________________________________________________________________________________________________________________________
    'SELECT No, NamaPegawai, Jabatan, StandarPendidikan, PendidikanReal, GapPendidikan, Skill, KebutuhanDiklat,
    'RealTraining1, RealTraining2, GapDiklat
    'From tempMonitoringGapKompetensi
    
    Dim no, nama, Nip, Pangkat1, Pangkat2, Jabatan1, Jabatan2, MasaKerja1, MasaKerja2, LatihanJabatan1, LatihanJabatan2, LatihanJabatan3, Pendidikan1, Pendidikan2, Pendidikan3, Usia, MutasiKerja As String
    Dim ii As Integer
    Dim BRS As Integer
    Dim BRS_AWAL As Integer
    Dim iii As Integer
    
    strSQLX = "delete from tempMonitoringGapKompetensi"
    Call msubRecFO(rsx, strSQLX)
    strSQL = "select * from V_GapKompetensi  where namalengkap like '%" & txtInstitusiCari.Text & "%'" '" 'WHERE     IdPegawai in ('L000000085')"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        fgLooping.Rows = 1
        fgLooping.Cols = 20
        BRS = 0
        BRS_AWAL = 1
        For ii = 0 To rs.RecordCount - 1
            fgLooping.Rows = fgLooping.Rows + 1
'            BRS = fgLooping.Rows - 1
            no = no + 1
'            Nama = IIf(IsNull(rs!Nama), "", rs!Nama) 'rs!Nama
'            Nip = IIf(IsNull(rs!Nip), "", rs!Nip) 'rs!Nip
            
            'BRS = BRS + 1
            BRS_AWAL = fgLooping.Rows - 1
            
            'For i = 1 To fgLooping.Rows - 1
                fgLooping.TextMatrix(BRS_AWAL, 1) = no
                fgLooping.TextMatrix(BRS_AWAL, 2) = IIf(IsNull(rs!NamaLengkap), "", rs!NamaLengkap)
                fgLooping.TextMatrix(BRS_AWAL, 3) = IIf(IsNull(rs!NamaJabatan), "", rs!NamaJabatan)
            'Next
            
            strsqlxx = "select * from V_GapKompetensiDetail2  WHERE     kdJabatan = '" & rs(2) & "'"
            Call msubRecFO(rsxx, strsqlxx)
            If rsxx.RecordCount <> 0 Then
                BRS = BRS_AWAL - 1
                For i = 1 To rsxx.RecordCount
                    BRS = BRS + 1
                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                    fgLooping.TextMatrix(BRS, 4) = IIf(IsNull(rsxx(2)), "", rsxx(2)) 'rsxx(0)
                    rsxx.MoveNext
                Next
            End If
            
            'Pendidikan Real
            fgLooping.TextMatrix(BRS_AWAL, 5) = IIf(IsNull(rs!Skill), "", rs!KualifikasiJurusan)
            'GapPEndidikan
            fgLooping.TextMatrix(BRS_AWAL, 6) = "-"
            'skill
            fgLooping.TextMatrix(BRS_AWAL, 7) = IIf(IsNull(rs!Skill), "", rs!Skill)
            
            DoEvents
            'KEbutuhan Diklat
            strsqlxx = "select * from V_GapKompetensiDetail  WHERE     kdJabatan = '" & rs(2) & "'"
            Call msubRecFO(rsxx, strsqlxx)
            If rsxx.RecordCount <> 0 Then
                BRS = BRS_AWAL - 1
                For i = 0 To rsxx.RecordCount - 1
                    BRS = BRS + 1
                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                    fgLooping.TextMatrix(BRS, 8) = IIf(IsNull(rsxx(2)), "", rsxx(2)) 'rsxx(0)
                    strSQLc1 = "select * from V_GapKompetensiDetail3  WHERE     kdJabatan = '" & rs(2) & "' and namadiklat='" & fgLooping.TextMatrix(BRS, 8) & "' and idpegawai='" & rs!idpegawai & "'"
                    Call msubRecFO(rsc1, strSQLc1)
                    If rsc1.RecordCount <> 0 Then
'                        BRS = BRS_AWAL - 1
                        For iii = 0 To rsc1.RecordCount - 1
'                            BRS = BRS + 1
'                            If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
                            fgLooping.TextMatrix(BRS, 9) = "v"
                            fgLooping.TextMatrix(BRS, 10) = ""
                            rsc1.MoveNext
                        Next
                    End If
                    rsxx.MoveNext
                Next
            End If
'
'            strsqlxx = "select * from V_GapKompetensiDetail3  WHERE     kdJabatan = '" & rs(2) & "' and namadiklat='" & fgLooping.TextMatrix(BRS, 8) & "' and idpegawai='" & rs!idpegawai & "'"
'            Call msubRecFO(rsxx, strsqlxx)
'            If rsxx.RecordCount <> 0 Then
'                BRS = BRS_AWAL - 1
'                For i = 0 To rsxx.RecordCount - 1
'                    BRS = BRS + 1
'                    If fgLooping.Rows - 1 < BRS Then fgLooping.Rows = BRS + 1
'                    fgLooping.TextMatrix(BRS, 9) = "v"
'                    fgLooping.TextMatrix(BRS, 10) = ""
'                    rsxx.MoveNext
'                Next
'            End If

            fgLooping.TextMatrix(BRS_AWAL, 11) = ""
            
            lblxx.Caption = ii & "/" & rs.RecordCount - 1
            rs.MoveNext
        Next
    End If
        
    Dim persen As Double
    Dim jmlAll As Integer
    Dim jmlOke As Integer
    Dim brsNama As String
    Dim namaa As String
    Dim brsAkhir As Integer
    For i = 1 To fgLooping.Rows - 1
        If fgLooping.TextMatrix(i, 2) <> "" Then
            jmlAll = 0
            jmlOke = 0
            brsNama = i
            brsAkhir = i
            namaa = fgLooping.TextMatrix(i, 2)
        End If
        jmlAll = jmlAll + 1
        If fgLooping.TextMatrix(i, 9) = "v" Then
            jmlOke = jmlOke + 1
            fgLooping.TextMatrix(i, 10) = ""
        Else
             fgLooping.TextMatrix(i, 10) = "x"
        End If
        If jmlAll > 0 And jmlOke > 0 Then
            'fgLooping.TextMatrix(brsNama, 11) = ((jmlOke / jmlAll) * 100) & " %"
            fgLooping.TextMatrix(brsNama, 11) = ((100 / jmlAll) * (jmlAll - jmlOke)) & " %"
        Else
            fgLooping.TextMatrix(brsNama, 11) = "100 %"
        End If
        fgLooping.TextMatrix(i, 2) = namaa
        
    Next
    
    For i = 1 To fgLooping.Rows - 1
        If fgLooping.TextMatrix(i, 1) = "" Then
            fgLooping.TextMatrix(i, 7) = fgLooping.TextMatrix(i - 1, 7)
            'fgLooping.TextMatrix(i - 1, 7) = ""
        End If
    Next
    
    With fgLooping
        For i = 1 To .Rows - 1
            strSQLX = "insert into tempMonitoringGapKompetensi values (" & _
                    "'" & .TextMatrix(i, 1) & "','" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "','" & .TextMatrix(i, 4) & "','" & .TextMatrix(i, 5) & "','" & .TextMatrix(i, 6) & "','" & .TextMatrix(i, 7) & "'," & _
                    "'" & .TextMatrix(i, 8) & "','" & .TextMatrix(i, 9) & "','" & .TextMatrix(i, 10) & "','" & .TextMatrix(i, 11) & "'" & _
                    ")"
            Call msubRecFO(rsx, strSQLX)
        Next
    End With
    
    vLaporan = ""
    If MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then vLaporan = "Print"
    
    strSQL = "select * from tempMonitoringGapKompetensi"
    frm_cetak_GapKompetensi.Show
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With
    Call LoadGrid
End Sub

Sub LoadGrid()
    'strSQL = "SELECT     ScheduleGapKompetensi.NoScheduleGapKompetensi, ScheduleGapKompetensi.TglScheduleGapKompetensi, DataPegawai.NamaLengkap FROM         ScheduleGapKompetensi INNER JOIN DataPegawai ON ScheduleGapKompetensi.IdUser = DataPegawai.IdPegawai where ScheduleGapKompetensi.TglScheduleGapKompetensi between '" & Format(dtpAwal.Value, "yyyy-MM-dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy-MM-dd 23:59:59") & "'"
    strSQL = "select IdPegawai, NamaLengkap, JenisKelamin as 'JK', TglLahir, TglMasuk, NamaJabatan from V_ListPegawai where namalengkap like '%" & txtInstitusiCari.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dgGapMapping.DataSource = rs
    
    dgGapMapping.Columns(0).Width = 1200
    dgGapMapping.Columns(1).Width = 3500
    dgGapMapping.Columns(2).Width = 700
    dgGapMapping.Columns(3).Width = 1500
    dgGapMapping.Columns(4).Width = 1500
    dgGapMapping.Columns(5).Width = 3500
End Sub

Private Sub txtInstitusiCari_Change()
    Call LoadGrid
End Sub
