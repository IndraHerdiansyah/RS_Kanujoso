VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbsensiPegawai_OffLine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Absensi Pegawai"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13815
   ControlBox      =   0   'False
   Icon            =   "frmAbsensiPegawai_OffLine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13815
   Begin VB.TextBox txtParameter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   28
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Timer tmrAbsensi 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   3960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Absensi"
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
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   13575
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   4920
         TabIndex        =   21
         Top             =   120
         Width           =   8535
         Begin VB.CommandButton cmdTutup 
            Caption         =   "Tutup"
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
            Left            =   6600
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdKoneksiDB 
            Caption         =   "Ambil Data Dari Alat"
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
            Left            =   4680
            TabIndex        =   22
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lbServer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "<Port>"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblDatabase 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "<Port>"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2280
            TabIndex        =   26
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Port:"
            Height          =   195
            Left            =   2280
            TabIndex        =   24
            Top             =   120
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "IP Addres:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.TextBox txttgl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtFingerPrint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Absen"
         Height          =   210
         Index           =   7
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finger Scan Client"
         Height          =   210
         Index           =   1
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   13575
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtJabatan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   6600
         MaxLength       =   50
         TabIndex        =   5
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtTempatTugas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   10080
         MaxLength       =   50
         TabIndex        =   4
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         Height          =   210
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Index           =   3
         Left            =   5400
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan"
         Height          =   210
         Index           =   4
         Left            =   10080
         TabIndex        =   8
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         Height          =   210
         Index           =   8
         Left            =   6600
         TabIndex        =   7
         Top             =   240
         Width           =   1215
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
      Height          =   5415
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9551
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   8115
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19182
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "15/12/2015"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "9:53"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
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
      TabIndex        =   29
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12000
      Picture         =   "frmAbsensiPegawai_OffLine.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmAbsensiPegawai_OffLine.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmAbsensiPegawai_OffLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fso As New Scripting.FileSystemObject
Private strServerFP As String, strDatabaseFP As String
Private strIdPegawaiAbsen As String, strNoRiwayat As String

Public Sub subLoadDBFP()
    strServerFP = funcGetFromINI("Database Finger Print", "Nama Server", "", strFileSettingDBFP)
    strDatabaseFP = funcGetFromINI("Database Finger Print", "Nama Database", "", strFileSettingDBFP)
    lbServer.Caption = strServerFP
    lblDatabase.Caption = strDatabaseFP
End Sub

Private Sub subLoadAbsensiHariIni()
    'strSQL = "select distinct IdPegawai,NamaLengkap,EmployeId,NIP,Tanggal,WaktuMasuk,WaktuKeluar,JmlJamMasuk,TypePegawai,NamaJabatan from v_AbsensiPegawai where GUId='" & frmKoneksiDb.guidId & "' order by EmployeId DESC"
    
    
    'strSQL = "select distinct IdPegawai,NamaLengkap,EmployeId,NIP,Tanggal,convert(varchar(8),WaktuMasuk,108) as WaktuMasuk ,convert(varchar(8),WaktuKeluar,108) WaktuKeluar ,dbo.hitungjamdarimenit(JmlJamMasuk) as JmlJamMasuk,TypePegawai,NamaJabatan from v_AbsensiPegawai where GUId='" & frmKoneksiDb.guidId & "' order by EmployeId DESC"
    
'    strSQL = "select distinct IdPegawai,NamaLengkap,NIP,JadwalKerja ,convert(varchar(8),TglMasuk ,108) as WaktuMasuk ,convert(varchar(8),TglPulang ,108) WaktuKeluar ,dbo.selisih_tanggal(tglmasuk,tglpulang) as JmlJamMasuk , TypePegawai,NamaJabatan " & _
             "from v_AbsensiPegawai2_1 " & _
             "where TglMasuk between '" & Format(TglAwal, "yyyy-mm-dd 00:00:00") & "' and '" & Format(TglAkhir, "yyyy-mm-dd 23:59:59") & "' " & _
             "and NamaLengkap like '%" & txtParameter.Text & "%'" & _
             "order by idPegawai  DESC" '//yayang.agus 2014-08-22
    strSQL = "select distinct IdPegawai,NamaLengkap,NIP,namajabatan,JadwalKerja ,convert(varchar(8),TglMasuk ,108) as WaktuMasuk ,convert(varchar(8),TglPulang ,108) WaktuKeluar ," & _
             "dbo.selisih_tanggal([JamIstirahatAwal],[JamIstirahatAkhir]) as JmlJamIstirahat , " & _
             "dbo.selisih_tanggal(tglmasuk,tglpulang) as JmlJamMasuk , dbo.JmlJamKerja(tglmasuk,tglpulang,[JamIstirahatAwal],[JamIstirahatAkhir]) as JmlJamKerja " & _
             "from v_AbsensiPegawai2_1 " & _
             "where TglMasuk between '" & Format(TglAwal, "yyyy-mm-dd 00:00:00") & "' and '" & Format(TglAkhir, "yyyy-mm-dd 23:59:59") & "' " & _
             "and NamaLengkap like '%" & txtParameter.Text & "%'" & _
             "order by idPegawai  DESC" '//yayang.agus 2014-08-22
    
    Call msubRecFO(rs, strSQL)
    With dgAbsensi
        Set .DataSource = rs
        If rs.RecordCount = 0 Then Exit Sub
        .Columns("IdPegawai").Visible = 1500
        .Columns("NamaLengkap").Visible = 2500
'        .Columns("EmployeId").Visible = False
        .Columns("NIP").Visible = False
        '.Columns("KdTypePegawai").Visible = False
'        .Columns("TypePegawai").Width = 1300
        '.Columns("KdJabatan").Width = False
'        .Columns("NamaJabatan").Visible = 1200
        '.Columns("KdRuanganKerja").Visible = False
        '.Columns("NamaRuangan").Width = 1500
        '.Columns("JadwalKerja").Width = False
        '.Columns("kdShift").Visible = False
        
        .Columns("jadwalkerja").Width = 1500
        .Columns("WaktuMasuk").Visible = 2500
        .Columns("WaktuKeluar").Visible = 2500
        .Columns("JmlJamMasuk").Width = 1500
'        .Columns("ThnTglAbsen").Width = 2000
'        .Columns("BlnTglAbsen").Width = 2000
'        .Columns("MntTglAbsen").Visible = 2000
'        .Columns("DtkTglAbsen").Visible = 2000
''        .Columns("TglMulai").Visible = False
    End With
'    strSQL = "select distinct * from v_AbsensiPegawai order by TglAbsen DESC"
'    Call msubRecFO(rs, strSQL)
'    With dgAbsensi
'        Set .DataSource = rs
'        .Columns("IdPegawai").Visible = False
'        .Columns("PINAbsensi").Visible = False
'        .Columns("NIP").Visible = False
'        .Columns("NamaLengkap").Visible = False
'        .Columns("KdJabatan").Visible = False
'        .Columns("NamaJabatan").Width = 2000
'        .Columns("KdJenisJabatan").Width = 2000
'        .Columns("KdTypePegawai").Visible = False
'        .Columns("TypePegawai").Visible = False
'        .Columns("KdRuanganKerja").Width = 1500
'        .Columns("NamaRuangan").Width = 1500
'        .Columns("JadwalKerja").Visible = False
'        .Columns("KdShift").Visible = False
'        .Columns("NamaShift").Visible = False
'        .Columns("JamMasuk").Visible = False
'        .Columns("JamPulang").Visible = False
'        .Columns("JamIstirahatAwal").Visible = False
'        .Columns("JamIstirahatAkhir").Width = 1500
'        .Columns("KdStatusAbsensi").Visible = False
'        .Columns("Tanggal").Visible = False
'        .Columns("TglAbsen").Width = 2000
'        .Columns("NamaJabatan").Width = 2000
'        .Columns("KdJabatan").Visible = False
'        .Columns("KdRuangan").Visible = False
'        .Columns("TglMulai").Visible = False
'    End With
End Sub

Private Sub subAmbilDataAbsenFP()
    Dim dbConnFP As New ADODB.connection
    Dim myConSTR As String
    Dim rsFP As ADODB.recordset
    Dim strRecord As String

    Dim strID As String
    Dim strIdFP As String
    Dim strNoFP As String
    Dim strTglWktAbsen As String

    Dim intTempID As Long
    Dim strTempID As String

    dbConnFP.CursorLocation = adUseServer
    myConSTR = "Provider=SQLOLEDB10;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & strDatabaseFP & ";Data Source=" & strServerFP
    dbConnFP.Open myConSTR

    If Err Then
        StatusBar1.Panels(1).Text = "Koneksi ke database Finger Print error!"
    Else
        StatusBar1.Panels(1).Text = "Terkoneksi ke database Finger Print.."
        strSQL = "select max(ID) as ID from TempIdRecordDariFP"
        Call msubRecFO(rs, strSQL)

        If IsNull(rs("ID")) Then
            intTempID = 0
        Else
            strTempID = Trim(rs("ID"))
            intTempID = strTempID
        End If

        strRecord = intTempID + 1

        strSQL = "select ID, Mach_Name, Per_Code, Date_Time" & _
        " from TA_Record_Info" & _
        " where ID='" & strRecord & "' "
        Set rsFP = New ADODB.recordset
        rsFP.Open strSQL, dbConnFP, adOpenForwardOnly, adLockReadOnly

        If Not rsFP.EOF Then
            With rsFP.Fields
                strID = .Item("ID").Value
                strNoFP = .Item("Mach_Name").Value
                strIdFP = .Item("Per_Code").Value
                strTglWktAbsen = Format(.Item("Date_Time").Value, "yyyy/MM/dd hh:mm:ss")
            End With

            strSQL = "select IdPegawai " & _
            " from PINAbsensiPegawai" & _
            " where PINAbsensi='" & strIdFP & "'"
            Call msubRecFO(rs, strSQL)

            If Not rs.EOF Then

                Dim strWaktuAbsen As String
                Dim strTempTglAbsen As String
                Dim strTempTglAbsenX As String
                Dim strJamMasuk As String, strJamPulang As String
                Dim strJamMasukToleransi As String, strJamPulangToleransi As String
                Dim IdShift As String

                strIdPegawaiAbsen = rs.Fields.Item("IdPegawai").Value
                strWaktuAbsen = Format(strTglWktAbsen, "dd/mm/yyyy HH:mm")
                strTempTglAbsen = CInt(Format(strWaktuAbsen, "dd") - 1)

                strSQL = "select TglAbsen from AbsensiPegawai where IdPegawai='" & strIdPegawaiAbsen & "' and day(TglAbsen)='" & Format(strWaktuAbsen, "dd") & "' and month(TglAbsen)='" & Format(strWaktuAbsen, "MM") & "' and year(TglAbsen)='" & Format(strWaktuAbsen, "yyyy") & "' "
                Call msubRecFO(rs, strSQL)

                If rs.EOF = False Then

                    If CInt(Format(strWaktuAbsen, "HH")) > CInt(Format(rs.Fields(0), "HH")) + 2 Then
                        If funcSimpanAbsensiPulang(strTglWktAbsen) Then Call subSimpanIDrecordFP(strID, strIdFP)
                    Else
                        Call subSimpanIDrecordFP(strID, strIdFP)
                    End If

                Else

                    strsqlx = "select * from AbsensiPegawai where IdPegawai='" & strIdPegawaiAbsen & "' and day(TglAbsen)='" & strTempTglAbsen & "' and month(TglAbsen)='" & Format(strWaktuAbsen, "MM") & "' and year(TglAbsen)='" & Format(strWaktuAbsen, "yyyy") & "'"
                    Call msubRecFO(rsx, strsqlx)
                    If rsx.EOF = False Then

                        If Format(rsx.Fields("TglAbsen").Value, "HH:mm") >= Format(rsx.Fields("TglAbsen").Value, "19:30") Then
                            If Format(strWaktuAbsen, "HH:mm") >= Format(strWaktuAbsen, "19:30") Then
                                Call funcSimpanRiwayat("A")
                                If funcSimpanAbsenMasuk(strTglWktAbsen, "A") Then Call subSimpanIDrecordFP(strID, strIdFP)
                            Else
                                strNoRiwayat = rsx.Fields("NoRiwayat").Value
                                If funcUpdateAbsensiMasukMalam(strWaktuAbsen, "U") Then Call subSimpanIDrecordFP(strID, strIdFP)
                            End If
                        Else

                            Call funcSimpanRiwayat("A")
                            If funcSimpanAbsenMasuk(strTglWktAbsen, "A") Then Call subSimpanIDrecordFP(strID, strIdFP)
                        End If

                    Else

                        Call funcSimpanRiwayat("A")
                        If funcSimpanAbsenMasuk(strTglWktAbsen, "A") Then Call subSimpanIDrecordFP(strID, strIdFP)

                    End If
                End If

            Else
                If strID = "" Then Exit Sub
                Call subSimpanIDrecordFP(strID, strIdFP)
            End If

        End If
        dbConnFP.Close
        Call subLoadAbsensiHariIni
    End If
    Set dbConnFP = Nothing
End Sub

Private Function funcSimpanRiwayat(strStatusSP) As Boolean
    On Error GoTo Errload
    funcSimpanRiwayat = True
    Set adoComm = New ADODB.Command
    With adoComm

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglRiwayat", adDate, adParamInput, , Format(Now, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, "181")
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatusSP)
        .Parameters.Append .CreateParameter("OutputNoRiwayat", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Riwayat"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            funcSimpanRiwayat = False
        Else
            If Not IsNull(.Parameters("OutputNoRiwayat").Value) Then strNoRiwayat = .Parameters("OutputNoRiwayat").Value
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
Errload:
End Function

Private Function funcSimpanAbsenMasuk(ByVal strTglMasuk As String, strStatusSP As String) As Boolean
    Dim status As String
    funcSimpanAbsenMasuk = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, IIf(strNoRiwayat = "", Null, strNoRiwayat))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIdPegawaiAbsen)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(strTglMasuk, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("KdStatusAbsensi", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatusSP)
        .Parameters.Append .CreateParameter("outputno", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_ABSENSIPEGAWAI"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Pegawai", vbCritical, "Validasi"
            funcSimpanAbsenMasuk = False
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function

Private Function funcInsertAbsenPulang(ByVal strTglMasuk As String, strStatusSP As String) As Boolean
    Dim status As String
    funcInsertAbsenPulang = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, IIf(strNoRiwayat = "", Null, strNoRiwayat))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIdPegawaiAbsen)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Format(strTglMasuk, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdStatusAbsensi", adChar, adParamInput, 2, "03")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatusSP)
        .Parameters.Append .CreateParameter("outputno", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_ABSENSIPEGAWAIPULANG"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data", vbCritical, "Validasi"
            funcInsertAbsenPulang = False
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function

Private Function funcSimpanAbsensiPulang(ByVal strTglPulang As String) As Boolean
    funcSimpanAbsensiPulang = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_Value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, strNoRiwayat)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIdPegawaiAbsen)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , IIf(strTglPulang = "", Null, Format(strTglPulang, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("kdtstatusabsensi", adChar, adParamInput, 2, "03")

        .ActiveConnection = dbConn
        .CommandText = "Update_AbsensiPegawaiPULANG"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            funcSimpanAbsensiPulang = False
            MsgBox "Error", vbExclamation, "Validasi"
        Else
        
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

Private Function funcUpdateAbsensiMasukMalam(ByVal strTglPulang As String, strStatusSP As String) As Boolean
    funcUpdateAbsensiMasukMalam = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_Value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, strNoRiwayat)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIdPegawaiAbsen)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , IIf(strTglPulang = "", Null, Format(strTglPulang, "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("kdtstatusabsensi", adChar, adParamInput, 2, "03")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatusSP)
        .Parameters.Append .CreateParameter("outputno", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_AbsensiPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            funcUpdateAbsensiMasukMalam = False
            MsgBox "Error", vbExclamation, "Validasi"
        Else
        
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

Private Sub subSimpanIDrecordFP(ByVal IdRecordFP As String, ByVal IdFP As String)
    Dim cmdIdCommand As New ADODB.Command

    IdFP = IIf(Trim(IdFP) = "", Null, IdFP)
    strSQL = "insert into TempIdRecordDariFP (ID, Per_Code, TglData)" & _
    " values ('" & IdRecordFP & "', " & IdFP & ", '" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "')"
    With cmdIdCommand
        .ActiveConnection = dbConn
        .CommandText = strSQL
        .CommandType = adCmdText
        .Execute
    End With
End Sub

Private Function Add_RekapAbsensiDariFingerPrint() As Boolean
    Dim status As String
'    TglAwal = "2014-08-15"
'    TglAkhir = "2014-08-20"
    Add_RekapAbsensiDariFingerPrint = True
    If (TglAwal = "") Then Add_RekapAbsensiDariFingerPrint = False
    If (Add_RekapAbsensiDariFingerPrint = True) Then
        Set adoComm = New ADODB.Command
        With adoComm
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("TglAwal", adDate, adParamInput, , Format(TglAwal, "yyyy/MM/dd 00:00:00"))
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Format(TglAkhir, "yyyy/MM/dd 23:59:59"))
    
            .ActiveConnection = dbConn
            .CommandText = "Add_RekapAbsensiDariFingerPrint"
            .CommandType = adCmdStoredProc
            .Execute
    
            If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                MsgBox "Ada kesalahan dalam penyimpanan data Pegawai", vbCritical, "Validasi"
                Add_RekapAbsensiDariFingerPrint = False
            End If
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End With
    End If
End Function


Private Sub cmdKoneksiDB_Click()
'    tmrAbsensi.Enabled = False
''    frmKoneksiDb.Show vbModal, MDIUtama
'    Call Add_RekapAbsensiDariFingerPrint
    Call subLoadAbsensiHariIni
    'frmKoneksiDb.guidId
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgAbsensi_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgAbsensi
    WheelHook.WheelHook dgAbsensi
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    'Call subLoadAbsensiHariIni
    lbServer.Caption = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Ip Finger Print ")
    lblDatabase.Caption = GetStringValue("HKEY_CURRENT_USER\Software\Medifirst2000\Standard", "Port Finger Print ")
    strFileSettingDBFP = fso.GetSpecialFolder(2) & "\mf2000dbfp.ini"
    'If fso.FileExists(strFileSettingDBFP) Then
    '    Call subLoadDBFP
    'Else
    '    strServerFP = ""
    '    strDatabaseFP = ""
    '    Me.lblDatabase.Caption = ""
    '    Me.lbServer.Caption = ""'

    'End If
    tmrAbsensi.Enabled = True

End Sub

Private Sub txtParameter_Change()
    Call subLoadAbsensiHariIni
End Sub
