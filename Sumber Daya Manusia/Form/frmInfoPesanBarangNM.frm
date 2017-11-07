VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInfoPesanBarangNM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Pemesanan Barang"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoPesanBarangNM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13830
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   7920
      Width           =   13815
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   520
         Left            =   12240
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Ceta&k"
         Height          =   520
         Left            =   10560
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informasi Pemesanan Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   13815
      Begin VB.Frame Frame4 
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
         Left            =   7920
         TabIndex        =   11
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton cmdTampilkan 
            Caption         =   "&Cari"
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
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpTglAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   118226947
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpTglAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   3
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   118226947
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status Terima"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   3255
         Begin VB.OptionButton optOrder 
            Caption         =   "Order"
            Height          =   375
            Left            =   480
            TabIndex        =   0
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTerima 
            Caption         =   "Terima"
            Height          =   375
            Left            =   2040
            TabIndex        =   1
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid dgInfoPesanBrg 
         Height          =   5535
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   9763
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
               LCID            =   1033
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
               LCID            =   1033
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
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      Left            =   12000
      Picture         =   "frmInfoPesanBarangNM.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmInfoPesanBarangNM.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInfoPesanBarangNM.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "frmInfoPesanBarangNM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
    On Error GoTo hell
    mdTglAwal = dtpTglAwal.Value
    mdTglAkhir = dtpTglAkhir.Value
    Call cmdTampilkan_Click
    vLaporan = ""
    
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
   
    If optOrder.Value = True Then
        frmCetakInfoPesanBarang.Show
    Else
    
       frmCetakInfoKirimBarang.Show
    End If
    
'Exit Sub
hell:
End Sub

Private Sub cmdTampilkan_Click()
    If mstrKdKelompokBarang = "02" Then     'medis
        If optOrder.Value = True Then
           Set rs = Nothing
           strSQL = "select * from V_InfoPemesananBrgRuangan where NoKirim is null and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
 
        Else
           Set rs = Nothing
           strnonmedis = False
          ' strSQL = "select * from V_InfoPemesananBrgRuangan where NoKirim is not null and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
           strSQL = "select * from V_InfoPengirimanBrgRuangan where NoKirim is not null and KdRuanganTujuan='" & mstrKdRuangan & "' and([Tgl. Kirim] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"

        End If
        
    ElseIf mstrKdKelompokBarang = "01" Then     'non medis
    
        If optOrder.Value = True Then
           Set rs = Nothing
           strSQL = "select * from V_InfoPemesananBrgRuanganNM where NoKirim is null and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
           ' strSQL = "select * from V_InfoPemesananBrgRuanganNM where NoKirim is not null AND NoKirim IS NULL and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
        Else
           Set rs = Nothing
           strnonmedis = True
            strSQL = "select * from V_InfoPengirimanBrgRuanganNM where NoKirim is not null and KdRuanganTujuan='" & mstrKdRuangan & "' and([Tgl. Kirim] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
          ' strSQL = "select * from V_InfoPemesananBrgRuanganNM where NoKirim is not null and KdRuangan='" & mstrKdRuangan & "' and([Tgl. Pesan] between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
        End If
    End If
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgInfoPesanBrg.DataSource = rs
    With dgInfoPesanBrg
        .Columns(0).Width = 1900
        
        .Columns(1).Width = 1900
        
       ' If optOrder.Value = True Then
         .Columns(2).Width = 1900
      '  Else
      '    .Columns(2).Width = 0
      '  End If
        
        .Columns(3).Width = 1200
        .Columns(4).Width = 2700
        .Columns(5).Width = 1000
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 1800
        .Columns(7).Width = 1200
        .Columns(8).Width = 0
    End With
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgInfoPesanBrg_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgInfoPesanBrg
    WheelHook.WheelHook dgInfoPesanBrg
End Sub

Private Sub dtpTglAkhir_Change()
    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       optOrder.SetFocus
    End If
End Sub

Private Sub dtpTglAwal_Change()
    dtpTglAwal.MaxDate = Now
End Sub

Private Sub dtpTglAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       dtpTglAkhir.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call openConnection
    optOrder.Value = True
    dtpTglAkhir.Value = Now
    dtpTglAwal.Value = Now
    Call cmdTampilkan_Click
End Sub

Private Sub optOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdTampilkan.SetFocus
    End If
End Sub



