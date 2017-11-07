VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLaporanBulananPegawai_VLoop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifist2000 - Laporan Jasa Pelayanan Dokter"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaporanBulananPegawai_VLoop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   14790
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   7560
      Width           =   14775
      Begin MSComctlLib.ProgressBar pbData 
         Height          =   495
         Left            =   600
         TabIndex        =   18
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   11760
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13320
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblPersen 
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9240
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   1080
      Width           =   14775
      Begin VB.Frame Frame5 
         Caption         =   "Periode Pembayaran Piutang/Mutasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         TabIndex        =   14
         Top             =   120
         Width           =   4455
         Begin MSComCtl2.DTPicker dtpAwalPiutang 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            CustomFormat    =   "MMM yyyy "
            Format          =   64487427
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAhkirPiutang 
            Height          =   375
            Left            =   2400
            TabIndex        =   16
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
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
            CustomFormat    =   "MMM yyyy "
            Format          =   64487427
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
            Caption         =   "s/d"
            Height          =   255
            Left            =   2040
            TabIndex        =   17
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.TextBox txtIdDokter 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtNamaDokter 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtJK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Caption         =   "Periode "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9960
         TabIndex        =   2
         Top             =   120
         Width           =   4695
         Begin MSComCtl2.DTPicker dtpPeriode 
            Height          =   375
            Left            =   480
            TabIndex        =   4
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
            CustomFormat    =   "MMMM yyyy "
            Format          =   64487427
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.CommandButton cmdProses 
            Caption         =   "&Proses"
            Height          =   495
            Left            =   3000
            TabIndex        =   0
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Nama Dokter"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Jenis Kelamin"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5295
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   9340
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
      Picture         =   "frmLaporanBulananPegawai_VLoop.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12960
      Picture         =   "frmLaporanBulananPegawai_VLoop.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLaporanBulananPegawai_VLoop.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13455
   End
End
Attribute VB_Name = "frmLaporanBulananPegawai_VLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public iCols As Integer
Public iRows As Integer
Public iCol As Integer
Public iRow As Integer
Public sNamaBulan As String
Public sBlnTglBKM As String
Public sThnTglBKM As String
Public dPeriodeSbelum As Date
Public sPeriodeSbelum As String
Public dPeriodeSbelum2 As Date
Public sPeriodeSbelum2 As String
Public cTotalPembayaranPiutang As Currency
Public rsHitung As ADODB.recordset
Public rsKelompokPasien As ADODB.recordset




Public Sub CreateTabel()
Dim sQuery As String
Dim sBulanField As String
Dim sFieldPembayaran As String
Dim sFieldPembayaranAll As String
Dim sFieldReport As String

    sFieldPembayaran = ""
    sFieldReport = ""
    sFieldPembayaranAll = ""
    sFieldPembayaran = ""
        For i = 1 To iJmlBulanPiutang
            sBulanField = DateAdd("m", -i, dtpPeriode.Value)
            sBulanField = MonthName(Month(sBulanField))
            sFieldPembayaran = sFieldPembayaran & "," & sBulanField & " Money Not Null"
            sFieldReport = sFieldReport & "," & sBulanField
        Next i
     
    
    sQuery = "create Table vRekapPembayaranJasaDokterTemp" & "_" & strNamaHostLocal & _
            " (idDokter char(10)," & _
            " idKelompokPasien char(2)," & _
            " JenisPasien varchar(50)," & _
            " IdPenjamin char(10)," & _
            " Penjamin varchar(50)," & _
            " PembayaranPasien money," & _
            " SaldoAwalHP money ," & _
            " SaldoAwalTRS money ," & _
            " PenambahanHP money," & _
            " PenambahanTRS money" & _
            " " & sFieldPembayaran & "," & _
            " SaldoAhkir money, " & _
            " Diterima money)"
          
    
    ' untuk kebutuhan cetak
    strSQL = "select " & _
            " idDokter," & _
            " idKelompokPasien ," & _
            " JenisPasien ," & _
            " IdPenjamin ," & _
            " Penjamin ," & _
            " PembayaranPasien ," & _
            " SaldoAwalHP ," & _
            " SaldoAwalTRS ," & _
            " PenambahanHP, " & _
            " PenambahanTRS " & _
            " " & sFieldReport & "," & _
            " SaldoAhkir , " & _
            " Diterima " & _
            " from vRekapPembayaranJasaDokterTemp" & "_" & strNamaHostLocal
    
    dbConn.Execute sQuery
    
    
    
    
End Sub
Public Sub DeleteTable()
On Error Resume Next
    sQuery = "drop Table vRekapPembayaranJasaDokterTemp" & "_" & strNamaHostLocal & ""
    dbConn.Execute sQuery

End Sub
Private Sub cmdCetak_Click()
Dim sQuery As String
Dim i As Integer
Dim j As Integer
Dim sValues As String
Dim sValuesMoney As String
Dim iNol As Integer
Dim jNol As Integer
On Error GoTo hell_
    Call DeleteTable
    Call CreateTabel
    With fgData
        For i = 1 To .Rows - 2
        pbData.Max = .Rows - 2
        DoEvents
        lblPersen.Caption = Int((i / pbData.Max) * 100) & "%"
            sValues = ""
            sValuesMoney = ""
            For j = 1 To iCols - 1
                If j = 1 Or j = 2 Or j = 3 Or j = 4 Then
                    sValues = sValues & "," & "'" & .TextMatrix(i, j) & "'"
                Else
                    sValuesMoney = sValuesMoney & "," & "" & msubKonversiKomaTitik(CCur((.TextMatrix(i, j)))) & ""
                    
                End If
            Next j
            
        sValues = "'" & txtIdDokter.Text & "' " & sValues & sValuesMoney
        
        
        
        sQuery = "Insert into vRekapPembayaranJasaDokterTemp" & "_" & strNamaHostLocal & _
                " values (" & _
                " " & sValues & "" & _
                " )"
        dbConn.Execute sQuery
        pbData.Value = Int(pbData.Value) + 1
        Next i

    
    End With
MsgBox "data telah tersimpan ke tabel temporary!!", vbInformation + vbOKOnly, "informasi"
pbData.Value = 0.0001
sPembayaranSebelumBulanPilih = MonthName(Month(DateAdd("m", -1, dtpPeriode.Value)))
sPembayaranSebelumBulanPilih2 = MonthName(Month(DateAdd("m", -2, dtpPeriode.Value)))
frmRekapKomponenJasaPerDokter_NEW.Show
Exit Sub
hell_:
    pbData.Value = 0.0001
    msubPesanError
End Sub


Private Sub cmdProses_Click()
Dim i As Integer
Dim j As Integer
On Error GoTo hell_
'dtpPeriode.Enabled = False
MousePointer = vbHourglass
'Call subPeriodeSbelumBulanAktif
'Call subLoadDefault
Call setGrid
Call LoadKolomSumberPendapatan
If MsgBox("yakin akan memproses data jasa dokter??", vbInformation + vbYesNo, "informasi") = vbNo Then MousePointer = vbDefault: Exit Sub
For iRow = 1 To iRows
pbData.Max = iRows
DoEvents
lblPersen.Caption = Int((iRow / iRows) * 100) & "%"
    For iCol = 1 To iCols
    'Select Case fgData.Col
        'Case 2
      If iCol = 5 Then Call LoadKolomBulanPembayaranPasien ' bulan
      If iCol = 6 Then Call LoadKolomSaldoSebelumBlnAktifHP 'saldo awal HP
      If iCol = 7 Then Call LoadKolomSaldoSebelumBlnAktifTRS 'saldo awal TRS
      If iCol = 8 Then Call LoadKolomSaldoBulanAktifHP ' penambahan
      If iCol = 9 Then Call LoadKolomSaldoBulanAktifTRS ' penambahan
      If iCol = 10 Then Call LoadKolomSaldoKlaimCairSebelumBulanAktif ' pembayaran
      If iCol = 10 + iJmlBulanPiutang Then Call LoadHitungSaldo ' saldo ahkir
      If iCol = 11 + iJmlBulanPiutang Then Call LoadKolomDiterimaDokter ' diterima
      
    'End Select
    Next iCol
   ' fgData.Rows = iRow + 2
pbData.Value = Int(pbData.Value) + 1
Next iRow
MousePointer = vbDefault
MsgBox "data berhasil diproses", vbInformation, "informasi"
pbData.Value = 0.0001
cmdCetak.SetFocus
Exit Sub
hell_:
    pbData.Value = 0.0001
    MousePointer = vbDefault
    msubPesanError
    Exit Sub
End Sub
Private Sub subPeriodeSbelumBulanAktif()
    dPeriodeSbelum = DateAdd("m", -1, Format(dtpPeriode.Value, dd - MM - yyyy))
    dtpAwalPiutang.Value = DateAdd("m", -1, Format(dPeriodeSbelum, dd - MM - yyyy))
    sPeriodeSbelum = MonthName(Month(dPeriodeSbelum))
    
    dPeriodeSbelum2 = DateAdd("m", -1, Format(dPeriodeSbelum, dd - MM - yyyy))
    sPeriodeSbelum2 = MonthName(Month(dPeriodeSbelum2))
End Sub

Private Sub cmdSimpan_Click()
    Call CreateTabel
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
Public Sub LoadHitungSaldo()
Dim cHitungSaldo As Currency
    ' (KolomSaldoSebelumBlnAktif + LoadKolomSaldoBulanAktif ) - kolomSaldoKlaimCairSebelumBulanAktif
    cHitungSaldo = (CCur(fgData.TextMatrix(iRow, 5)) + CCur(fgData.TextMatrix(iRow, 6)) + CCur(fgData.TextMatrix(iRow, 7)) + CCur(fgData.TextMatrix(iRow, 8)) + CCur(fgData.TextMatrix(iRow, 9))) - (CCur(fgData.TextMatrix(iRow, 10)) + CCur(fgData.TextMatrix(iRow, 11)))
    If cHitungSaldo < 0 Then fgData.TextMatrix(iRow, iCol) = Format(CCur(fgData.TextMatrix(iRow, 8)), "#,###.00") Else fgData.TextMatrix(iRow, iCol) = Format(cHitungSaldo, "#,###.00")
    
    cGrandTotalSaldoAhkir = cGrandTotalSaldoAhkir + cHitungSaldo
End Sub
Public Sub LoadKolomSumberPendapatan()
    Set rs = Nothing
    If sp_LoadKelompokPasien = False Then Exit Sub
    
    iRows = rsKelompokPasien.RecordCount
    For i = 1 To rsKelompokPasien.RecordCount
        fgData.Rows = 2 + i
        fgData.TextMatrix(i, 1) = rsKelompokPasien.Fields("KdKelompokPasien") ' kolom 1
        fgData.TextMatrix(i, 2) = rsKelompokPasien.Fields("JenisPasien") ' kolom 2
        fgData.TextMatrix(i, 3) = rsKelompokPasien.Fields("IdPenjamin") ' kolom 3
        fgData.TextMatrix(i, 4) = rsKelompokPasien.Fields("NamaPenjamin") ' kolom 4
        rsKelompokPasien.MoveNext
    Next i
    
    rsKelompokPasien.Close
End Sub



Private Sub LoadKolomBulanPembayaranPasien()
'Uang Tunai/Cash yang bukan berasal dari Piutang Klaim Cair dan sudah dibayarkan ke Dokter (biasanya pada Bulan Y)
' kolom 3
' NoBKK is null = belum dibayarkan ke dokter
 '   Set rs = Nothing

  Set rs = Nothing
  If HitungKolom("JmlBayar", Month(Format(dPeriodeSbelum, "dd MM yyyy")), Year(Format(dPeriodeSbelum, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
       If rsHitung.Fields("JmlBayar") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlBayar")) Then
            fgData.TextMatrix(iRow, iCol) = 0
        Else
            fgData.TextMatrix(iRow, iCol) = Format(rsHitung.Fields("JmlBayar"), "#,###.00")
            
        End If
        
       rsHitung.Close
       
End Sub
Private Function HitungKolom(sKolom As String, sBulan As String, sTahun As String, sKdKelompokPasien As String, sIdPenjamin As String) As Boolean
    On Error GoTo hell_
    HitungKolom = True
    Set dbcmd = New ADODB.Command
    Set rsHitung = New ADODB.recordset
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, Trim(txtIdDokter.Text))
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, sKdKelompokPasien)
        .Parameters.Append .CreateParameter("idPenjamin", adChar, adParamInput, 10, sIdPenjamin)
        .Parameters.Append .CreateParameter("Bulan", adVarChar, adParamInput, 20, sBulan)
        .Parameters.Append .CreateParameter("Tahun", adVarChar, adParamInput, 4, sTahun)
        .Parameters.Append .CreateParameter("Kolom", adVarChar, adParamInput, 30, sKolom)
        
        .ActiveConnection = dbConn
        .CommandText = "HitungRekapitulasiJasaDokterPerKolom"
        .CommandType = adCmdStoredProc
        .Execute
        
        
        Set rsHitung = .Execute
        
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Hapus Rekap Komponen Biaya Pelayanan", vbCritical, "Validasi"
            HitungKolom = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        
       
        
    End With
    
Exit Function
hell_:
    HitungKolom = False
    Call msubPesanError("-HitungRekapitulasiJasaDokterPerKolom")
End Function
Private Function sp_LoadKelompokPasien() As Boolean
    On Error GoTo hell_
    sp_LoadKelompokPasien = True
    Set dbcmd = New ADODB.Command
    Set rsKelompokPasien = New ADODB.recordset
    With dbcmd
    
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, Trim(txtIdDokter.Text))
        .ActiveConnection = dbConn
        
        .CommandText = "RekapKomponenBPRemunerasiKelompokPasien_S"
        .CommandType = adCmdStoredProc
        .Execute
        
        
        Set rsKelompokPasien = .Execute
        
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam load kelompok pasien", vbCritical, "Validasi"
            sp_LoadKelompokPasien = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        
       
        
    End With
    
Exit Function
hell_:
    sp_LoadKelompokPasien = False
    Call msubPesanError("-RekapKomponenBPRemunerasiKelompokPasien_S")
End Function
Private Sub LoadKolomSaldoSebelumBlnAktifHP()
'Piutang Klaim yang belum dan sudah cair (Lunas + Sisa Tagihan Piutang Klaim) pada bulan-bulan sebelum bulan Aktif/Pilih tapi belum dibayarkan ke Dokter
' kolom 4
' noBKK is null = blum di bayarkan
Dim cSaldoAwalHp As Currency
Dim cSaldiAwalHpMin As Currency
Dim cSaldoAwalNow As Currency
Dim iPeriodeMax As Integer
Dim dCateoff As Date
Dim cPembayaran As Currency
Dim cSaldoAwalHpMinNow As Currency
' ambil pembayran sebelumnya
dCuteoff = Format("1/7/2009", "dd/MM/yyyy")
cSaldoAwalHpMin = 0
cPembayaran = 0
iPeriodeMax = DateDiff("m", dCuteoff, dtpPeriode.Value)


If iPeriodeMax = 1 Then GoTo langsung_
For i = 1 To iPeriodeMax
If fgData.TextMatrix(iRow, 2) = "ASKES LOKAL GOL III" Then
    MsgBox "ok"
End If
If HitungKolom("SaldoAwalHPMinus", Month(Format(DateAdd("m", -i, Format(dtpPeriode.Value, dd - MM - yyyy)), "dd MM yyyy")), Year(Format(dtpPeriode.Value, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
If rsHitung.Fields("JmlBayarKlaimHPMin") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlBayarKlaimHPMin")) Then
    'tampung total pembayaran
Else
    cPembayaran = cPembayaran + rsHitung.Fields("JmlBayarKlaimHPMin")
End If

If HitungKolom("SaldoAwalHP", Month(Format(DateAdd("m", -i, Format(dPeriodeSbelum2, dd - MM - yyyy)), "dd MM yyyy")), Year(Format(dPeriodeSbelum2, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
If rsHitung.Fields("JmlHutangPenjamin") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlHutangPenjamin")) Then
Else
    cSaldoAwalHpMin = cSaldoAwalHpMin + CCur(Format(rsHitung.Fields("JmlHutangPenjamin"), "#,###.00"))
End If
        
rsHitung.Close
Next i
cSaldoAwalHpMinNow = cSaldoAwalHpMin - cPembayaran
If cSaldoAwalHpMinNow <= 0 Then cSaldoAwalHpMinNow = 0
'        Else
'           cSaldiAwalHpMin = Format(rsHitung.Fields("JmlBayarKlaimHPMin"), "#,###.00")
'        End If
langsung_:
If HitungKolom("SaldoAwalHP", Month(Format(dPeriodeSbelum2, "dd MM yyyy")), Year(Format(dPeriodeSbelum2, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
        If rsHitung.Fields("JmlHutangPenjamin") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlHutangPenjamin")) Then
            cSaldoAwalHp = 0
        Else
            cSaldoAwalHp = Format(rsHitung.Fields("JmlHutangPenjamin"), "#,###.00")
        End If
rsHitung.Close

cSaldoAwalNow = cSaldoAwalHp + cSaldoAwalHpMinNow
If cSaldoAwalNow <= 0 Then
    fgData.TextMatrix(iRow, iCol) = 0
Else
    fgData.TextMatrix(iRow, iCol) = Format(cSaldoAwalNow, "#,###.00")
End If


End Sub
Private Sub LoadKolomSaldoSebelumBlnAktifTRS()
'Piutang Klaim yang belum dan sudah cair (Lunas + Sisa Tagihan Piutang Klaim) pada bulan-bulan sebelum bulan Aktif/Pilih tapi belum dibayarkan ke Dokter
' kolom 4
' noBKK is null = blum di bayarkan
Dim cJmlTanggunganRSMin As Currency
Dim cJmlTanggunganRS As Currency
Dim cJmlTanggunganRSNow As Currency
Dim iPeriodeMax As Integer
Dim dCateoff As Date
Dim cPembayaran As Currency
Dim cJmlTanggunganRSMinNow As Currency
' ambil pembayran sebelumnya
dCuteoff = Format("1/7/2009", "dd/MM/yyyy")
cJmlTanggunganRSMin = 0
iPeriodeMax = DateDiff("m", dCuteoff, dtpPeriode.Value)

If iPeriodeMax = 1 Then GoTo langsung_
For i = 1 To iPeriodeMax
If HitungKolom("SaldoAwalTRSMin", Month(Format(DateAdd("m", -i, Format(dtpPeriode.Value, dd - MM - yyyy)), "dd MM yyyy")), Year(Format(dtpPeriode.Value, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
If rsHitung.Fields("JmlTanggunganRSMin") = 0 Or rsHitung.Fields("JmlTanggunganRSMin") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlTanggunganRSMin")) Then
Else
    cPembayaran = cPembayaran + rsHitung.Fields("JmlTanggunganRSMin")
End If
If HitungKolom("SaldoAwalTRS", Month(Format(DateAdd("m", -i, Format(dPeriodeSbelum2, dd - MM - yyyy)), "dd MM yyyy")), Year(Format(dPeriodeSbelum2, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
If rsHitung.Fields("JmlTanggunganRS") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlTanggunganRS")) Then
Else
    fgData.TextMatrix(iRow, iCol) = Format(rsHitung.Fields("JmlTanggunganRS"), "#,###.00")
    cJmlTanggunganRSMin = cJmlTanggunganRSMin + fgData.TextMatrix(iRow, iCol)
End If
rsHitung.Close

Next i
cJmlTanggunganRSMinNow = cJmlTanggunganRSMin - cPembayaran
If cJmlTanggunganRSMinNow <= 0 Then cJmlTanggunganRSMinNow = 0
langsung_:
If HitungKolom("SaldoAwalTRS", Month(Format(dPeriodeSbelum2, "dd MM yyyy")), Year(Format(dPeriodeSbelum2, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
        If rsHitung.Fields("JmlTanggunganRS") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlTanggunganRS")) Then
            fgData.TextMatrix(iRow, iCol) = 0
        Else
            fgData.TextMatrix(iRow, iCol) = Format(rsHitung.Fields("JmlTanggunganRS"), "#,###.00")
            cJmlTanggunganRS = fgData.TextMatrix(iRow, iCol)
        End If
rsHitung.Close

cJmlTanggunganRSNow = cJmlTanggunganRS + cJmlTanggunganRSMinNow
If cJmlTanggunganRSNow <= 0 Then
    fgData.TextMatrix(iRow, iCol) = 0
Else
    fgData.TextMatrix(iRow, iCol) = Format(cJmlTanggunganRSNow, "#,###.00")
End If

End Sub
Private Sub LoadKolomSaldoBulanAktifHP()
'Piutang Klaim yang belum dan sudah cair (Lunas + Sisa Tagihan Piutang Klaim) pada bulan Aktif/Pilih tapi belum dibayarkan ke Dokter
' kolom 5
' noBKK is null = blum di bayarkan
'    Set rs = Nothing
'
If HitungKolom("PenambahanHP", Month(Format(dtpAwalPiutang.Value, "dd MM yyyy")), Year(Format(dtpAwalPiutang.Value, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
       If rsHitung.Fields("JmlHutangPenjamin") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlHutangPenjamin")) Then
           fgData.TextMatrix(iRow, iCol) = 0
        Else
            fgData.TextMatrix(iRow, iCol) = Format(rsHitung.Fields("JmlHutangPenjamin"), "#,###.00")
        End If
rsHitung.Close
End Sub
Private Sub LoadKolomSaldoBulanAktifTRS()
'Piutang Klaim yang belum dan sudah cair (Lunas + Sisa Tagihan Piutang Klaim) pada bulan Aktif/Pilih tapi belum dibayarkan ke Dokter
' kolom 5
' noBKK is null = blum di bayarkan
'
If HitungKolom("PenambahanTRS", Month(Format(dtpAwalPiutang.Value, "dd MM yyyy")), Year(Format(dtpAwalPiutang.Value, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
       If rsHitung.Fields("JmlTanggunganRS") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlTanggunganRS")) Then
           fgData.TextMatrix(iRow, iCol) = 0
        Else
            fgData.TextMatrix(iRow, iCol) = Format(rsHitung.Fields("JmlTanggunganRS"), "#,###.00")
        End If
rsHitung.Close
End Sub
Private Sub LoadKolomSaldoKlaimCairSebelumBulanAktif()
'Piutang Klaim Cair yang belum dibayarkan ke Dokter
' kolom 6
' noBKK is null = blum di bayarkan

Dim i As Integer
On Error GoTo hell_
cTotalPembayaranPiutang = 0
    dPeriodeSbelum = DateAdd("m", -1, Format(dtpPeriode.Value, dd - MM - yyyy))
    If HitungKolom("Pembayaran", Month(Format(dPeriodeSbelum, "dd MM yyyy")), Year(Format(dPeriodeSbelum, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
         On Error Resume Next
         If rsHitung.Fields("JmlBayarKlaim") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlBayarKlaim")) Then
             fgData.TextMatrix(iRow, iCol) = 0
         Else
             fgData.TextMatrix(iRow, iCol) = Format(rsHitung.Fields("JmlBayarKlaim"), "#,###.00")
             cTotalPembayaranPiutang = cTotalPembayaranPiutang + fgData.TextMatrix(iRow, iCol)
'             fgData.TextMatrix(iRow, iCol) = CCur(fgData.TextMatrix(iRow, 6)) + CCur(fgData.TextMatrix(iRow, 7))
'             cTotalPembayaranPiutang = Format(fgData.TextMatrix(iRow, iCol), "#,###.00")
         End If
    rsHitung.Close

    If HitungKolom("PembayaranHP", Month(Format(dtpPeriode.Value, "dd MM yyyy")), Year(Format(dtpPeriode.Value, " dd MM yyyy")), fgData.TextMatrix(iRow, 1), fgData.TextMatrix(iRow, 3)) = False Then Exit Sub
         On Error Resume Next
         If rsHitung.Fields("JmlBayarKlaimHP") = "" Or rsHitung.EOF = True Or IsNull(rsHitung.Fields("JmlBayarKlaimHP")) Then
             fgData.TextMatrix(iRow, iCol + 1) = 0
         Else
             fgData.TextMatrix(iRow, iCol + 1) = Format(rsHitung.Fields("JmlBayarKlaimHP"), "#,###.00")
             cTotalPembayaranPiutang = cTotalPembayaranPiutang + fgData.TextMatrix(iRow, iCol + 1)
             'fgData.TextMatrix(iRow, iCol + 1) = CCur(fgData.TextMatrix(iRow, 6)) + CCur(fgData.TextMatrix(iRow, 7))
             'cTotalPembayaranPiutang = cTotalPembayaranPiutang + Format(fgData.TextMatrix(iRow, iCol + 1), "#,###.00")
             
             
         End If
    rsHitung.Close
Exit Sub
hell_:
    
End Sub
Private Sub LoadKolomDiterimaDokter()
'jasa dokter yang diterima oleh dokter
' kolom 8
Dim cDiterima As Currency
        cDiterima = cTotalPembayaranPiutang
     
        If cDiterima = 0 Then
            fgData.TextMatrix(iRow, iCol) = 0
        Else
            fgData.TextMatrix(iRow, iCol) = Format(cDiterima, "#,###.00")
            
        End If
        
        
   
End Sub

Private Sub dtpPeriode_Change()
    dtpAhkirPiutang.Value = Format(DateAdd("m", -1, dtpPeriode.Value), "MM yyyy")
    dtpAwalPiutang.Value = Format(DateAdd("m", -2, dtpPeriode.Value), "MM yyyy")
    
    Call subPeriodeSbelumBulanAktif
    Call subLoadDefault
    Call setGrid
    cmdProses.SetFocus
End Sub

Private Sub dtpPeriode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   ' dtpAwalPiutang.Value = Format(DateAdd("m", -1, dtpPeriode.Value), "MM yyyy")
    dtpAhkirPiutang.Value = Format(DateAdd("m", -1, dtpPeriode.Value), "MM yyyy")
    dtpAwalPiutang.Value = Format(DateAdd("m", -2, dtpPeriode.Value), "MM yyyy")
    
    Call subPeriodeSbelumBulanAktif
    Call subLoadDefault
    Call setGrid
    cmdProses.SetFocus
End If
End Sub

Private Sub dtpPeriode_KeyPress(KeyAscii As Integer)
 '   Call dtpPeriode_KeyDown(13)
End Sub

Private Sub Form_Load()
On Error GoTo errFormLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    dtpPeriode.Value = Format(Now, "MMMM yyyy")
    dtpAwalPiutang.Value = Format(DateAdd("m", -1, dtpPeriode.Value), "MM yyyy")
    dtpAhkirPiutang.Value = Format(dtpPeriode.Value, "MM yyyy")
    dtpAhkirPiutang.Enabled = False
    dtpAwalPiutang.Enabled = False
    
    dtpPeriode.Enabled = True
    Call subPeriodeSbelumBulanAktif
    Call subLoadDefault
    Call setGrid
    Call LoadDataDokter
 
    
    
Exit Sub
errFormLoad:
    msubPesanError
End Sub
Private Sub subLoadDefault()
'set default
    cTotalPembayaranPiutang = 0
    iJmlBulanPiutang = DateDiff("m", dtpAwalPiutang.Value, dtpAhkirPiutang.Value)
    iJmlBulanPiutang = iJmlBulanPiutang
   ' dPeriodeSbelum = DateAdd("m", -1, Format(dtpPeriode.Value, dd - MM - yyyy))
    dtpAwalPiutang.Value = Format(DateAdd("m", -1, dtpPeriode.Value), "MM yyyy")
    dtpAhkirPiutang.Value = Format(dtpPeriode.Value, "MM yyyy")
    iCols = 12 + iJmlBulanPiutang 'nilai default
    sNamaBulan = MonthName(Month(dPeriodeSbelum))
    sSumberPendapatan = "Sumber Pendapatan "  ' jenis pasien
    sSumberPendapatanPenjamin = "Nama Penjamin "  ' penjamin jenis pasien
    sBulanPembayaranPasien = sNamaBulan   'Uang Tunai/Cash yang BUKAN berasal dari Piutang Klaim Cair tapi belum dibayarkan ke Dokter
    sSaldoBulanSebelumBulanPilih = "Saldo HP " & sPeriodeSbelum2 'Piutang Klaim yang belum dan sudah cair (Lunas + Sisa Tagihan Piutang Klaim) pada bulan-bulan sebelum bulan Aktif/Pilih tapi belum dibayarkan ke Dokter
    sSaldoBulanSebelumBulanPilihTRS = "Saldo TRS " & sPeriodeSbelum2 'Piutang Klaim yang belum dan sudah cair (Lunas + Sisa Tagihan Piutang Klaim) pada bulan-bulan sebelum bulan Aktif/Pilih tapi belum dibayarkan ke Dokter
    sSaldoBulanPilih = "Pen. HP " & sNamaBulan 'Piutang Klaim yang belum dan sudah cair (Lunas + Sisa Tagihan Piutang Klaim) pada bulan Aktif/Pilih tapi belum dibayarkan ke Dokter
    sSaldoBulanPilihTRS = "Pen. TRS " & sNamaBulan
    sPembayaranSebelumBulanPilih = "Pembayaran " & sPeriodeSbelum
    sSaldoTotal = "Saldo Akhir "
End Sub
Private Sub LoadDataDokter()
    Set rs = Nothing
    strSQL = "SELECT    idPegawai, NamaLengkap, JenisKelamin" & _
             " FROM  DataPegawai where IdPegawai = '" & strIDPegawaiAktif & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF Then Exit Sub
    txtNamaDokter.Text = rs.Fields("NamaLengkap")
    If rs.Fields("JenisKelamin") = "P" Then
        txtJK.Text = "Perempuan"
    Else
        txtJK.Text = "Laki-Laki"
    End If
    txtIdDokter.Text = rs.Fields("idPegawai")
End Sub

Private Sub setGrid()
'Dim i As Integer
    With fgData
        .clear
        .Rows = 2
        .Cols = 14
        .Row = 0
        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 500
            
        Next
       
        .TextMatrix(0, 1) = "KodeSumberPendapatan"
        .TextMatrix(0, 2) = sSumberPendapatan ' jenis pasien
        .TextMatrix(0, 3) = "Kode Penjamin Pasien" ' penjamin jenis pasien
        .TextMatrix(0, 4) = sSumberPendapatanPenjamin ' penjamin jenis pasien
        .TextMatrix(0, 5) = sBulanPembayaranPasien ' bulan
        .TextMatrix(0, 6) = sSaldoBulanSebelumBulanPilih ' saldo awal HP
        .TextMatrix(0, 7) = sSaldoBulanSebelumBulanPilihTRS ' saldo awal TRS
        .TextMatrix(0, 8) = sSaldoBulanPilih ' penambahan
        .TextMatrix(0, 9) = sSaldoBulanPilihTRS ' penambahan
        sPembayaranSebelumBulanPilih = DateAdd("m", -1, dtpPeriode.Value)
        .TextMatrix(0, 10) = MonthName(Month(sPembayaranSebelumBulanPilih))  ' pembayaran piutang
        .TextMatrix(0, 11) = MonthName(Month(DateAdd("m", -1, sPembayaranSebelumBulanPilih)))  ' pembayaran piutang
       
        .TextMatrix(0, 12) = sSaldoTotal  ' saldo ahkir
        .TextMatrix(0, 13) = "Diterima"
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 1800
        .ColWidth(3) = 0
        .ColWidth(4) = 2000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1200
        .ColWidth(13) = 1200
        
        
    End With
End Sub
