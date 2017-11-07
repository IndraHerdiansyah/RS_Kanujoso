VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSlipInsentifPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Slip Insentif Pegawai"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "frmSlipInsentifPegawaii.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8550
   Begin MSComCtl2.DTPicker dtpBulan 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   6720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      CustomFormat    =   "MMM yyyy"
      Format          =   94568451
      CurrentDate     =   42199
   End
   Begin VB.CheckBox chkSemuaa 
      Caption         =   "&Cek Semua"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
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
      Left            =   5640
      TabIndex        =   5
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   8895
      Visible         =   0   'False
      Width           =   1335
   End
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
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   8895
      Visible         =   0   'False
      Width           =   1335
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
   Begin MSComctlLib.ListView lvPenjamin 
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Bulan Tahun :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmSlipInsentifPegawaii.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmSlipInsentifPegawaii.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSlipInsentifPegawaii.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmSlipInsentifPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSemuaa_Click()
    If chkSemuaa.Value = Checked Then
        Call LoadDataCombo(True)
    Else
        Call LoadDataCombo(False)
    End If
End Sub


Private Sub cmdCetak_Click()
    On Error GoTo errLoad

    Dim i As Integer
    Dim boolXXX As Boolean
    Dim boolX As Boolean
    Dim strDftRuang As String
    Dim strDftPenjamin As String
    Dim mstrFilterr As String
    Dim mstrFilter As String

    TglCetak = Now
    mstrFilter = ""
    mstrFilterr = ""
    strDaftarRuangan = ""
    strDaftarPenjamin = ""


    

        boolX = False
        For i = 1 To lvPenjamin.ListItems.Count
            If lvPenjamin.ListItems(i).Checked = True Then
                strDftPenjamin = strDftPenjamin & "'" & lvPenjamin.ListItems(i).Text & "',"
                strDaftarPenjamin = strDaftarPenjamin & lvPenjamin.ListItems(i).Text & ", "
                boolX = True
            End If
        Next i
    
        If boolX = True Then
            strDftPenjamin = "  Nama In (" & Mid(Trim(strDftPenjamin), 1, Len(Trim(strDftPenjamin)) - 1) & ")"
            strDaftarPenjamin = Mid(Trim(strDaftarPenjamin), 1, Len(Trim(strDaftarPenjamin)) - 1)
            boolX = False
            mstrFilter = mstrFilter & strDftPenjamin
        End If
    

            strSQL = "select * from SlipInsentifPegawai  where BulanTahun ='" & Format(dtpBulan.Value, "yymm") & "' and " & mstrFilter
            Call msubRecFO(rs, strSQL)
            If rs.RecordCount = 0 Then
                MsgBox "Tidak ada Pelayanan dengan Dokter ini", vbInformation, "Infromasi"
            Else
            'Exit Sub
                frm_cetak_SlipInsentif.tgl = UCase(Format(dtpBulan.Value, "MMMM yyyy"))
                frm_cetak_SlipInsentif.Show
            End If
'    ElseIf strCetak = "LaporanPenerimaanperTanggal" Then
'        boolXXX = False
'        For i = 1 To lvRuanganKasir.ListItems.Count
'            If lvRuanganKasir.ListItems.Item(i).Checked = True Then
'                mstrFilter = mstrFilter & "'" & lvRuanganKasir.ListItems(i).Text & "',"
'                mstrFilterr = mstrFilterr & lvRuanganKasir.ListItems(i).Text & "',"
'                boolXXX = True
'            End If
'        Next i
'
'        If boolXXX = True Then
'            mstrFilter = " And NamaRuanganKasir In (" & Mid(Trim(mstrFilter), 1, Len(Trim(mstrFilter)) - 1) & ")"
'            mstrFilterr = Mid(Trim(mstrFilterr), 1, Len(Trim(mstrFilterr)) - 1)
'            boolXXX = False
'        End If
'
'        If chkInstalasi.value = 1 Then
'            If Periksa("datacombo", dcInstalasi, "Nama instalasi kosong") = False Then Exit Sub
'            If dcInstalasi.BoundText = "02" Then
'                mstrFilter = mstrFilter & " AND instalasiPelayanan IN('Instalasi Ibu & Anak','Instalasi Rehabilitasi Medis','Instalasi Rawat Jalan' )"
'            Else
'                mstrFilter = mstrFilter & " AND instalasiPelayanan = '" & dcInstalasi.Text & "'"
'            End If
'        End If
'
'        strDftRuang = ""
'        boolXXX = False
'        For i = 1 To lvRuangan.ListItems.Count
'            If lvRuangan.ListItems(i).Checked = True Then
'                strDftRuang = strDftRuang & "'" & lvRuangan.ListItems(i).Text & "',"
'                strDaftarRuangan = strDaftarRuangan & lvRuangan.ListItems(i).Text & ", "
'                boolXXX = True
'            End If
'        Next i
'
'        If boolXXX = True Then
'            strDftRuang = " And NamaRuangan In (" & Mid(Trim(strDftRuang), 1, Len(Trim(strDftRuang)) - 1) & ")"
'            strDaftarRuangan = Mid(Trim(strDaftarRuangan), 1, Len(Trim(strDaftarRuangan)) - 1)
'            boolXXX = False
'            mstrFilter = mstrFilter & strDftRuang
'        Else
'        End If
'
'        strDftPenjamin = ""
'        boolX = False
'        For i = 1 To lvPenjamin.ListItems.Count
'            If lvPenjamin.ListItems(i).Checked = True Then
'                strDftPenjamin = strDftPenjamin & "'" & lvPenjamin.ListItems(i).Text & "',"
'                strDaftarPenjamin = strDaftarPenjamin & lvPenjamin.ListItems(i).Text & ", "
'                boolX = True
'            End If
'        Next i
'
'        If boolX = True Then
'            strDftPenjamin = " And Kasir In (" & Mid(Trim(strDftPenjamin), 1, Len(Trim(strDftPenjamin)) - 1) & ")"
'            strDaftarPenjamin = Mid(Trim(strDaftarPenjamin), 1, Len(Trim(strDaftarPenjamin)) - 1)
'            boolX = False
'            mstrFilter = mstrFilter & strDftPenjamin
'        Else
'        End If
'        strSQL = "select DISTINCT * from V_PenerimaanHarianPerTanggal where TglBKM Between '" & Format(dtpAwal.value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.value, "yyyy/MM/dd 23:59:59") & "'" & mstrFilter
'        Call msubRecFO(rs, strSQL)
'        If rs.RecordCount = 0 Then
'            MsgBox "Tidak ada Pembayaran dengan user ini", vbInformation, "Infromasi"
'        'Exit Sub
'        Else
'            frm_cetak_LaporanPenerimaanHarian.Show
'        End If
'    End If
    

    Exit Sub
errLoad:
    cmdCetak.Enabled = False
End Sub


Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call LoadDataCombo(False)
    dtpBulan.Value = Now()
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub lvPenjamin_BeforeLabelEdit(Cancel As Integer)
    If lvPenjamin.ListItems(Item.key).Checked = True Then
        lvPenjamin.ListItems(Item.key).ForeColor = vbBlue
    Else
        lvPenjamin.ListItems(Item.key).ForeColor = vbBlack
    End If
End Sub



Private Sub LoadPetugasKasir()
    On Error Resume Next
    strSQL = " SELECT   DISTINCT  DataPegawai.NamaLengkap,biayapelayanan.IdPegawai" & _
             " FROM         biayapelayanan INNER JOIN " & _
             " DataPegawai ON biayapelayanan.IdPegawai = DataPegawai.IdPegawai   where DataPegawai.IdPegawai not in ('8888888888','1111111111','2222222222') and DataPegawai.KdJenisPegawai='001'" & _
             " order by DataPegawai.NamaLengkap"
    Call msubRecFO(rs, strSQL)
    Do
        lvPenjamin.ListItems.add , "A" & rs(1).Value, rs(0).Value
        rs.MoveNext
    Loop Until rs.EOF
    lvPenjamin.ListItems("A" & rs(1)).Checked = True
    lvPenjamin.ListItems("A" & rs(1)).ForeColor = vbBlue
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    Call LoadDataCombo(False)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub LoadDataCombo(chk As Boolean)
    On Error GoTo errLoad

strSQL = "SELECT distinct nik,nama from SlipInsentifPegawai "
Call msubRecFO(rs, strSQL)
lvPenjamin.ListItems.clear
If rs.RecordCount <> 0 Then
    For i = 0 To rs.RecordCount - 1
        lvPenjamin.ListItems.add , "A" & rs(0).Value, rs(1).Value
            If chk = True Then
                lvPenjamin.ListItems("A" & rs(0)).Checked = True
                lvPenjamin.ListItems("A" & rs(0)).ForeColor = vbBlue
            Else
                lvPenjamin.ListItems("A" & rs(0)).Checked = False
                lvPenjamin.ListItems("A" & rs(0)).ForeColor = vbBlack
            End If
        rs.MoveNext
    Next
    lvPenjamin.Sorted = True
End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub


