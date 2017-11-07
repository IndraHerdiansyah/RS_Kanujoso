VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPembayaranGajiPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pembayaran Gaji Pegawai"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9075
   Icon            =   "frmPembayaranGajiPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   9075
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   7800
      Width           =   9015
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutup"
         Height          =   375
         Left            =   6840
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   3960
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Batal"
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTabPembayaranGaji 
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pembayaran Gaji"
      TabPicture(0)   =   "frmPembayaranGajiPegawai.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Potongan Gaji"
      TabPicture(1)   =   "frmPembayaranGajiPegawai.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   6255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   8775
         Begin VB.Frame Frame4 
            Height          =   1695
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   8535
            Begin VB.TextBox txtJumlahPotongan 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   6600
               TabIndex        =   18
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtKeteranganPotongan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   17
               Top             =   1200
               Width           =   8295
            End
            Begin MSComCtl2.DTPicker dtpTglPotongan 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
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
               Format          =   130678787
               UpDown          =   -1  'True
               CurrentDate     =   37760
            End
            Begin MSDataListLib.DataCombo dcPegawai 
               Height          =   315
               Left            =   2280
               TabIndex        =   20
               Top             =   480
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcKomponenPotongan 
               Height          =   315
               Left            =   4560
               TabIndex        =   21
               Top             =   480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin VB.Label Label8 
               Caption         =   "Tgl Pembayaran"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label9 
               Caption         =   "Nama Pegawai"
               Height          =   255
               Left            =   2280
               TabIndex        =   25
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Jumlah Potongan"
               Height          =   255
               Left            =   6600
               TabIndex        =   24
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label13 
               Caption         =   "Komponen Potongan"
               Height          =   255
               Left            =   4560
               TabIndex        =   23
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label14 
               Caption         =   "Keterangan"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   960
               Width           =   1095
            End
         End
         Begin MSDataGridLib.DataGrid dgPotonganGaji 
            Height          =   4215
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   7435
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
               Locked          =   -1  'True
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   8775
         Begin VB.Frame Frame2 
            Height          =   1695
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   8535
            Begin VB.TextBox txtJumlah 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6600
               TabIndex        =   5
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtKeterangan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Top             =   1200
               Width           =   8295
            End
            Begin MSComCtl2.DTPicker dtpTglBayar 
               Height          =   330
               Left            =   120
               TabIndex        =   6
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
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
               Format          =   130678787
               UpDown          =   -1  'True
               CurrentDate     =   37760
            End
            Begin MSDataListLib.DataCombo dcNamaPegawai 
               Height          =   315
               Left            =   2280
               TabIndex        =   7
               Top             =   480
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcKomponenGaji 
               Height          =   315
               Left            =   4560
               TabIndex        =   8
               Top             =   480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin VB.Label Label1 
               Caption         =   "Tgl Pembayaran"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Nama Pegawai"
               Height          =   255
               Left            =   2280
               TabIndex        =   12
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Jumlah"
               Height          =   255
               Left            =   6600
               TabIndex        =   11
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label4 
               Caption         =   "KomponenGaji"
               Height          =   255
               Left            =   4560
               TabIndex        =   10
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label7 
               Caption         =   "Keterangan"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   960
               Width           =   1095
            End
         End
         Begin MSDataGridLib.DataGrid dgPembayaranGaji 
            Height          =   4095
            Left            =   120
            TabIndex        =   14
            Top             =   1920
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   7223
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
               Locked          =   -1  'True
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
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
      Left            =   7200
      Picture         =   "frmPembayaranGajiPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPembayaranGajiPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmPembayaranGajiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglBayar.Value = Format(Now, "yyyy/MMMM/dd")
    dtpTglPotongan.Value = Format(Now, "yyyy/MMMM/dd")
    Call subLoadGridSource
    Call subLoadDcSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    Call msubDcSource(dcNamaPegawai, rs, "SELECT IdPegawai, NamaLengkap FROM DataPegawai ORDER BY NamaLengkap")
    If rs.EOF = False Then dcNamaPegawai.BoundText = rs(0).Value
    Call msubDcSource(dcPegawai, rs, "SELECT IdPegawai, NamaLengkap FROM DataPegawai ORDER BY NamaLengkap")
    If rs.EOF = False Then dcPegawai.BoundText = rs(0).Value
    Call msubDcSource(dcKomponenGaji, rs, "SELECT KdKomponenGaji, KomponenGaji FROM KomponenGaji Where StatusEnabled = 1 ORDER BY KomponenGaji")
    If rs.EOF = False Then dcKomponenGaji.BoundText = rs(0).Value
    Call msubDcSource(dcKomponenPotongan, rs, "SELECT KdKomponenPotonganGaji, KomponenPotonganGaji FROM KomponenPotonganGaji where StatusEnabled = 1 ORDER BY KomponenPotonganGaji")
    If rs.EOF = False Then dcKomponenPotongan.BoundText = rs(0).Value
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgPotonganGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    dtpTglPotongan.Value = dgPotonganGaji.Columns("tglPembayaran")
    dcKomponenPotongan.Text = dgPotonganGaji.Columns("KomponenPotonganGaji")
    dcPegawai.Text = dgPotonganGaji.Columns("Namalengkap")
    txtJumlahPotongan.Text = dgPotonganGaji.Columns("Jumlah")
    txtKeteranganPotongan.Text = dgPotonganGaji.Columns("Keterangan")
End Sub

Public Sub subLoadGridSource()
    On Error GoTo bawah
    Select Case SSTabPembayaranGaji.Tab
        Case 0
            Set rs = Nothing
            strSQL = "SELECT*FROM V_PembayaranGajiPegawai"
            Call msubRecFO(rs, strSQL)
            Set dgPembayaranGaji.DataSource = rs
        Case 1
            Set rs = Nothing
            strSQL = "SELECT*FROM V_PembayaranPotonganGajiPegawai"
            Call msubRecFO(rs, strSQL)
            Set dgPotonganGaji.DataSource = rs
    End Select
    Exit Sub
bawah:
    Exit Sub
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo bawah
    Select Case SSTabPembayaranGaji.Tab
        Case 0
            If Periksa("text", dcNamaPegawai, "Silahkan Pilih Nama Pegawai") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "delete PembayaranGajiPegawai where IdPegawai = '" & dcNamaPegawai.BoundText & "' and kdKomponenGaji = '" & dcKomponenGaji.BoundText & "' and TglPembayaran='" & Format(dtpTglBayar.Value, "yyyy/MM/dd HH:mm:ss") & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
            Call cmdBatal_Click
            Call subLoadGridSource
            MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
        Case 1
            Set rs = Nothing
            strSQL = "delete PembayaranPotonganGajiPegawai where IdPegawai = '" & dcPegawai.BoundText & "' and kdKomponenPotonganGaji = '" & dcKomponenPotongan.BoundText & "'and TglPembayaran='" & Format(dtpTglPotongan.Value, "yyyy/MM/dd HH:mm:ss") & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
            Call cmdBatal_Click
            Call subLoadGridSource
            MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    End Select
    Exit Sub
bawah:
    Call msubPesanError
End Sub

Private Sub dcNamaPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKomponenGaji.SetFocus
End Sub

Private Sub dcKomponenGaji_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJumlah.SetFocus
End Sub

Private Sub SSTabPembayaranGaji_Click(PreviousTab As Integer)
    Call subLoadGridSource
End Sub

Private Sub txtJumlah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub dcPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKomponenPotongan.SetFocus
End Sub

Private Sub dcKomponenPotongan_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        strSQL = "SELECT JumlahPotongan FROM RiwayatPotonganGaji where idPegawai='" & dcPegawai.BoundText & "' and kdKomponenPotonganGaji = '" & dcKomponenPotongan.BoundText & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then txtJumlahPotongan.Text = rs(0).Value
        txtKeteranganPotongan.SetFocus
    End If
End Sub

Private Sub txtJumlahPotongan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeteranganPotongan.SetFocus
End Sub

Private Sub txtKeteranganPotongan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dgPembayaranGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    dtpTglBayar.Value = dgPembayaranGaji.Columns("tglPembayaran")
    dcKomponenGaji.Text = dgPembayaranGaji.Columns("KomponenGaji")
    dcNamaPegawai.Text = dgPembayaranGaji.Columns("namalengkap")
    txtJumlah.Text = dgPembayaranGaji.Columns("Jumlah")
    txtKeterangan.Text = dgPembayaranGaji.Columns("Keterangan")
End Sub

Private Sub cmdBatal_Click()
    dtpTglBayar.Value = Format(Now, "yyyy/MMMM/dd HH:mm:ss")
    dcNamaPegawai.BoundText = ""
    dcPegawai.BoundText = ""
    dcKomponenGaji.BoundText = ""
    dcKomponenPotongan.BoundText = ""
    txtJumlahPotongan.Text = ""
    txtJumlah.Text = ""
    txtKeterangan.Text = ""
    txtKeteranganPotongan.Text = ""
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    Select Case SSTabPembayaranGaji.Tab
        Case 0
            If Periksa("text", dcKomponenGaji, "Silahkan isi Nama Komponen Gaji") = False Then Exit Sub
            Call sp_SimpanPembayaranGajiPegawai("A")
            Call cmdBatal_Click
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
            Call subLoadGridSource
        Case 1
            If Periksa("text", dcKomponenPotongan, "Silahkan isi Nama Komponen Potongan Gaji ") = False Then Exit Sub
            Call sp_SimpanPembayaranPotonganGaji("A")
            Call cmdBatal_Click
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
            Call subLoadGridSource
    End Select
    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Function sp_SimpanPembayaranGajiPegawai(f_Status As String) As Boolean
    sp_SimpanPembayaranGajiPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglPembayaran", adDate, adParamInput, , Format(dtpTglBayar.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, dcNamaPegawai.BoundText)
        .Parameters.Append .CreateParameter("KdKomponenGaji", adVarChar, adParamInput, 50, dcKomponenGaji.BoundText)
        .Parameters.Append .CreateParameter("Jumlah", adCurrency, adParamInput, , txtJumlah.Text)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, txtKeterangan.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("kdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .ActiveConnection = dbConn
        .CommandText = "AUD_PembayaranGajiPegawai"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Daftar Layanan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AUD_PembayaranGajiPegawai")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    sp_SimpanPembayaranGajiPegawai = False
End Function

Private Function sp_SimpanRiwayatPotonganGaji(f_Status As String) As Boolean
    sp_SimpanRiwayatPotonganGaji = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, dcPegawai.BoundText)
        .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adVarChar, adParamInput, 50, dcKomponenPotongan.BoundText)
        .Parameters.Append .CreateParameter("JumlahPotongan", adCurrency, adParamInput, , txtJumlahPotongan.Text)
        .Parameters.Append .CreateParameter("TglPotongan", adDate, adParamInput, , Format(dtpTglPotongan.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, txtKeteranganPotongan.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .ActiveConnection = dbConn
        .CommandText = "AUD_RiwayatPotonganGaji"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Daftar Layanan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AUD_RiwayatPotonganGaji")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    sp_SimpanRiwayatPotonganGaji = False
End Function

Public Function sp_SimpanPembayaranPotonganGaji(f_Status As String) As Boolean
    sp_SimpanPembayaranPotonganGaji = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglPembayaran", adDate, adParamInput, , Format(dtpTglPotongan, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, dcPegawai.BoundText)
        .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adChar, adParamInput, 2, dcKomponenPotongan.BoundText)
        .Parameters.Append .CreateParameter("Jumlah", adCurrency, adParamInput, , txtJumlahPotongan.Text)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 30, txtKeteranganPotongan.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("kdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .ActiveConnection = dbConn
        .CommandText = "AUD_PembayaranPotonganGajiPegawai"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Daftar Layanan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AUD_PembayaranPotonganGajiPegawai")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    sp_SimpanPembayaranPotonganGaji = False
End Function

Private Sub cmdTutupRiwayat_Click()
    FrRiwayat.Visible = False
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
