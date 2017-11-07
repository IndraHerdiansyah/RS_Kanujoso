VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatPotonganGaji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Potongan Gaji"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670
   Icon            =   "frmRiwayatPotonganGaji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8670
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   8655
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   6120
         Width           =   8415
         Begin VB.CommandButton cmdTutup 
            Caption         =   "Tutup"
            Height          =   375
            Left            =   6840
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdSimpan 
            Caption         =   "Simpan"
            Height          =   375
            Left            =   5400
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdHapus 
            Caption         =   "Hapus"
            Height          =   375
            Left            =   3960
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdBatal 
            Caption         =   "Batal"
            Height          =   375
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6015
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8415
         Begin VB.TextBox txtKodeKomponen 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   -120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtIdPegawai 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            TabIndex        =   20
            Top             =   -120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chkStatusAktif 
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6840
            TabIndex        =   19
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtNamaPegawai 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            TabIndex        =   18
            Top             =   480
            Width           =   2055
         End
         Begin MSDataGridLib.DataGrid dgPotonganGaji 
            Height          =   4215
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   7435
            _Version        =   393216
            HeadLines       =   2
            RowHeight       =   19
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtJumlah 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6720
            TabIndex        =   6
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtKeterangan 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   6495
         End
         Begin MSComCtl2.DTPicker dtpTglPotongan 
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   122093568
            UpDown          =   -1  'True
            CurrentDate     =   38448
         End
         Begin MSDataListLib.DataCombo dcKomponenPotongan 
            Height          =   315
            Left            =   4560
            TabIndex        =   4
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
         End
         Begin VB.Label Label5 
            Caption         =   "Keterangan"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Jumlah Potongan"
            Height          =   255
            Left            =   6720
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Komponen Potongan"
            Height          =   255
            Left            =   4560
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Nama Pegawai"
            Height          =   255
            Left            =   2400
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tanggal Potongan"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1455
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
      Left            =   6840
      Picture         =   "frmRiwayatPotonganGaji.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPotonganGaji.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatPotonganGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBatal_Click()
    dcKomponenPotongan.Text = ""
    txtJumlah.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo bawah
    Set rs = Nothing
    'strSQL = "delete from RiwayatPotonganGaji where IdPegawai = '" & mstrIdPegawai & "' "
    If MsgBox("Yakin akan menghapus data potongan gaji?", vbYesNo) = vbNo Then Exit Sub
    
    strSQL = "delete from RiwayatPotonganGaji where KdKomponenPotonganGaji = '" & dcKomponenPotongan.BoundText & "' "
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    MsgBox "Data berhasil dihapus", vbInformation
    
    Set rs = Nothing
    Call subLoadGrid
    Call cmdBatal_Click
bawah:
End Sub

Private Sub cmdTutup_Click()
    Call frmRiwayatPegawai.subLoadRiwayatPotongan
    Unload Me
End Sub

Private Sub dcKomponenPotongan_Change()
    txtKodeKomponen.Text = dcKomponenPotongan.BoundText
End Sub

Private Sub dcKomponenPotongan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKomponenPotongan.Text)) = 0 Then txtJumlah.SetFocus: Exit Sub
        If dcKomponenPotongan.MatchedWithList = True Then txtJumlah.SetFocus: Exit Sub
        strSQL = "SELECT KdKomponenPotonganGaji, KomponenPotonganGaji FROM KomponenPotonganGaji WHERE (KomponenPotonganGaji LIKE '%" & dcKomponenPotongan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKomponenPotongan.BoundText = rs(0).Value
        dcKomponenPotongan.Text = rs(1).Value
        txtKodeKomponen.Text = rs(0).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgPotonganGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtIdPegawai.Text = dgPotonganGaji.Columns("IdPegawai")
    dcKomponenPotongan.Text = dgPotonganGaji.Columns("KomponenPotonganGaji")
    txtJumlah.Text = dgPotonganGaji.Columns("JumlahPotongan")
    txtKeterangan.Text = dgPotonganGaji.Columns("Keterangan")
    chkStatusAktif.Value = dgPotonganGaji.Columns("statusEnabled")
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglPotongan.Value = Format(Now, "yyyy/MMMM/dd HH:mm:ss")
    txtIdPegawai.Text = frmRiwayatPegawai.txtIdPegawai.Text
    txtNamaPegawai.Text = frmRiwayatPegawai.txtNamaPegawai.Text
    Call subLoadDcSource
    Call subLoadGrid
End Sub

Public Sub subLoadDcSource()
    On Error GoTo Errload
    Call msubDcSource(dcKomponenPotongan, rs, "SELECT KdKomponenPotonganGaji, KomponenPotonganGaji FROM KomponenPotonganGaji where StatusEnabled = 1 ORDER BY KomponenPotonganGaji")
    If rs.EOF = False Then dcKomponenPotongan.BoundText = rs(0).Value
    Exit Sub
Errload:
End Sub

Public Sub subLoadGrid()
    On Error GoTo bawah
    Set rs = Nothing
    strSQL = "SELECT idPegawai, KomponenPotonganGaji, JumlahPotongan, Keterangan, statusEnabled FROM V_RiwayatPotonganGaji where idPegawai = '" & frmRiwayatPegawai.txtIdPegawai.Text & "'"
    Call msubRecFO(rs, strSQL)
    Set dgPotonganGaji.DataSource = rs
    Exit Sub
bawah:
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo bawah
    If dcKomponenPotongan.Text <> "" Then
        If Periksa("datacombo", dcKomponenPotongan, "Komponen Potongan Tidak Terdaftar") = False Then Exit Sub
    End If
    If txtJumlah.Text = "" Then
        If Periksa("text", txtJumlah, "Jumlah potongan harus di isi") = False Then Exit Sub
    End If
    
    Call sp_simpan("A")
    Call subLoadGrid
    Call cmdBatal_Click
bawah:
End Sub

Private Function sp_simpan(f_Status As String) As Boolean
    sp_simpan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIdPegawai.Text)
        If dcKomponenPotongan.BoundText = "" Then
        .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adChar, adParamInput, 2, Null)
        Else
'        .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adChar, adParamInput, 2, dcKomponenPotongan.BoundText)
         .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adChar, adParamInput, 2, txtKodeKomponen.Text)
        End If
        .Parameters.Append .CreateParameter("JumlahPotongan", adCurrency, adParamInput, , txtJumlah.Text)
        .Parameters.Append .CreateParameter("TglPotongan", adDate, adParamInput, , Format(dtpTglPotongan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 30, txtKeterangan.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatusAktif.Value)
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
    sp_simpan = False
End Function

Private Sub txtJumlah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
