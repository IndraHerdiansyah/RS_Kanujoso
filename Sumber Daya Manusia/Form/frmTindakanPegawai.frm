VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTindakanPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pelayanan Tindakan Pegawai"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTindakanPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8430
   Begin VB.Frame fraPelayanan 
      Caption         =   "Daftar Pelayanan Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      Top             =   7800
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid dgPelayanan 
         Height          =   2415
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
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
   End
   Begin VB.Frame fraPegawai 
      Caption         =   "Daftar Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid dgPegawai 
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Daftar Pelayanan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   14
      Top             =   3120
      Width           =   8415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPelayanan 
         Height          =   1575
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   65535
         BackColorSel    =   8388608
         BackColorBkg    =   16777215
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame fraButton 
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   5280
      Width           =   8415
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   360
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Tutu&p"
         Height          =   360
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraPPelayanan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   2160
      Width           =   8415
      Begin VB.TextBox txtKuantitas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtNamaPelayanan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4695
      End
      Begin VB.CheckBox chkAPBD 
         Caption         =   "Pos APBD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   518
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   360
         Left            =   5880
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   360
         Left            =   7080
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   240
         Left            =   5040
         TabIndex        =   19
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame fraPDokter 
      Height          =   1095
      Left            =   0
      TabIndex        =   20
      Top             =   1080
      Width           =   8415
      Begin VB.TextBox txtKodeJabatan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         MaxLength       =   5
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         TabIndex        =   24
         Top             =   525
         Width           =   3135
      End
      Begin VB.TextBox txtPegawai 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   525
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   525
         Width           =   1935
         _ExtentX        =   3413
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   115015683
         UpDown          =   -1  'True
         CurrentDate     =   37823
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan Fungsional"
         Height          =   240
         Index           =   3
         Left            =   5040
         TabIndex        =   25
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai"
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   23
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pelayanan"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1620
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgPerawatPerPelayanan 
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTindakanPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6600
      Picture         =   "frmTindakanPegawai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTindakanPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmTindakanPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterPelayanan As String
Dim strKodePelayanan As String
Dim intJmlPelayanan As Integer
Dim mstrKdPegawai As String
Dim subJmlTotal As Integer
Dim strPilihGrid As String
Dim mintJmlPegawai As Integer
Dim mstrFilterPegawai As String
Dim intRowNow As Integer

Private Sub cmdBatal_Click()
    frmDaftarPelayananPegawai.Enabled = True
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    Dim h As Integer
    With fgPelayanan
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        Call msubRemoveItem(fgPelayanan, .Row)
    End With
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim tempKdRuanganAsal As String
    If txtPegawai.Text = "" Then
        MsgBox "Nama Pegawai kosong, silahkan lengkapi terlebih dahulu", vbCritical, "Validasi"
        Exit Sub
    End If

    If funcCekValidasi = False Then Exit Sub
    Call subEnableButtonReg(False)

    For i = 1 To fgPelayanan.Rows - 2
        'simpan biaya pelayanan
        If sp_PenilaianPegawai(dbcmd, fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 2), dtpTglPeriksa) = False Then Exit Sub

    Next i

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTambah_Click()
    Dim i As Integer
    Dim adocmd As New ADODB.Command
    On Error GoTo hell

    If strKodePelayanan = "" Then Exit Sub

    With fgPelayanan
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, 0) = strKodePelayanan) And _
                (.TextMatrix(i, 4) = dtpTglPeriksa.Value) Then txtNamaPelayanan.SetFocus: txtNamaPelayanan.SelStart = 0: txtNamaPelayanan.SelLength = Len(txtNamaPelayanan.Text): Exit Sub
            Next i
            intRowNow = .Rows - 1
            .TextMatrix(intRowNow, 0) = strKodePelayanan
            .TextMatrix(intRowNow, 1) = txtNamaPelayanan.Text
            .TextMatrix(intRowNow, 2) = CInt(txtKuantitas.Text)
            .TextMatrix(intRowNow, 3) = mstrKdPegawai

            .Rows = .Rows + 1
            .SetFocus
        End With

        txtNamaPelayanan.Text = ""
        txtKuantitas.Text = 1
        fraPelayanan.Visible = False

        Exit Sub
hell:
        Call msubPesanError
End Sub

Private Sub dgPegawai_DblClick()
    Call dgPegawai_KeyPress(13)
End Sub

Private Sub dgPegawai_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        If mintJmlPegawai = 0 Then Exit Sub
        txtKodeJabatan.Text = ""
        txtJabatan.Text = ""
        txtPegawai.Text = dgPegawai.Columns(1).Value
        txtJabatan.Text = dgPegawai.Columns(3).Value
        mstrKdPegawai = dgPegawai.Columns(0).Value

        txtKodeJabatan.Text = dgPegawai.Columns(4).Value
        If mstrKdPegawai = "" Then
            MsgBox "Silahkan pilih Nama Pegawai", vbCritical, "Validasi"
            txtPegawai.Text = ""
            dgPegawai.SetFocus
            Exit Sub
        End If
        If txtKodeJabatan.Text = "" Then
            MsgBox "Jabatan Fungsional pegawai kosong, silahkan lengkapi terlebih dahulu", vbCritical, "Validasi"
            fraPegawai.Visible = False
            Exit Sub
        End If
        fraPegawai.Visible = False
        txtNamaPelayanan.SetFocus
    End If

    If KeyAscii = 27 Then
        fraPegawai.Visible = False
    End If

End Sub

Private Sub dgPelayanan_DblClick()
    Call dgPelayanan_KeyPress(13)
End Sub

Private Sub dgPelayanan_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        Dim strkd As String
        strkd = dgPelayanan.Columns(0).Value
        txtNamaPelayanan.Text = dgPelayanan.Columns(1).Value
        strKodePelayanan = strkd

        If strKodePelayanan = "" Then
            MsgBox "Pilih dulu tindakan pelayanan Pasien", vbCritical, "Validasi"
            txtNamaPelayanan.Text = ""
            dgPelayanan.SetFocus
            Exit Sub
        End If

        fraPelayanan.Visible = False

    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dtpTglPeriksa_Change()
    dtpTglPeriksa.MaxDate = Now
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPegawai.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subSetGidPelayanan
    dtpTglPeriksa.Value = Now

    subJmlTotal = 0

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarPelayananPegawai.Enabled = True
End Sub

Private Sub txtPegawai_Change()
    mstrFilterPegawai = " WHERE [Nama Lengkap] like '%" & txtPegawai.Text & "%'"
    fraPegawai.Visible = True
    Call subLoadDokter
End Sub

Private Sub txtPegawai_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If mintJmlPegawai = 0 Then Exit Sub
        If fraPegawai.Visible = True Then
            dgPegawai.SetFocus

        End If
        txtNamaPelayanan.SetFocus
    End If
    If KeyAscii = 27 Then
        fraPegawai.Visible = False
    End If
    Exit Sub
hell:
End Sub

Private Sub txtKuantitas_Change()
    On Error Resume Next
    If txtKuantitas.Text = "" Or txtKuantitas.Text = 0 Then txtKuantitas.Text = 1
End Sub

Private Sub txtKuantitas_GotFocus()
    txtKuantitas.SelStart = 0
    txtKuantitas.SelLength = Len(txtKuantitas.Text)
End Sub

Private Sub txtKuantitas_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)

End Sub

Private Sub txtKuantitas_LostFocus()
    If txtKuantitas.Text = "" Then txtKuantitas.Text = 1: Exit Sub
    If txtKuantitas.Text = 0 Then txtKuantitas.Text = 1
End Sub

Private Sub txtNamaPelayanan_Change()
    strFilterPelayanan = "WHERE [Nama Pelayanan] like '%" & txtNamaPelayanan.Text _
    & "%' AND [Kode Jabatan]='" & txtKodeJabatan.Text & "' "
    strKodePelayanan = ""
    fraPelayanan.Visible = True
    Call subLoadPelayanan
End Sub

Private Sub txtNamaPelayanan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        If fraPelayanan.Visible = True Then
            dgPelayanan.SetFocus
        Else
            txtKuantitas.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
hell:
End Sub

'untuk load data pegawai
Private Sub subLoadDokter()
    On Error GoTo errLoad
    strSQL = "SELECT [Id Pegawai], [Nama Lengkap],JK,[Jabatan Fungsi],KdDetailJabatanF FROM V_M_DataPegawai" & mstrFilterPegawai
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlPegawai = rs.RecordCount
    With dgPegawai
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
        .Columns(4).Width = 0
    End With
    fraPegawai.Left = 0
    fraPegawai.Top = 1920
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'untuk meload data pelayanan di grid
Private Sub subLoadPelayanan()
    On Error GoTo errLoad
    strSQL = "SELECT * FROM V_LoadPelayananFungsional " & strFilterPelayanan
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlPelayanan = rs.RecordCount
    With dgPelayanan
        Set .DataSource = rs
        .Columns(0).Width = 0
        .Columns(1).Width = 4000
        .Columns(2).Width = 0
        .Columns(3).Width = 2000
        .Columns(4).Width = 0
        .Columns(5).Width = 0
        .Columns(6).Width = 1000
        .Columns(6).Alignment = dbgRight
    End With
    fraPelayanan.Left = 0
    fraPelayanan.Top = 3120
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_PenilaianPegawai(ByVal adoCommand As ADODB.Command, strKdPelayanan As String, intJmlPel As Integer, dtTanggalPelayanan As Date) As Boolean
    On Error GoTo errLoad
    sp_PenilaianPegawai = True
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrKdPegawai)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananF", adChar, adParamInput, 6, strKdPelayanan)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , intJmlPel)
        .Parameters.Append .CreateParameter("TglPelayananF", adDate, adParamInput, , Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "Add_DetailPenilaianPegawai"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            sp_PenilaianPegawai = False
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

    Exit Function
errLoad:
    sp_PenilaianPegawai = False
    Call msubPesanError
End Function

'simpan data perawat
Private Function sp_PetugasPemeriksaBP(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, F_IdUser As String) As Boolean
    On Error GoTo errLoad

    sp_PetugasPemeriksaBP = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, F_IdUser)

        .ActiveConnection = dbConn
        .CommandText = "Add_PetugasPemeriksaBP"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas pemeriksa BP", vbExclamation, "Validasi"
            sp_PetugasPemeriksaBP = False
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_PetugasPemeriksaBP = False
    Call msubPesanError
End Function

'untuk set grid pelayanan
Private Sub subSetGidPelayanan()
    With fgPelayanan
        .clear
        .Rows = 2
        .Cols = 5
        .TextMatrix(0, 0) = "Kode Pelayanan"
        .TextMatrix(0, 1) = "Nama Pelayanan"
        .TextMatrix(0, 2) = "Jumlah"
        .TextMatrix(0, 3) = "Id Pegawai"
        .TextMatrix(0, 4) = "Tgl Pelayanan"

        .ColWidth(0) = 0
        .ColWidth(1) = 6000
        .ColWidth(2) = 700
        .ColWidth(3) = 0
        .ColWidth(4) = 0

    End With
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
    If fgPelayanan.TextMatrix(1, 0) = "" Then
        MsgBox "Pilihan Pelayanan harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNamaPelayanan.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    fraPDokter.Enabled = blnStatus
    fraPPelayanan.Enabled = blnStatus
    fgPelayanan.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

Private Sub subSetGridPerawatPerPelayanan()
    With fgPerawatPerPelayanan
        .Cols = 6
        .Rows = 1

        .MergeCells = flexMergeFree

        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"

    End With
End Sub

