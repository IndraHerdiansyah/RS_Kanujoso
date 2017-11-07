VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmStokBrg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Stok Barang Ruangan"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStokBrg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   12990
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   7560
      Width           =   12975
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   9360
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   7560
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   11160
         TabIndex        =   12
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Daftar Stok Barang Ruangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   12975
      Begin VB.TextBox txtCariAsalBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7185
         MaxLength       =   50
         TabIndex        =   9
         Top             =   5160
         Width           =   1680
      End
      Begin VB.TextBox txtCariJenisBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   7
         Top             =   5160
         Width           =   1680
      End
      Begin VB.TextBox txtCariBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4185
         MaxLength       =   50
         TabIndex        =   8
         Top             =   5160
         Width           =   1680
      End
      Begin VB.Frame frameDataBrg 
         Caption         =   "Data Barang"
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
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   8775
         Begin MSDataGridLib.DataGrid dgBarang 
            Height          =   1935
            Left            =   240
            TabIndex        =   1
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
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
      End
      Begin MSDataGridLib.DataGrid dgBrg 
         Height          =   4695
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8281
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Index           =   1
         Left            =   6120
         TabIndex        =   25
         Top             =   5205
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Barang"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   5205
         Width           =   1005
      End
      Begin VB.Label lblJmlData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Barang"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   11520
         TabIndex        =   23
         Top             =   5220
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Cari Barang"
         Height          =   210
         Index           =   6
         Left            =   3240
         TabIndex        =   22
         Top             =   5205
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   12975
      Begin VB.TextBox txtKdBarang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   27
         Text            =   "txtkdbarang"
         Top             =   60
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox txtLokasi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   11280
         MaxLength       =   12
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNamaBrg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   6855
      End
      Begin VB.TextBox txtJmlStok 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   10440
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtJmlMin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   9600
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcAsalBrg 
         Height          =   330
         Left            =   7200
         TabIndex        =   2
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lokasi"
         Height          =   210
         Left            =   11280
         TabIndex        =   21
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Left            =   7200
         TabIndex        =   19
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jml. Stok"
         Height          =   210
         Left            =   10440
         TabIndex        =   18
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jml. Min"
         Height          =   210
         Left            =   9600
         TabIndex        =   17
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   26
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
      Left            =   11160
      Picture         =   "frmStokBrg.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStokBrg.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStokBrg.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmStokBrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilter As String
Dim strkdbarang As String
Dim intJmlBarang As Integer
Dim kodebarang As String
Dim kodeasal As String
Dim tempbolTampil As Boolean
Dim i As Integer

Private Sub cmdBatal_Click()
    txtNamaBrg.Text = ""
    dcAsalBrg.Text = ""
    txtJmlMin.Text = ""
    txtJmlStok.Text = ""
    txtLokasi.Text = ""
    frameDataBrg.Visible = False
    txtCariBarang.Text = ""
    txtCariAsalBarang.Text = ""
    txtCariJenisBarang.Text = ""
    Call subLoadGridSource
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frm_cetak_stokBarang.Show
hell:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo a:
    Dim msg As String
    If dgBrg.Row = -1 Then Exit Sub
    If txtKdBarang = "" Then
        MsgBox "Pilih Dulu data yamg mau di hapus", vbInformation, "Informasi"
        Exit Sub
    End If
    msg = MsgBox("Apakah Benar Data akan di hapus", vbQuestion + vbYesNo, "Konfirmasi")
    If msg = vbYes Then
        strSQL = "delete StokRuangan where KdBarang='" & txtKdBarang & "' and KdAsal='" & dcAsalBrg.BoundText & "' and KdRuangan='" & mstrKdRuangan & "'"
        dbConn.Execute strSQL
        Call subLoadGridSource
    End If
    Exit Sub
a:
    MsgBox "Maaf Data tidak bisa di Hapus", vbCritical, "error"
End Sub

'SP AU_StockBarangRuangan sudah tidak ada jadi diubah menjadi AUD_StokRuangan
'dan penambahan parameter status
Private Sub cmdSimpan_Click()
    If txtKdBarang.Text = "" Then
        MsgBox "Nama Barang Harus dipilih", vbInformation, "Informasi"
        txtNamaBrg.SetFocus
        Exit Sub
    End If
    If dcAsalBrg.Text = "" Then
        MsgBox "Asal Barang Harus diisi", vbInformation, "Informasi"
        dcAsalBrg.SetFocus
        Exit Sub
    End If
    If txtNamaBrg.Text = "" Then
        MsgBox "Nama Barang Harus diisi", vbInformation, "Informasi"
        txtNamaBrg.SetFocus
        Exit Sub
    End If
    If txtJmlMin.Text = "" Then
        MsgBox "Jumlah Minimal Harus diisi", vbInformation, "Informasi"
        txtJmlMin.SetFocus
        Exit Sub
    End If
    If txtJmlStok.Text = "" Then
        MsgBox "Jumlah Stok Harus diisi", vbInformation, "Informasi"
        txtJmlStok.SetFocus
        Exit Sub
    End If
    If sp_StockBarang("A") = False Then Exit Sub

    txtJmlMin = ""
    txtJmlStok = ""
    dcAsalBrg.Text = ""
    txtNamaBrg.Text = ""
    txtLokasi.Text = ""
    frameDataBrg.Visible = False
    txtNamaBrg.SetFocus
    Call subLoadGridSource
End Sub

'Function untuk memanggil SP AUD_StokRuangan
Private Function sp_StockBarang(f_Status As String) As Boolean
    On Error GoTo errLoad

    sp_StockBarang = True
    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBrg.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("JmlMin", adInteger, adParamInput, , CInt(txtJmlMin))
        .Parameters.Append .CreateParameter("JmlStok", adDouble, adParamInput, , CDec(txtJmlStok))
        .Parameters.Append .CreateParameter("Lokasi", adVarChar, adParamInput, 12, IIf(txtLokasi.Text = "", Null, txtLokasi.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_StokRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_StockBarang = False
        End If
    End With

    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalBrg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtJmlMin.SetFocus
    End If
End Sub

Private Sub dgBarang_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgBarang
    WheelHook.WheelHook dgBarang
End Sub

Private Sub dgBarang_DblClick()
    Call dgBarang_KeyPress(13)
End Sub

Private Sub dgBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlBarang = 0 Then Exit Sub
        Dim strkd As String
        strkd = dgBarang.Columns(0).Value
        txtNamaBrg.Text = dgBarang.Columns(1).Value
        strkdbarang = strkd
        If strkdbarang = "" Then
            MsgBox "Pilih dulu Nama Barang yg ingin di proses", vbCritical, "Validasi"
            txtNamaBrg.Text = ""
            dgBarang.SetFocus
            Exit Sub
        End If
        frameDataBrg.Visible = False
    End If
    If KeyAscii = 27 Then
        frameDataBrg.Visible = False
    End If
End Sub

Private Sub dgBrg_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgBrg
    WheelHook.WheelHook dgBrg
End Sub

Private Sub dgBrg_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next

    With dgBrg
        If .ApproxCount = 0 Then Exit Sub
        txtKdBarang.Text = .Columns("KdBarang")
        txtNamaBrg.Text = .Columns("Nama Barang")
        dcAsalBrg.BoundText = .Columns("KdAsal")
        txtJmlStok.Text = .Columns("Jml. Stock")
    End With
    frameDataBrg.Visible = False

    lblJmlData.Caption = dgBrg.Bookmark & " / " & dgBrg.ApproxCount & " Data"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc("[") Or KeyAscii = Asc("]") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Set rs = Nothing
    rs.Open "select * from asalbarang where statusenabled<>0", dbConn, adOpenDynamic, adLockOptimistic
    Set dcAsalBrg.RowSource = rs
    dcAsalBrg.ListField = rs.Fields(1).Name
    dcAsalBrg.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
    Call subLoadGridSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtCariAsalBarang_Change()
    On Error GoTo errLoad
    Call subLoadGridSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtCariBarang_Change()
    On Error GoTo errLoad
    Call subLoadGridSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtCariJenisBarang_Change()
    On Error GoTo errLoad
    Call subLoadGridSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtJmlMin_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        txtJmlStok.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtJmlStok_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtLokasi.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtLokasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dgBarang.SetFocus
End Sub

Private Sub txtNamaBrg_Change()
    If tempbolTampil = True Then Exit Sub
    strFilter = "WHERE [nama barang] like '" & txtNamaBrg.Text & "%'"
    strkdbarang = ""
    frameDataBrg.Visible = True
End Sub

Private Sub subLoadbarang()
    On Error Resume Next
    strSQL = "SELECT * FROM V_Databarang " & strFilter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlBarang = rs.RecordCount
    Set dgBarang.DataSource = rs
    With dgBarang
        .Columns(0).Width = 1000
        .Columns(1).Width = 3500
        .Columns(2).Width = 2000
        .Columns(3).Width = 1000
        .Columns(4).Width = 0
        .Columns(5).Width = 0
        .Columns(6).Width = 0
        .Columns(7).Width = 0
    End With
End Sub

Sub subLoadGridSource()
    On Error GoTo errLoad

    strSQL = "SELECT NamaBarang as [Nama Barang],NamaAsal as Asal, JmlStok as [Jml. Stock],DetailJenisBarang as [Jenis Barang],KdBarang,KdAsal,NamaRuangan as Ruangan FROM V_InfoStokApotikNRuanganFIFO where KdRuangan='" & mstrKdRuangan & "' and (DetailJenisBarang like '%" & txtCariJenisBarang & "%' and NamaBarang like '%" & txtCariBarang & "%' and NamaAsal like '" & txtCariAsalBarang & "%')"

    Call msubRecFO(rs, strSQL)
    Set dgBrg.DataSource = rs
    With dgBrg
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("Nama Barang").Width = 3000
        .Columns("Asal").Width = 1000

        .Columns("Jml. Stock").Width = 1000
        .Columns("Jenis Barang").Width = 2000
    End With
    lblJmlData.Caption = 0 & " / " & dgBrg.ApproxCount & " Data"

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaBrg_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 27 Then
        txtNamaBrg = ""
        txtJmlStok = ""
        txtJmlMin = ""
        dcAsalBrg.Text = ""
        frameDataBrg.Visible = False
    End If
    If KeyAscii = 13 Then
        dgBarang.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
hell:
End Sub

