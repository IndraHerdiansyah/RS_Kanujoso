VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmKinerja_Transaksi 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kinerja Pegawai"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKinerja_Transaksi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   17010
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   14280
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   11760
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   13080
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   15600
      TabIndex        =   4
      Top             =   7680
      Width           =   1215
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
   Begin VB.Frame frmSasaran 
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   16695
      Begin VB.ListBox lstPegawai 
         Appearance      =   0  'Flat
         Height          =   3600
         Left            =   3240
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5055
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   5160
         TabIndex        =   19
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIdPegawai 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtNmPegawai 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3240
         TabIndex        =   14
         Top             =   720
         Width           =   5055
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   5055
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid fgdata2 
         Height          =   5055
         Left            =   8400
         TabIndex        =   8
         Top             =   1200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   104660995
         CurrentDate     =   42325
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM-yyyy"
         Format          =   104660995
         CurrentDate     =   42325
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Pegawai"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sasaran"
         Height          =   375
         Left            =   14640
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame frmPenilaian 
      Height          =   6495
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   16695
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5055
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   5055
         Left            =   8400
         TabIndex        =   12
         Top             =   1200
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Penilaian"
         Height          =   375
         Left            =   15120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Image Image4 
      Height          =   945
      Left            =   15120
      Picture         =   "frmKinerja_Transaksi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKinerja_Transaksi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16455
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKinerja_Transaksi.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmKinerja_Transaksi.frx":5A71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmKinerja_Transaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BHU As Boolean
Public statusForm As String
Public kode As String

Private Sub cmdCancel_Click()
    frmKinerja_Transaksi.fgData.Rows = 1
    frmKinerja_Transaksi.fgData.Cols = 1
    frmKinerja_Transaksi.fgdata2.Rows = 1
    frmKinerja_Transaksi.fgdata2.Cols = 1
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.Cols = 1
End Sub

Private Sub cmdCetak_Click()
'    FrmCetakNilaiKinerja.Show
    Dim a, d As Integer
    Dim b, c, e, f As String
    kode = Format(Now(), "yyMMddHHmm")
    'tempCetakKinerjaBulan
    strSQL = "SELECT     MasterKinerja.KdKategoryKinerja, MasterKinerja.NamaKinerja, NilaiKinerjaPegawai.Nilai , (SasaranKinerjaPegawai.Nilai / 12)  AS nilai, " & _
             "NilaiKinerjaPegawai.idpegawai , NilaiKinerjaPegawai.bulan, NilaiKinerjaPegawai.tahun " & _
             "FROM         NilaiKinerjaPegawai INNER JOIN MasterKinerja ON NilaiKinerjaPegawai.KdKinerja = MasterKinerja.KdKinerja INNER JOIN " & _
             "SasaranKinerjaPegawai ON NilaiKinerjaPegawai.KdKinerja = SasaranKinerjaPegawai.KdKinerja AND " & _
             "NilaiKinerjaPegawai.idpegawai = SasaranKinerjaPegawai.idpegawai And NilaiKinerjaPegawai.tahun = SasaranKinerjaPegawai.tahun " & _
             "WHERE     (NilaiKinerjaPegawai.IdPegawai = '" & txtIdPegawai.Text & "' and  NilaiKinerjaPegawai.bulan ='" & Format(DTPicker2.Value, "MM") & "' and NilaiKinerjaPegawai.tahun='" & Format(DTPicker2.Value, "yyyy") & "')"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        MSFlexGrid1.Cols = 7
        MSFlexGrid1.Rows = rs.RecordCount + 1
        For i = 0 To rs.RecordCount - 1
'            MSFlexGrid1.TextMatrix(i, 0) = ""
'            MSFlexGrid1.TextMatrix(i, 1) = ""
'            MSFlexGrid1.TextMatrix(i, 2) = ""
'            MSFlexGrid1.TextMatrix(i, 3) = ""
'            MSFlexGrid1.TextMatrix(i, 4) = ""
'            MSFlexGrid1.TextMatrix(i, 5) = ""

            If rs(0) = "01" Then
                a = a + 1
                MSFlexGrid1.TextMatrix(a, 0) = a
                MSFlexGrid1.TextMatrix(a, 1) = rs(1)
                MSFlexGrid1.TextMatrix(a, 2) = FormatNumber((CDbl(rs(2)) / CDbl(rs(3))) * 100, 0)
            End If
            If rs(0) = "02" Then
                d = d + 1
                MSFlexGrid1.TextMatrix(d, 3) = d
                MSFlexGrid1.TextMatrix(d, 4) = rs(1)
                MSFlexGrid1.TextMatrix(d, 5) = FormatNumber((CDbl(rs(2)) / CDbl(rs(3))) * 100, 0)
            End If
            rs.MoveNext
        Next
        
        strSQL = "Delete from tempCetakKinerjaBulan where kode ='" & txtIdPegawai.Text & "'"
        Call msubRecFO(rs, strSQL)
        For i = 1 To MSFlexGrid1.Rows - 1
            If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
                strSQL = "insert into tempCetakKinerjaBulan values ('" & MSFlexGrid1.TextMatrix(i, 0) & "','" & MSFlexGrid1.TextMatrix(i, 1) & "','" & MSFlexGrid1.TextMatrix(i, 2) & "','" & MSFlexGrid1.TextMatrix(i, 3) & "','" & MSFlexGrid1.TextMatrix(i, 4) & "','" & MSFlexGrid1.TextMatrix(i, 5) & "','" & txtIdPegawai.Text & "')"
                Call msubRecFO(rs, strSQL)
            End If
        Next
        
    End If
    
    strSQL = "select * from tempCetakKinerjaBulan where kode ='" & txtIdPegawai.Text & "'"
    vLaporan = "View"
    FrmCetakNilaiKinerja.Show
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If frmSasaran.Visible = True Then
        If statusForm = "SASARAN" Then
            'SASARAN
            strSQL = "select * from sasaranKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and tahun='" & Format(DTPicker1.Value, "yyyy") & "'"
            Call msubRecFO(rs, strSQL)
            If rs.RecordCount <> 0 Then
                strSQL = "delete from sasaranKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and tahun='" & Format(DTPicker1.Value, "yyyy") & "'"
                Call msubRecFO(rs, strSQL)
            End If
            For i = 1 To fgData.Rows - 1
                If fgData.TextMatrix(i, 1) <> "Total Bobot BHU" Then
                    strSQL = "insert into sasaranKinerjaPegawai values ('" & fgData.TextMatrix(i, 4) & "','" & txtIdPegawai.Text & "','" & Format(DTPicker1.Value, "yyyy") & "','" & fgData.TextMatrix(i, 2) & "')"
                    Call msubRecFO(rs, strSQL)
                End If
            Next
            For i = 1 To fgdata2.Rows - 1
                If fgdata2.TextMatrix(i, 1) <> "Total Bobot BPU" Then
                    strSQL = "insert into sasaranKinerjaPegawai values ('" & fgdata2.TextMatrix(i, 4) & "','" & txtIdPegawai.Text & "','" & Format(DTPicker1.Value, "yyyy") & "','" & fgdata2.TextMatrix(i, 2) & "')"
                    Call msubRecFO(rs, strSQL)
                End If
            Next
        Else
            'PENILAIAN
            strSQL = "select * from NilaiKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and tahun='" & Format(DTPicker2.Value, "yyyy") & "'  and Bulan='" & Format(DTPicker2.Value, "MM") & "'"
            Call msubRecFO(rs, strSQL)
            If rs.RecordCount <> 0 Then
                strSQL = "delete from NilaiKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and tahun='" & Format(DTPicker2.Value, "yyyy") & "'  and Bulan='" & Format(DTPicker2.Value, "MM") & "'"
                Call msubRecFO(rs, strSQL)
            End If
            For i = 1 To fgData.Rows - 1
                strSQL = "insert into NilaiKinerjaPegawai values ('" & fgData.TextMatrix(i, 4) & "','" & txtIdPegawai.Text & "','" & Format(DTPicker2.Value, "MM") & "','" & fgData.TextMatrix(i, 2) & "','" & Format(DTPicker2.Value, "yyyy") & "')"
                Call msubRecFO(rs, strSQL)
            Next
            For i = 1 To fgdata2.Rows - 1
                strSQL = "insert into NilaiKinerjaPegawai values ('" & fgdata2.TextMatrix(i, 4) & "','" & txtIdPegawai.Text & "','" & Format(DTPicker2.Value, "MM") & "','" & fgdata2.TextMatrix(i, 2) & "','" & Format(DTPicker2.Value, "yyyy") & "')"
                Call msubRecFO(rs, strSQL)
            Next
        End If
    Else
        'PENILAIAN
        
    End If
End Sub

Private Sub DTPicker2_Change()
    If txtIdPegawai.Text <> "" Then
        Call txtIdPegawai_KeyDown(13, False)
    End If
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then
        
    'End If
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    If fgData.Col = 2 Then
        txtIsi.Left = fgData.ColWidth(1) + fgData.ColWidth(0) + fgData.Left
        txtIsi.Top = (fgData.RowHeight(1) * (fgData.row)) + fgData.Top
        txtIsi.Width = fgData.ColWidth(2)
        txtIsi.Height = fgData.RowHeight(2)
        txtIsi.Visible = True
        txtIsi.SelStart = Len(txtIsi.Text)
        txtIsi.SetFocus
        txtIsi.Text = Chr(KeyAscii)
        txtIsi.SelStart = Len(txtIsi.Text)
        BHU = True
    End If
End Sub

Private Sub fgdata2_KeyPress(KeyAscii As Integer)
    If fgdata2.Col = 2 Then
        txtIsi.Left = fgdata2.ColWidth(1) + fgdata2.ColWidth(0) + fgdata2.Left
        txtIsi.Top = (fgdata2.RowHeight(1) * (fgdata2.row)) + fgdata2.Top
        txtIsi.Width = fgdata2.ColWidth(2)
        txtIsi.Height = fgdata2.RowHeight(2)
        txtIsi.Visible = True
        txtIsi.SelStart = Len(txtIsi.Text)
        txtIsi.SetFocus
        txtIsi.Text = Chr(KeyAscii)
        txtIsi.SelStart = Len(txtIsi.Text)
        BHU = False
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
'    If statusForm = "SASARAN" Then
'        Label3.Caption = "Tahun"
'        DTPicker2.Visible = False
'        DTPicker1.Visible = True
'    Else
'        Label3.Caption = "Bulan"
'        DTPicker2.Visible = True
'        DTPicker1.Visible = False
'    End If
End Sub

Private Sub frmSasaran_DblClick()
    MSFlexGrid1.Visible = Not MSFlexGrid1.Visible
End Sub

Private Sub List1_Click()

End Sub

Private Sub txtIdPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdCancel_Click
    Call LoadKinerja
End If
End Sub
Private Sub SetGrid()
    fgData.Cols = 5
    fgData.TextMatrix(0, 0) = "No"
    fgData.TextMatrix(0, 1) = "Materi Kinerja BHU 70%"
    fgData.TextMatrix(0, 2) = "Ukuran Pencapaian"
    fgData.TextMatrix(0, 3) = "Bobot"
    
    fgData.ColWidth(0) = 600
    fgData.ColWidth(1) = 5000
    fgData.ColWidth(2) = 1600
    If statusForm = "SASARAN" Then
        fgData.ColWidth(3) = 700
    Else
        fgData.ColWidth(3) = 0
    End If
    fgData.ColWidth(4) = 0
    
    fgdata2.Cols = 5
    fgdata2.TextMatrix(0, 0) = "No"
    fgdata2.TextMatrix(0, 1) = "Materi Kinerja BPU 30%"
    fgdata2.TextMatrix(0, 2) = "Ukuran Pencapaian"
    fgdata2.TextMatrix(0, 3) = "Bobot"
    
    fgdata2.ColWidth(0) = 600
    fgdata2.ColWidth(1) = 5000
    fgdata2.ColWidth(2) = 1600
    If statusForm = "SASARAN" Then
        fgdata2.ColWidth(3) = 700
    Else
        fgdata2.ColWidth(3) = 0
    End If
    fgdata2.ColWidth(4) = 0
    
End Sub
Private Sub LoadKinerja()
    If frmSasaran.Visible = True Then
        'SASARAN
        
        strSQL = "select * from DataPegawai where idpegawai='" & txtIdPegawai.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            txtIdPegawai.Text = rs(0)
            txtNmPegawai.Text = rs(3)
        End If
        
        Call SetGrid
        
        strSQL = "SELECT KinerjaPegawai.KdKinerja, KinerjaPegawai.idPegawai, MasterKinerja.NamaKinerja, MasterKinerja.KdKategoryKinerja, MasterKinerja.StatusEnabled " & _
                 "FROM KinerjaPegawai INNER JOIN MasterKinerja ON KinerjaPegawai.KdKinerja = MasterKinerja.KdKinerja " & _
                 "where KinerjaPegawai.idPegawai='" & txtIdPegawai.Text & "'"
        'strSQL = "SELECT     KinerjaPegawai.KdKinerja, KinerjaPegawai.idPegawai, MasterKinerja.NamaKinerja, MasterKinerja.KdKategoryKinerja, MasterKinerja.StatusEnabled, " & _
                 "SasaranKinerjaPegawai.tahun , SasaranKinerjaPegawai.Nilai " & _
                 "FROM         KinerjaPegawai INNER JOIN " & _
                 "MasterKinerja ON KinerjaPegawai.KdKinerja = MasterKinerja.KdKinerja  LEFT OUTER JOIN " & _
                 "SasaranKinerjaPegawai ON KinerjaPegawai.KdKinerja = SasaranKinerjaPegawai.KdKinerja AND " & _
                 "KinerjaPegawai.idpegawai = SasaranKinerjaPegawai.idpegawai " & _
                 "WHERE     (KinerjaPegawai.idPegawai = '" & txtIdPegawai.Text & "' and SasaranKinerjaPegawai.tahun='" & Format(DTPicker1.Value, "yyyy") & "')"
        Call msubRecFO(rs, strSQL)
        fgData.Rows = 1
        fgdata2.Rows = 1
        If rs.RecordCount <> 0 Then
            For i = 0 To rs.RecordCount - 1
                If rs(3) = "01" Then 'BHU
                    fgData.Rows = fgData.Rows + 1
                    fgData.TextMatrix(fgData.Rows - 1, 0) = fgData.Rows - 1
                    fgData.TextMatrix(fgData.Rows - 1, 1) = rs(2)
                    If statusForm = "SASARAN" Then
                        strSQLIdentifikasi = "select nilai from sasaranKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and KdKinerja='" & rs(0) & "' and tahun='" & Format(DTPicker1.Value, "yyyy") & "'"
                        Call msubRecFO(rsAplikasi, strSQLIdentifikasi)
                        If rsAplikasi.RecordCount <> 0 Then
                            fgData.TextMatrix(fgData.Rows - 1, 2) = rsAplikasi(0)
                        Else
                            fgData.TextMatrix(fgData.Rows - 1, 2) = ""
                        End If
                    Else
                        strSQLIdentifikasi = "select nilai from NilaiKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and KdKinerja='" & rs(0) & "' and tahun='" & Format(DTPicker1.Value, "yyyy") & "'  and Bulan='" & Format(DTPicker2.Value, "MM") & "'"
                        Call msubRecFO(rsAplikasi, strSQLIdentifikasi)
                        If rsAplikasi.RecordCount <> 0 Then
                            fgData.TextMatrix(fgData.Rows - 1, 2) = rsAplikasi(0)
                        Else
                            fgData.TextMatrix(fgData.Rows - 1, 2) = ""
                        End If
                    End If
                    fgData.TextMatrix(fgData.Rows - 1, 4) = rs(0)
                Else 'BPU
                    fgdata2.Rows = fgdata2.Rows + 1
                    fgdata2.TextMatrix(fgdata2.Rows - 1, 0) = fgdata2.Rows - 1
                    fgdata2.TextMatrix(fgdata2.Rows - 1, 1) = rs(2)
                    If statusForm = "SASARAN" Then
                        strSQLIdentifikasi = "select nilai from sasaranKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and KdKinerja='" & rs(0) & "' and tahun='" & Format(DTPicker1.Value, "yyyy") & "'"
                        Call msubRecFO(rsAplikasi, strSQLIdentifikasi)
                        If rsAplikasi.RecordCount <> 0 Then
                            fgdata2.TextMatrix(fgdata2.Rows - 1, 2) = rsAplikasi(0)
                        Else
                            fgdata2.TextMatrix(fgdata2.Rows - 1, 2) = ""
                        End If
                    Else
                        strSQLIdentifikasi = "select nilai from NilaiKinerjaPegawai where IdPegawai='" & txtIdPegawai.Text & "' and KdKinerja='" & rs(0) & "' and tahun='" & Format(DTPicker2.Value, "yyyy") & "' and Bulan='" & Format(DTPicker2.Value, "MM") & "'"
                        Call msubRecFO(rsAplikasi, strSQLIdentifikasi)
                        If rsAplikasi.RecordCount <> 0 Then
                            fgdata2.TextMatrix(fgdata2.Rows - 1, 2) = rsAplikasi(0)
                        Else
                            fgdata2.TextMatrix(fgdata2.Rows - 1, 2) = ""
                        End If
                    End If
                    'fgdata2.TextMatrix(fgdata2.Rows - 1, 2) = ""
                    fgdata2.TextMatrix(fgdata2.Rows - 1, 4) = rs(0)
                End If
                rs.MoveNext
            Next
            
        End If
        
        If fgData.Rows > 1 And statusForm = "SASARAN" Then
            Dim bobot As Double
            Dim b1, b2 As Double
            bobot = FormatNumber(70 / (fgData.Rows - 1), 2)
            For i = 1 To fgData.Rows - 1
                fgData.TextMatrix(i, 3) = CStr(bobot) & " %"
                b1 = b1 + bobot
            Next
            bobot = FormatNumber(30 / (fgdata2.Rows - 1), 2)
            For i = 1 To fgdata2.Rows - 1
                fgdata2.TextMatrix(i, 3) = CStr(bobot) & " %"
                b2 = b2 + bobot
            Next
            fgData.Rows = fgData.Rows + 1
            fgData.TextMatrix(fgData.Rows - 1, 1) = "Total Bobot BHU"
            fgData.TextMatrix(fgData.Rows - 1, 3) = CStr(FormatNumber(b1, 0)) & " %"
            fgdata2.Rows = fgdata2.Rows + 1
            fgdata2.TextMatrix(fgdata2.Rows - 1, 1) = "Total Bobot BPU"
            fgdata2.TextMatrix(fgdata2.Rows - 1, 3) = CStr(FormatNumber(b2, 0)) & " %"
        End If
    Else
        'PENILAIAN
        
    End If
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txtIsi.Visible = False
    If KeyCode = 13 Then
        If BHU = True Then
            fgData.TextMatrix(fgData.row, 2) = txtIsi.Text
            txtIsi.Visible = False
            If fgData.row + 1 < fgData.Rows Then
                fgData.row = fgData.row + 1
                fgData.SetFocus
            End If
        Else
            fgdata2.TextMatrix(fgdata2.row, 2) = txtIsi.Text
            txtIsi.Visible = False
            If fgdata2.row + 1 < fgdata2.Rows Then
                fgdata2.row = fgdata2.row + 1
                fgdata2.SetFocus
            End If
        End If
    End If
End Sub

Private Sub lstPegawai_DblClick()
    lstPegawai.Visible = False
    strSQL = "select * from datapegawai where namalengkap ='" & lstPegawai.List(lstPegawai.ListIndex) & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        txtNmPegawai.Text = rs(3)
        txtIdPegawai.Text = rs(0)
        txtNmPegawai.SetFocus
        Call txtIdPegawai_KeyDown(13, False)
    End If
End Sub

Private Sub lstPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call lstPegawai_DblClick
End Sub

Private Sub txtNmPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 And lstPegawai.Visible = True Then lstPegawai.SetFocus
    If KeyCode = 13 Then
        If txtNmPegawai.Text <> "" Then
        strSQL = "select * from datapegawai where NamaLengkap like '%" & txtNmPegawai.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            lstPegawai.Visible = True
            lstPegawai.clear
            For i = 0 To rs.RecordCount - 1
                lstPegawai.AddItem rs(3), i
                rs.MoveNext
            Next
            
        End If
    End If
    End If
End Sub
