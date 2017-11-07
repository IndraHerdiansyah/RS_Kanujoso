VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmMonitoringStokBarangNonMedis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Monitoring Stok Barang Non Medis"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMonitoringStokBarangNonMedis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   12795
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
      Left            =   120
      TabIndex        =   10
      Top             =   7440
      Width           =   12615
      Begin VB.CommandButton cmdPesanBrgRuangan 
         Caption         =   "Pesan Barang &Supplier"
         Height          =   495
         Left            =   7680
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   10800
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   9240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPesanBrgSupp 
         Caption         =   "Pesan Barang &Supplier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Daftar Stok Barang Non Medis Ruangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12615
      Begin VB.TextBox txtCariAsalBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7185
         MaxLength       =   50
         TabIndex        =   5
         Top             =   5640
         Width           =   1680
      End
      Begin VB.TextBox txtCariJenisBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   4
         Top             =   5640
         Width           =   1680
      End
      Begin VB.TextBox txtCariBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   5640
         Width           =   1680
      End
      Begin VB.CheckBox chkCheck 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   200
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   2
         Top             =   640
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid fxDaftarStokBarangRuanganNonMedis 
         Height          =   5055
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
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
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Index           =   1
         Left            =   6120
         TabIndex        =   9
         Top             =   5685
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Barang"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   5685
         Width           =   1005
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Index           =   6
         Left            =   3000
         TabIndex        =   7
         Top             =   5685
         Width           =   1065
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
      Height          =   975
      Left            =   10920
      Picture         =   "frmMonitoringStokBarangNonMedis.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMonitoringStokBarangNonMedis.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "frmMonitoringStokBarangNonMedis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkCheck_Click()
On Error GoTo errLoad
    If chkCheck.Value = vbChecked Then
        fxDaftarStokBarangRuanganNonMedis.TextMatrix(fxDaftarStokBarangRuanganNonMedis.Row, 0) = Chr$(187)
        fxDaftarStokBarangRuanganNonMedis.TextMatrix(fxDaftarStokBarangRuanganNonMedis.Row, 21) = 1
    Exit Sub
    Else
        fxDaftarStokBarangRuanganNonMedis.TextMatrix(fxDaftarStokBarangRuanganNonMedis.Row, 0) = ""
        fxDaftarStokBarangRuanganNonMedis.TextMatrix(fxDaftarStokBarangRuanganNonMedis.Row, 21) = 0
    Exit Sub
    End If
    
errLoad:
    msubPesanError
End Sub

Private Sub chkCheck_LostFocus()
    chkCheck.Visible = False
End Sub

Private Sub cmdBatal_Click()
    txtCariJenisBarang.Text = ""
    txtCariBarang.Text = ""
    txtCariAsalBarang.Text = ""
    Call subLoadFgSource
End Sub

Private Sub cmdPesanBrgRuangan_Click()
Dim m As Integer
On Error GoTo errLoad
'    If fxDaftarStokBarangRuanganNonMedis = "" Or fxDaftarStokBarangRuanganNonMedis <> "»" Then Exit Sub
    
    
    With fxDaftarStokBarangRuanganNonMedis
        For m = 1 To .Rows - 1
            If fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 0) <> "" Then
                strNoTerima = fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 19)
            End If
        Next m
    End With
    
    With frmPemesanankeSupplier
'        .dcStatusBarang.BoundText = "01"
'        Call msubDcSource(.dcRuanganTujuan, rs, "SELECT KdRuangan,NamaRuangan FROM V_StrukOrderRuanganTujuan WHERE KdKelompokBarang='" & .dcStatusBarang.BoundText & "' AND StatusEnabled=1 AND KdRuangan <> '" & mstrKdRuangan & "' ORDER BY NamaRuangan")
'        If rs.EOF = False Then .dcRuanganTujuan.BoundText = rs(0).Value
        .Show
        
            Dim a As Integer
            a = 1
            For m = 1 To fxDaftarStokBarangRuanganNonMedis.Rows - 1
                If fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 0) <> "" Then

                    strsqlx = "SELECT TOP (1) MasterBarangNonMedis.KdBarang, SupplierPabrik.KdPabrik, StrukTerima.KdSupplier " & _
                             "FROM StrukTerima INNER JOIN " & _
                             "SupplierPabrik ON StrukTerima.KdSupplier = SupplierPabrik.KdSupplier INNER JOIN " & _
                             "DetailTerimaBarangNonMedis ON StrukTerima.NoTerima = DetailTerimaBarangNonMedis.NoTerima INNER JOIN " & _
                             "MasterBarangNonMedis ON DetailTerimaBarangNonMedis.KdBarang = MasterBarangNonMedis.KdBarang " & _
                             "WHERE (MasterBarangNonMedis.KdBarang = '" & fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 1) & "') And DetailTerimaBarangNonMedis.kdAsal = '" & fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 3) & "' order by StrukTerima.noterima desc"
                    Set rsC = Nothing
                    Call msubRecFO(rsC, strsql)

                    .dcSumberDana.BoundText = fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 3)
                    .fgData.TextMatrix(a, 0) = fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 1)
                    .fgData.TextMatrix(a, 1) = fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 2)
                    .fgData.TextMatrix(a, 2) = fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 4)
                    
                    
                    strsql = "Select * From V_CariBarangNM Where Kdbarang='" & .fgData.TextMatrix(a, 0) & "'"
                    Set rs = Nothing
                    Call msubRecFO(rs, strsql)
                    .fgData.TextMatrix(a, 3) = IIf(IsNull(rs("Satuan").Value), "", rs("Satuan").Value)
                    
                    strsql = "select  * from V_CariBarangNonMedis " & _
                             " where kdbarang like '%" & fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 1) & "%' And KdRuangan='" & mstrKdRuangan & "' "
                    Set rsD = Nothing
                    Call msubRecFO(rsD, strsql)
                    
                    If rsD.EOF = False Then
                        '.fgData.TextMatrix(a, 3) = IIf(IsNull(rsD("Satuan").Value), "", rsD("Satuan").Value)
                        .fgData.TextMatrix(a, 6) = IIf(IsNull(rsD("Discount").Value), "", rsD("Discount").Value)
                    Else
                        '.fgData.TextMatrix(a, 3) = ""
                        .fgData.TextMatrix(a, 6) = "0"
                    End If
                    
                    If rsD.EOF = False Then
                        .fgData.TextMatrix(a, 4) = rs(5).Value
                    Else
                        .fgData.TextMatrix(a, 4) = rs(5).Value
                    End If
                    
                    strsql = "Select sum(JmlStok) as Stok From V_InfoStokNonMedisRuanganFIFO Where  Kdbarang='" & fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 1) & "' and KdAsal='" & fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 3) & "' and KdRuangan='" & mstrKdRuangan & "'"
                    Set rsD = Nothing
                    Call msubRecFO(rsD, strsql)
                    If rsD.EOF = False Then
                        .fgData.TextMatrix(a, 5) = IIf(IsNull(rsD("Stok").Value), "", rsD("Stok").Value)
                    Else
                        .fgData.TextMatrix(a, 5) = 0
                    End If
        
                    .fgData.TextMatrix(a, 7) = fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 3)
                    .fgData.TextMatrix(a, 8) = fxDaftarStokBarangRuanganNonMedis.TextMatrix(m, 12)
                    .fgData.TextMatrix(a, 9) = ""
                    .fgData.TextMatrix(a, 10) = ""
                    
                    a = a + 1
                    .fgData.Rows = .fgData.Rows + 1
                    .fgData.SetFocus
                    .fgData.Col = 2
                    
                End If
            Next m
        .fgData.Rows = .fgData.Rows - 1
    End With
errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadFgSource
errLoad:
End Sub

Private Sub subclear_Grid()
    Dim i As Integer
    With fxDaftarStokBarangRuanganNonMedis
        .Clear
        .Rows = 2
        .Cols = 22
        .ColWidth(0) = 400
        .ColWidth(1) = 0        'KdBarang
        .ColWidth(2) = 4000     'NamaBarang
        .ColAlignment(2) = flexAlignLeftCenter
        .ColWidth(3) = 0        'KdAsal
        .ColWidth(4) = 1700     'NamaAsal
        .ColWidth(5) = 0        'KdRuangan
        .ColWidth(6) = 0        'NamaRuangan
        .ColWidth(7) = 0        'JmlStok
        .ColWidth(8) = 0        'JmlStokFIFO
        .ColWidth(9) = 1800     'JmlStokRuangan
        .ColWidth(10) = 0       'KdDetailJenisBarang
        .ColWidth(11) = 1800    'DetailJenisBarang
        .ColWidth(12) = 0       'KdSatuanJmlB
        .ColWidth(13) = 0       'SatuanJmlB
        .ColWidth(14) = 0       'JmlKemasan
        .ColWidth(15) = 1000    'JmlMinimum
        .ColWidth(16) = 0       'Lokasi
        .ColWidth(17) = 1500    'KdJenisAsset
        .ColWidth(18) = 2000    'JenisBarang
        .ColWidth(19) = 2000    'NoTerima
        .ColWidth(20) = 2000    'NoRegisterAsset
        .ColWidth(21) = 0
        Call SetGridPesanMenu
    End With
    
End Sub

Private Sub SetGridPesanMenu()
    On Error Resume Next
    Dim i As Integer
    With fxDaftarStokBarangRuanganNonMedis
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KdBarang"
        .TextMatrix(0, 2) = "NamaBarang"
        .TextMatrix(0, 3) = "KdAsal"
        .TextMatrix(0, 4) = "NamaAsal"
        .TextMatrix(0, 5) = "KdRuangan"
        .TextMatrix(0, 6) = "NamaRuangan"
        .TextMatrix(0, 7) = "JmlStok"
        .TextMatrix(0, 8) = "JmlStokFIFO"
        .TextMatrix(0, 9) = "JmlStokRuangan"
        .TextMatrix(0, 10) = "KdDetailJenisBarang"
        .TextMatrix(0, 11) = "DetailJenisBarang"
        .TextMatrix(0, 12) = "KdSatuanJmlB"
        .TextMatrix(0, 13) = "SatuanJmlB"
        .TextMatrix(0, 14) = "JmlKemasan"
        .TextMatrix(0, 15) = "JmlMinimum"
        .TextMatrix(0, 16) = "Lokasi"
        .TextMatrix(0, 17) = "KdJenisAsset"
        .TextMatrix(0, 18) = "JenisBarang"
        .TextMatrix(0, 19) = "NoTerima"
        .TextMatrix(0, 20) = "NoRegisterAsset"
        .TextMatrix(0, 21) = ""
    End With
End Sub

Sub subLoadFgSource()
    Dim i As Integer
    Dim strJmlMin As Integer
    On Error GoTo errLoad
    
    strsql = "SELECT * " & _
    " FROM V_InfoStokNonMedisRuanganFIFO " & _
    " WHERE NamaBarang LIKE '%" & txtCariBarang & "%' AND DetailJenisBarang LIKE '%" & txtCariJenisBarang & "%' AND NamaAsal LIKE '%" & txtCariAsalBarang & "%'AND kdRuangan = '" & mstrKdRuangan & "'"
    
    
    Set rs = Nothing
    rs.Open strsql, dbConn, adOpenForwardOnly, adLockReadOnly
    Call subclear_Grid
    fxDaftarStokBarangRuanganNonMedis.Rows = rs.RecordCount + 1
    For i = 1 To rs.RecordCount
    With fxDaftarStokBarangRuanganNonMedis
        .TextMatrix(i, 0) = ""
        .TextMatrix(i, 1) = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0))
        .TextMatrix(i, 2) = IIf(IsNull(rs.Fields(1).Value), 0, rs.Fields(1))
        .TextMatrix(i, 3) = IIf(IsNull(rs.Fields(2).Value), 0, rs.Fields(2))
        .TextMatrix(i, 4) = IIf(IsNull(rs.Fields(3).Value), 0, rs.Fields(3))

        .TextMatrix(i, 5) = IIf(IsNull(rs.Fields(4).Value), "-", rs.Fields(4))
        .TextMatrix(i, 6) = IIf(IsNull(rs.Fields(5).Value), "-", rs.Fields(5))
        .TextMatrix(i, 7) = IIf(IsNull(rs.Fields(6).Value), 0, rs.Fields(6))
        .TextMatrix(i, 8) = IIf(IsNull(rs.Fields(7).Value), "-", rs.Fields(7))
        .TextMatrix(i, 9) = IIf(IsNull(rs.Fields(8).Value), "-", rs.Fields(8))
        .TextMatrix(i, 10) = IIf(IsNull(rs.Fields(9).Value), "-", rs.Fields(9))
        .TextMatrix(i, 11) = IIf(IsNull(rs.Fields(10).Value), 0, rs.Fields(10))
        .TextMatrix(i, 12) = IIf(IsNull(rs.Fields(11).Value), "-", rs.Fields(11))
        .TextMatrix(i, 13) = IIf(IsNull(rs.Fields(12).Value), 0, rs.Fields(12))

        .TextMatrix(i, 14) = IIf(IsNull(rs.Fields(13).Value), 0, rs.Fields(13))
        .TextMatrix(i, 15) = IIf(IsNull(rs.Fields(14).Value), 0, rs.Fields(14))
        .TextMatrix(i, 16) = IIf(IsNull(rs.Fields(15).Value), 0, rs.Fields(15))
        .TextMatrix(i, 17) = IIf(IsNull(rs.Fields(16).Value), "-", rs.Fields(16))

        .TextMatrix(i, 18) = IIf(IsNull(rs.Fields(17).Value), "-", rs.Fields(17))
        .TextMatrix(i, 19) = IIf(IsNull(rs.Fields(18).Value), "-", rs.Fields(18))
        .TextMatrix(i, 20) = IIf(IsNull(rs.Fields(19).Value), "-", rs.Fields(19))
        .TextMatrix(i, 21) = 1
        
        strJmlMin = .TextMatrix(i, 15)
        'lblJmlData.Caption = 0 & " / " & .ApproxCount & " Data"
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 0
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 1
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then

            .Row = i
            .Col = 2
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 3
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then

            .Row = i
            .Col = 4
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 5
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 6
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 7
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 8
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 9
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 10
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 11
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 12
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 13
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 14
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 15
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 16
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 17
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 18
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 19
            .CellBackColor = vbRed
        End If

        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 20
            .CellBackColor = vbRed
        End If
        
        If .TextMatrix(i, 7) <= strJmlMin Then
            .Row = i
            .Col = 21
            .CellBackColor = vbRed
        End If
        
    End With
    rs.MoveNext
    Next i
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub fxDaftarStokBarangRuanganNonMedis_Click()
On Error GoTo hell
    If fxDaftarStokBarangRuanganNonMedis.Rows = 1 Then Exit Sub
    If fxDaftarStokBarangRuanganNonMedis.Col = 0 Then
        chkCheck.Visible = True
        chkCheck.Top = fxDaftarStokBarangRuanganNonMedis.RowPos(fxDaftarStokBarangRuanganNonMedis.Row) + 375
        Dim intChk As Integer
        intChk = ((fxDaftarStokBarangRuanganNonMedis.ColPos(fxDaftarStokBarangRuanganNonMedis.Col + 1) - fxDaftarStokBarangRuangan.ColPos(fxDaftarStokBarangRuangan.Col)) / 2)
        chkCheck.Left = fxDaftarStokBarangRuanganNonMedis.ColPos(fxDaftarStokBarangRuanganNonMedis.Col) + intChk - 10 ' - 250  '+ intChk
        chkCheck.SetFocus
        If fxDaftarStokBarangRuanganNonMedis.Col <> 0 Then
            If fxDaftarStokBarangRuanganNonMedis.TextMatrix(fxDaftarStokBarangRuanganNonMedis.Row, 0) <> "" Then
                chkCheck.Value = 1
            Else
                chkCheck.Value = 0
            End If
        End If
    End If
hell:
End Sub

Private Sub txtCariAsalBarang_Change()
On Error GoTo errLoad
    Call subLoadFgSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtCariBarang_Change()
On Error GoTo errLoad
    Call subLoadFgSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtCariJenisBarang_Change()
On Error GoTo errLoad
    Call subLoadFgSource
    Exit Sub
errLoad:
    Call msubPesanError
End Sub
