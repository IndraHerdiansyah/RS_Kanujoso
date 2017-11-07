VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataPerhitunganInsentif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Perhitungan Insentif"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   Icon            =   "frmDataPerhitunganInsentif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14055
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7200
      TabIndex        =   21
      Text            =   "0"
      Top             =   8040
      Width           =   2265
   End
   Begin VB.TextBox txtNoBKK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      MaxLength       =   15
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtTotalPrestasi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   18
      Top             =   8040
      Width           =   2505
   End
   Begin VB.TextBox txtTotalDasar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   11
      Top             =   8040
      Width           =   2265
   End
   Begin VB.CommandButton cmdBaru 
      Caption         =   "&Batal"
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
      Left            =   13800
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
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
      Left            =   10680
      TabIndex        =   2
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
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
      Left            =   12360
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdHitung 
      Caption         =   "Proses"
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
      Left            =   11400
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   11175
      Begin VB.TextBox txtTotalPrestasiDibagi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Width           =   2385
      End
      Begin VB.TextBox txtTotalDasarDibagi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Insentif Prestasi Dibagikan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Insentif Dasar Dibagikan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pengisian Insentif Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   13935
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   2760
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgIndex 
         Height          =   5175
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   9128
         _Version        =   393216
         BackColor       =   -2147483624
         WordWrap        =   -1  'True
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
   Begin MSComctlLib.ProgressBar pbData 
      Height          =   480
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   847
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   200
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtpBlnHitung 
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   8040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
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
      CustomFormat    =   "MMMM, yyyy"
      Format          =   55181315
      UpDown          =   -1  'True
      CurrentDate     =   38231
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Total Insentif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7200
      TabIndex        =   22
      Top             =   7800
      Width           =   1710
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Score Insentif Prestasi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4560
      TabIndex        =   19
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Score Insentif Dasar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2160
      TabIndex        =   14
      Top             =   7800
      Width           =   2115
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Bulan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   435
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataPerhitunganInsentif.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12240
      Picture         =   "frmDataPerhitunganInsentif.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPerhitunganInsentif.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataPerhitunganInsentif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String
Dim strSQL As String
Dim strQuerySQL As String
Dim strLFilterPegawai As String
Dim intLJmlPegawai As Integer
Dim strLIdPegawai As String
Dim strLKdJabatan As String
Dim strLKdPendidikan As String
Const strLOrder As String = "ORDER BY NamaLengkap"
Dim blnLPegawaiFocus As Boolean
Dim intLJmlIndex As Integer
Dim j As Integer

Private Sub cmdHitung_Click()
    If txtTotalDasarDibagi.Text = "" Then MsgBox "Silahkan lengkapi Jumlah Total Insentif dasar yang dibagikan", vbExclamation, "Validasi": Exit Sub
    If txtTotalPrestasiDibagi.Text = "" Then MsgBox "Silahkan lengkapi Jumlah Total Insentif Prestasi yang dibagikan", vbExclamation, "Validasi": Exit Sub

    Call subLoadIndexPegawai
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer
    Dim Jml As Currency
    If hgIndex.TextMatrix(1, 0) = "" Then MsgBox "Data tidak ada", vbExclamation, "Validasi": Exit Sub

    If sp_AddStrukBuktiKasKeluar() = False Then Exit Sub
    For i = 1 To hgIndex.Rows - 2
        With hgIndex
            Jml = CCur(.TextMatrix(i, 7)) + CCur(.TextMatrix(i, 11))
            If sp_DetailPembayaranInsentif(.TextMatrix(i, 1), CCur(Jml)) = False Then Exit Sub
        End With
    Next i
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_AddStrukBuktiKasKeluar() As Boolean
    On Error GoTo errLoad

    sp_AddStrukBuktiKasKeluar = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglBKK", adDate, adParamInput, , Format(dtpBlnHitung, "yyyy/MM/01 HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdCaraBayar", adChar, adParamInput, 2, "01")
        .Parameters.Append .CreateParameter("NamaBank", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("NoAccount", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("AtasNama", adVarChar, adParamInput, 50, Null)
        .Parameters.Append .CreateParameter("JmlBayar", adCurrency, adParamInput, , CCur(txtTotal.Text))
        .Parameters.Append .CreateParameter("Administrasi", adCurrency, adParamInput, , 0)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("KdTransaksiKasir", adChar, adParamInput, 6, Null)
        .Parameters.Append .CreateParameter("OutputNoBKK", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "Add_StrukBuktiKasKeluar"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Data", vbCritical, "Validasi"
            sp_AddStrukBuktiKasKeluar = False
        Else
            If Not IsNull(.Parameters("OutputNoBKK").Value) Then txtNoBKK = .Parameters("OutputNoBKK").Value
        End If
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_AddStrukBuktiKasKeluar = False
    Call msubPesanError
End Function

Private Function sp_DetailPembayaranInsentif(f_IdPegawai As String, f_Total As Currency) As Boolean
    On Error GoTo errLoad
    sp_DetailPembayaranInsentif = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoBKK", adChar, adParamInput, 10, txtNoBKK.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_IdPegawai)
        .Parameters.Append .CreateParameter("Periode", adDate, adParamInput, , Format(dtpBlnHitung, "yyyy/MM/01 HH:mm:ss"))
        .Parameters.Append .CreateParameter("Jumlah", adCurrency, adParamInput, , CCur(f_Total))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PembayaranInsentif"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            sp_DetailPembayaranInsentif = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError(" sp_DetailPembayaranInsentif")
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpBlnHitung_Change()
    dtpBlnHitung.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subSetGrid

    dtpBlnHitung.Value = Format(Now, "MMMM, yyyy")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDataPerhitunganInsentif = Nothing
End Sub

Private Sub subLoadIndexPegawai()
    On Error Resume Next   '' 0/0 overflow

    pbData.Value = 0

    strSQL = "Select distinct IdPegawai,NamaLengkap,SUM(Pendidikan) as Pendidikan, SUM(Jabatan) as Jabatan, SUM(MasaKerja) as MasaKerja, SUM(Absensi) as Absensi, SUM(BebanKerja) as BebanKerja " & _
    "from V_InsentifPegawai Group by IdPegawai,NamaLengkap order by NamaLengkap"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intLJmlIndex = rs.RecordCount
    With hgIndex

        pbData.Value = 0
        pbData.Max = rs.RecordCount

        For j = 1 To rs.RecordCount
            hgIndex.Rows = rs.RecordCount + 1
            hgIndex.TextMatrix(j, 0) = j
            hgIndex.TextMatrix(j, 1) = rs("IdPegawai").Value
            hgIndex.TextMatrix(j, 2) = rs("NamaLengkap").Value
            hgIndex.TextMatrix(j, 3) = rs("Pendidikan").Value
            hgIndex.TextMatrix(j, 4) = rs("Jabatan").Value
            hgIndex.TextMatrix(j, 5) = rs("MasaKerja").Value

            txtTotalDasar.Text = Val(txtTotalDasar.Text) + hgIndex.TextMatrix(j, 5)
            hgIndex.TextMatrix(j, 6) = (rs("Pendidikan").Value * rs("Jabatan").Value * rs("MasaKerja").Value)
            hgIndex.TextMatrix(j, 7) = ((hgIndex.TextMatrix(j, 6) / Val(txtTotalDasar.Text)) * Val(txtTotalDasarDibagi.Text))

            hgIndex.TextMatrix(j, 8) = 0
            hgIndex.TextMatrix(j, 9) = rs("BebanKerja").Value
            hgIndex.TextMatrix(j, 10) = hgIndex.TextMatrix(j, 9) * hgIndex.TextMatrix(j, 8)
            txtTotalPrestasi.Text = Val(txtTotalPrestasi.Text) + hgIndex.TextMatrix(j, 10)
            hgIndex.TextMatrix(j, 11) = ((hgIndex.TextMatrix(j, 10) / Val(txtTotalPrestasi.Text)) * Val(txtTotalPrestasiDibagi.Text))

            txtTotal.Text = FormatCurrency(Val(txtTotal.Text) + hgIndex.TextMatrix(j, 7) + hgIndex.TextMatrix(j, 11), 2)

            rs.MoveNext
            hgIndex.row = j + 1
        Next j

    End With
End Sub

Private Sub subHitungIndex()
    txtTotalDasar.Text = 0
    txtTotalPrestasi.Text = 0
    For i = 1 To hgIndex.Rows - 1
        If hgIndex.TextMatrix(i, 6) = "" Then GoTo nexti
        txtTotalDasar.Text = CInt(txtTotalDasar.Text) + CInt(hgIndex.TextMatrix(i, 6))
        If hgIndex.TextMatrix(i, 10) = "" Then GoTo nexti
        txtTotalDasar.Text = CInt(txtTotalPrestasi.Text) + CInt(hgIndex.TextMatrix(i, 10))
nexti:
    Next i
End Sub

Sub kosong()

End Sub

Private Sub hgIndex_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    txtisi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    Select Case hgIndex.Col

        Case 8 'score absensi
            txtisi.MaxLength = 5
            Call subLoadText
            txtisi.Text = Chr(KeyAscii)
            txtisi.SelStart = Len(txtisi.Text)

    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtisi.Left = hgIndex.Left

    For i = 0 To hgIndex.Col - 1
        txtisi.Left = txtisi.Left + hgIndex.ColWidth(i)
    Next i
    txtisi.Visible = True
    txtisi.Top = hgIndex.Top - 7

    For i = 0 To hgIndex.row - 1
        txtisi.Top = txtisi.Top + hgIndex.RowHeight(i)
    Next i

    If hgIndex.TopRow > 1 Then
        txtisi.Top = txtisi.Top - ((hgIndex.TopRow - 1) * hgIndex.RowHeight(1))
    End If

    txtisi.Width = hgIndex.ColWidth(hgIndex.Col)
    txtisi.Height = hgIndex.RowHeight(hgIndex.row)

    txtisi.Visible = True
    txtisi.SelStart = Len(txtisi.Text)
    txtisi.SetFocus

End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim i As Integer
    If KeyAscii = 27 Then
        txtisi.Visible = False
        hgIndex.SetFocus
    End If

    If KeyAscii = 13 Then
        Call SetKeyPressToNumber(KeyAscii)
        Select Case hgIndex.Col

            Case 8 'score absensi
                hgIndex.TextMatrix(hgIndex.row, hgIndex.Col) = txtisi.Text

                With hgIndex
                    .TextMatrix(hgIndex.row, 11) = CDbl(hgIndex.TextMatrix(hgIndex.row, 8)) * CDbl(hgIndex.TextMatrix(hgIndex.row, 9))
                    txtTotal.Text = FormatCurrency(CCur(txtTotal.Text) + hgIndex.TextMatrix(hgIndex.row, 7) + hgIndex.TextMatrix(hgIndex.row, 11), 2)
                End With
                txtisi.Visible = False
                hgIndex.SetFocus
        End Select

    End If

    If hgIndex.Col = 6 Then
        If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(",")) Then KeyAscii = 0
    End If

End Sub

Private Sub subSetGrid()
    On Error GoTo errLoad
    With hgIndex
        .clear
        .Rows = 2
        .Cols = 12
        .Refresh

        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "IdPegawai"
        .TextMatrix(0, 2) = "Nama Pegawai"
        .TextMatrix(0, 3) = "Score Pendidikan"
        .TextMatrix(0, 4) = "Score Jabatan"
        .TextMatrix(0, 5) = "Score Masa Kerja"
        .TextMatrix(0, 6) = "Total Score Insentif Dasar"
        .TextMatrix(0, 7) = "Total Terima Insentif Dasar"
        .TextMatrix(0, 8) = "Score Absensi"
        .TextMatrix(0, 9) = "Score Beban Kerja"
        .TextMatrix(0, 10) = "Total Score Insentif Prestasi"
        .TextMatrix(0, 11) = "Total Terima Insentif Prestasi"

        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1300
        .ColWidth(7) = 1300
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1300
        .ColWidth(11) = 1300

        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter

        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .Refresh
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtisi.Visible = False
End Sub

Private Sub txtTotalDasarDibagi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTotalPrestasiDibagi.SetFocus
End Sub

Private Sub txtTotalDasarDibagi_LostFocus()
    On Error Resume Next
    txtTotalDasarDibagi.Text = IIf(Val(txtTotalDasarDibagi.Text) = 0, 0, FormatCurrency(txtTotalDasarDibagi.Text, 2))
End Sub

Private Sub txtTotalPrestasiDibagi_LostFocus()
    On Error Resume Next
    txtTotalPrestasiDibagi.Text = IIf(Val(txtTotalPrestasiDibagi.Text) = 0, 0, FormatCurrency(txtTotalPrestasiDibagi.Text, 2))
End Sub

Private Sub txtTotalPrestasiDibagi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdHitung.SetFocus
End Sub
