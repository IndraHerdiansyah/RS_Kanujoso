VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmInsentifKaryawan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Jasa Remunerasi /Insentif Karyawan"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11955
   Icon            =   "frmInsentifKaryawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11955
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   11685
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
         Height          =   615
         Left            =   9720
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Cetak"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8040
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin MSComctlLib.ProgressBar pbData 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
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
         Left            =   7200
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kriteria Pencarian "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   11835
      Begin MSComCtl2.DTPicker dtpPeriode 
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
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
         CustomFormat    =   "MMMM yyyy"
         Format          =   61865987
         CurrentDate     =   40282
      End
      Begin VB.CommandButton cmdProses 
         Caption         =   "&Proses.."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9960
         TabIndex        =   10
         Top             =   240
         Width           =   1665
      End
      Begin VB.TextBox txtIdAkhir 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "0000000001"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcUnitKerja 
         Height          =   360
         Left            =   720
         TabIndex        =   8
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Unit Kerja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ID Pegawai Akhir"
         Height          =   195
         Left            =   11880
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1230
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
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   11775
      _ExtentX        =   20770
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
      Picture         =   "frmInsentifKaryawan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   10080
      Picture         =   "frmInsentifKaryawan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInsentifKaryawan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInsentifKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilter As String
Dim sKodePegawai As String
Dim sKdRuangan As String
Dim sNamaPegawai As String
Dim sNamaRuangan As String
Dim cInsentif As Currency
Dim cPPh As Currency
Dim cDiterima As Currency
Dim iIndex As Integer

Private Function HitungKolom(f_Bulan As String, f_Tahun As String, f_BulanPelayanan As String, f_TahunPelayanan As String, f_KdRuangan As String, F_IdPegawai As String) As Boolean
    On Error GoTo hell_
    HitungKolom = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adVarChar, adParamInput, 20, Trim(F_IdPegawai))
        .Parameters.Append .CreateParameter("KdKelompokPasien", adVarChar, adParamInput, 6, f_KdRuangan)
        .Parameters.Append .CreateParameter("Bulan", adVarChar, adParamInput, 20, f_Bulan)
        .Parameters.Append .CreateParameter("Tahun", adVarChar, adParamInput, 20, f_Tahun)
         .Parameters.Append .CreateParameter("BulanPelayanan", adVarChar, adParamInput, 20, f_BulanPelayanan)
         .Parameters.Append .CreateParameter("TahunPelayanan", adVarChar, adParamInput, 20, f_TahunPelayanan)
        .Parameters.Append .CreateParameter("outputInsentif", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("outputIndexOrang", adInteger, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("outputPph", adCurrency, adParamOutput, , Null)
        
        .ActiveConnection = dbConn
        .CommandText = "HitungJasaRemunerasiPegawai"
        .CommandType = adCmdStoredProc
        .Execute
        
        
      
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Hapus Rekap Komponen Biaya Pelayanan", vbCritical, "Validasi"
            HitungKolom = False
        Else
            If Not IsNull(.Parameters("outputInsentif").Value) Then cInsentif = .Parameters("outputInsentif").Value Else cInsentif = 0
            If Not IsNull(.Parameters("outputIndexOrang").Value) Then iIndex = .Parameters("outputIndexOrang").Value Else iIndex = 0
            If Not IsNull(.Parameters("outputPph").Value) Then cPPh = .Parameters("outputPph").Value Else cPPh = 0
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    
    End With
    
Exit Function
hell_:
    HitungKolom = False
    Call msubPesanError("-HitungJasaRemunerasiPegawai")
End Function

Private Sub setgrid()
'Dim i As Integer
    With fgData
        .clear
        .Rows = 2
        .Cols = 10
        .Row = 0
        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 500
            
        Next
       
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "No"
        .TextMatrix(0, 2) = "Nama Ruangan"
        .TextMatrix(0, 3) = "Nama Pegawai"
        .TextMatrix(0, 4) = "Insentif" '
        .TextMatrix(0, 5) = "Pph ps 21" '
        .TextMatrix(0, 6) = "Yang diterima"
        .TextMatrix(0, 7) = "Index"
        
        .TextMatrix(0, 8) = "KdRuangan"
        .TextMatrix(0, 9) = "kdPegawai"
        
        .ColWidth(0) = 0
        .ColWidth(1) = 400
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 1800
        .ColWidth(6) = 2000
        .ColWidth(7) = 1400
         .ColWidth(8) = 0
        .ColWidth(9) = 0
        
    End With
End Sub

Private Sub cmdCetak_Click()
On Error GoTo hell
    For i = 1 To fgData.Rows - 2
    With fgData
        strQuery = "Insert into vJasaRemunerasiPegawai " & _
                " values ('" & strNamaHostLocal & "', '" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "'," & CCur(.TextMatrix(i, 4)) & "," & CCur(.TextMatrix(i, 5)) & "," & CCur(.TextMatrix(i, 6)) & ",'" & .TextMatrix(i, 7) & "'" & _
                ", '" & MonthName(Month(Me.dtpPeriode.Value)) & "' ,'" & Year(Me.dtpPeriode.Value) & "'" & _
                " )"
        dbConn.Execute strQuery
    End With
    Next i

    frmCetakIndexPegawaiPerRuangan.Show
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdProses_Click()
On Error GoTo hell_
Dim i As Integer
Dim dBulanPelayanan As Date

dBulanPelayanan = DateAdd("m", -1, Format(dtpPeriode.Value, "dd MMM yyyy"))
Call setgrid
strsql = "SELECT IdPegawai, NamaLengkap, NamaRuangan, KdRuangan" & _
        " FROM  V_DataPegawaiInsentif where KdRuangan = '" & dcUnitKerja.BoundText & "' "
Call msubRecFO(rs, strsql)
If rs.EOF = True Then Exit Sub
MousePointer = vbHourglass
cmdProses.Enabled = False
cmdCetak.Enabled = True
For i = 1 To rs.RecordCount
pbData.Max = rs.RecordCount
DoEvents
lblPersen.Caption = Int((i / rs.RecordCount) * 100) & "%"
sKodePegawai = rs.Fields("IdPegawai")
sKdRuangan = rs.Fields("KdRuangan")
sNamaPegawai = rs.Fields("NamaLengkap")
sNamaRuangan = rs.Fields("NamaRuangan")

With fgData
    .TextMatrix(i, 1) = i ' nomor
    .TextMatrix(i, 2) = sNamaRuangan ' nama pegawai
    .TextMatrix(i, 3) = sNamaPegawai
   
    If HitungKolom(Month(dtpPeriode.Value), Year(dtpPeriode.Value), Month(dBulanPelayanan), Year(dBulanPelayanan), sKdRuangan, sKodePegawai) = False Then Exit Sub
          

    .TextMatrix(i, 4) = FormatCurrency(cInsentif)  'hitung insentif
    .TextMatrix(i, 5) = FormatCurrency(cPPh) 'pph
    .TextMatrix(i, 6) = FormatCurrency(cInsentif - cPPh) 'yang diterima
    .TextMatrix(i, 7) = iIndex 'index
    .Rows = 2 + i

End With
rs.MoveNext
pbData.Value = Int(pbData.Value) + 1
Next i
cmdProses.Enabled = True
MousePointer = vbDefault
MsgBox "Proese Penghitungan Jasa Remunerasi Pegawai Berhasil..", vbInformation + vbOKOnly, "Informasi"
pbData.Value = 0.0001
Exit Sub
hell_:
    cmdProses.Enabled = True
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call setgrid
    Call subLoadDcSource
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmKriteriaLaporan = Nothing
End Sub

Private Sub subLoadDcSource()
    'strSQL = "Select * from Ruangan order by NamaRuangan"
    strsql = "Select kdsubruangkerja, subruangkerja from subruangkerja where StatusEnabled = 1 order by subruangkerja"
    Call msubDcSource(dcUnitKerja, rs, strsql)
End Sub

