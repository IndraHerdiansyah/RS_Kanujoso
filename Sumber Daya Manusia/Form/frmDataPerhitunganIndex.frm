VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataPerhitunganIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Perhitungan Index"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   Icon            =   "frmDataPerhitunganIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   11190
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
      Left            =   6240
      TabIndex        =   7
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
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
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   6240
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
      Left            =   9600
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdHitung 
      Caption         =   "&Hitung Index"
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
      Left            =   1680
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
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
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   1080
      Width           =   11175
      Begin VB.TextBox txtIdPegawai 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPegawai 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4800
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtJnsPeg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7920
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ID. Pegawai"
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
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pegawai"
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
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pegawai"
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
         Left            =   5400
         TabIndex        =   17
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "JK"
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
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan"
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
         Left            =   7920
         TabIndex        =   15
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pengisian Index Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   11175
      Begin VB.TextBox txtTotalHasil 
         Alignment       =   2  'Center
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
         TabIndex        =   23
         Top             =   3600
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo dcDetKomIndex 
         Height          =   315
         Left            =   3240
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgIndex 
         Height          =   2535
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   -2147483624
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
      Begin MSComCtl2.DTPicker dtpBlnHitung 
         Height          =   360
         Left            =   120
         TabIndex        =   24
         Top             =   3600
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
         Format          =   129957891
         UpDown          =   -1  'True
         CurrentDate     =   38231
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bulan Penghitungan"
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
         TabIndex        =   26
         Top             =   3360
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Index"
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
         TabIndex        =   25
         Top             =   3360
         Width           =   945
      End
   End
   Begin VB.Frame fraPegawai 
      Caption         =   "Data Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   11175
      Begin MSDataGridLib.DataGrid dgPegawai 
         Height          =   3615
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgPegawai 
         Height          =   2535
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4471
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   19
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
      TabIndex        =   21
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
   Begin VB.Label Label8 
      Caption         =   "Grid Harus di double Klik"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataPerhitunganIndex.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9360
      Picture         =   "frmDataPerhitunganIndex.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPerhitunganIndex.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataPerhitunganIndex"
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

Private Sub cmdBaru_Click()

    cmdHitung.Enabled = True
    cmdsimpan.Enabled = False
    Call kosong
    Call subLoadDataPegawai
    txtnamapegawai.SetFocus
    Frame2.Visible = False
    frapegawai.Visible = True
    txtTotalHasil.Text = ""
    txtidpegawai.Text = ""
End Sub

Private Sub cmdHitung_Click()
    cmdHitung.Enabled = False
    Call dgPegawai_KeyPress(13)
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    Dim intRow As Integer
    Dim strKdDetailKomponenIndex As String
    Dim strKdKomponenIndex As String

    strSQL = "SELECT * FROM TotalScoreIndex WHERE IdPegawai = '" & dgPegawai.Columns(0) & "' AND Month(TglHitung) = '" & dtpBlnHitung.Month & "' AND Year(TglHitung) = '" & dtpBlnHitung.Year & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        If Year(rs(0).Value) = Year(dtpBlnHitung.Value) And Month(rs(0).Value) = Month(dtpBlnHitung.Value) Then
            strQuerySQL = "DELETE FROM TotalScoreIndex WHERE IdPegawai = '" & dgPegawai.Columns(0) & "' AND Month(TglHitung) = '" & dtpBlnHitung.Month & "' AND Year(TglHitung) = '" & dtpBlnHitung.Year & "'"
            dbConn.Execute strQuerySQL
            rs.Close
        End If
    End If

    Set rs = Nothing
    With hgIndex
        For intRow = 2 To hgIndex.Rows
            If hgIndex.TextMatrix(intRow - 1, 3) <> "" Then
                strLSQL = "SELECT KdKomponenIndex,KdDetailKomponenIndex FROM V_DetailKomponenIndex WHERE DetailKomponenIndex = '" & hgIndex.TextMatrix(intRow - 1, 2) & "' AND KomponenIndex = '" & hgIndex.TextMatrix(intRow - 1, 6) & "'"
                rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                strKdDetailKomponenIndex = rs.Fields("KdDetailKomponenIndex")
                strKdKomponenIndex = rs.Fields("KdKomponenIndex")
                rs.Close
                Dim adoCommand As New ADODB.Command
                With adoCommand
                    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                    .Parameters.Append .CreateParameter("TglHitung", adDate, adParamInput, , Format(dtpBlnHitung.Value, "yyyy/MM/dd"))
                    .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, dgPegawai.Columns(0))
                    .Parameters.Append .CreateParameter("KdDetailKomponenIndex", adVarChar, adParamInput, 6, strKdDetailKomponenIndex)
                    .Parameters.Append .CreateParameter("NilaiIndex", adInteger, adParamInput, 4, hgIndex.TextMatrix(intRow - 1, 5))
                    .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                    .Parameters.Append .CreateParameter("KdKomponenIndex", adVarChar, adParamInput, 4, strKdKomponenIndex)
                    .ActiveConnection = dbConn

                    .CommandText = "AU_HRD_HitungIndex"
                    .CommandType = adCmdStoredProc
                    .Execute

                    Call deleteADOCommandParameters(adoCommand)
                    Set adoCommand = Nothing
                End With
            End If
        Next
        'sukses:
        MsgBox "Proses Simpan sukses", vbInformation
    End With
    cmdBaru_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcDetKomIndex_LostFocus()
    dcDetKomIndex.Visible = False
End Sub

Private Sub dgPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPegawai
    WheelHook.WheelHook dgPegawai
End Sub

Private Sub dgPegawai_DblClick()
    Call cmdHitung_Click
End Sub

Private Sub dgPegawai_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    cmdsimpan.Enabled = True
    If dgPegawai.Columns(5) = "" Then MsgBox "Jabatan kosong ", vbCritical, "validasi": Exit Sub
    If dgPegawai.Columns(6) = "" Then MsgBox "Pendidikan kosong ", vbCritical, "validasi": Exit Sub
    If KeyAscii = 13 Then
        If intLJmlPegawai = 0 Then Exit Sub

        txtidpegawai.Text = dgPegawai.Columns(0)
        txtnamapegawai.Text = dgPegawai.Columns(1)
        txtJK.Text = dgPegawai.Columns(2)
        txtJnsPeg.Text = dgPegawai.Columns(3)
        txtJabatan.Text = dgPegawai.Columns(5)
        frapegawai.Visible = False
        txtidpegawai.Enabled = False
        txtnamapegawai.Enabled = False
        txtJK.Enabled = False
        txtJnsPeg.Enabled = False
        txtJabatan.Enabled = False

        Call subLoadIndexPegawai

    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcDetKomIndex_DblClick(Area As Integer)
    Call dcDetKomIndex_KeyPress(13)
End Sub

Private Sub dcDetKomIndex_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        hgIndex.TextMatrix(hgIndex.row, 2) = dcDetKomIndex.Text
        dcDetKomIndex.Visible = False
        hgIndex.Col = 3
        hgIndex.SetFocus
        If dcDetKomIndex.Text <> "" Then

            strSQL = "SELECT NilaiIndexStandar,RateIndex,JenisKomponenIndex,KomponenIndex FROM V_KomponenIndexKaryawan WHERE " _
            & "DetailKomponenIndex='" & hgIndex.TextMatrix(hgIndex.row, 2) & "' "
            Set rs = Nothing
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            hgIndex.TextMatrix(hgIndex.row, 3) = rs(0).Value
            hgIndex.TextMatrix(hgIndex.row, 4) = rs(1).Value
            hgIndex.TextMatrix(hgIndex.row, 5) = (rs(0).Value * rs(1).Value)
            hgIndex.TextMatrix(hgIndex.row, 6) = rs(3).Value
            Call subHitungIndex

            hgIndex.SetFocus
            SendKeys "{DOWN}"

        End If
    End If
    If rs("JenisKomponenIndex").Value = "Risk Index" Then
        cmdsimpan.Enabled = True
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dtpBlnHitung_Change()
    dtpBlnHitung.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subLoadDataPegawai
    frapegawai.Visible = True
    Frame2.Visible = False
    dtpBlnHitung.Value = Format(Now, "MMMM, yyyy")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDataPerhitunganIndex = Nothing
End Sub

Private Sub hgIndex_Click()
    Call hgIndex_EnterCell
End Sub

Private Sub hgIndex_EnterCell()
    On Error GoTo hunny

    With dcDetKomIndex
        If hgIndex.Col <> 2 Then
            If hgIndex.TextMatrix(hgIndex.row, 1) = "" Then
                .Visible = False
                Exit Sub
            End If
        End If
        If hgIndex.TextMatrix(hgIndex.row, 1) = "Risk Index" Then
            .Visible = True
            .Width = hgIndex.ColWidth(2)
            .Top = hgIndex.RowPos(hgIndex.row) + 370
            .Left = hgIndex.ColPos(2) + 250
            .Text = ""

            strSQL = "SELECT DetailKomponenIndex FROM V_KomponenIndexKaryawan where JenisKomponenIndex='" & hgIndex.TextMatrix(hgIndex.row, 1) & "'"
            Set rs = Nothing
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set .RowSource = rs
            .ListField = rs(0).Name
            .SetFocus
            If rs.RecordCount = 1 Then .Text = rs(0).Value
            If hgIndex.TextMatrix(hgIndex.row, 2) <> "" Then
                .Text = hgIndex.TextMatrix(hgIndex.row, 2)
            End If
        End If
        If hgIndex.TextMatrix(hgIndex.row, 1) = "Emergency Index" Then
            .Visible = True
            .Width = hgIndex.ColWidth(2)
            .Top = hgIndex.RowPos(hgIndex.row) + 370
            .Left = hgIndex.ColPos(2) + 250
            .Text = ""

            strSQL = "SELECT DetailKomponenIndex FROM V_KomponenIndexKaryawan where JenisKomponenIndex='" & hgIndex.TextMatrix(hgIndex.row, 1) & "'"
            Set rs = Nothing
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set .RowSource = rs
            .ListField = rs(0).Name
            .SetFocus
            If rs.RecordCount = 1 Then .Text = rs(0).Value
            If hgIndex.TextMatrix(hgIndex.row, 2) <> "" Then
                .Text = hgIndex.TextMatrix(hgIndex.row, 2)
            End If
        End If

    End With

    Exit Sub
hunny:
    Call msubPesanError
End Sub

Private Sub hgIndex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = 13 Then
        Call hgIndex_EnterCell
    Else
        Call hgIndex_EnterCell
        Call dcDetKomIndex_KeyPress(KeyAscii)
    End If
End Sub

Private Sub hgIndex_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then

        hgIndex.TextMatrix(hgIndex.row, 2) = ""
        hgIndex.TextMatrix(hgIndex.row, 3) = ""
        hgIndex.TextMatrix(hgIndex.row, 4) = ""
        hgIndex.TextMatrix(hgIndex.row, 5) = ""
        hgIndex.TextMatrix(hgIndex.row, 6) = ""
        Call subHitungIndex
    End If
End Sub

Private Sub dgPegawai_GotFocus()
    If dgPegawai.Col < 2 Then dgPegawai.Col = 2
End Sub

Private Sub txtNamaPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgPegawai.Visible = True Then dgPegawai.SetFocus
End Sub

Private Sub txtNamaPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dgPegawai.Visible = True Then dgPegawai.SetFocus
    End If
End Sub

Private Sub txtNamaPegawai_GotFocus()
    txtnamapegawai.SelStart = 0
    txtnamapegawai.SelLength = Len(txtnamapegawai.Text)
End Sub

Private Sub subLoadIndexPegawai()
    On Error Resume Next

    pbData.Value = 0

    strLSQL = "SELECT distinct JenisKomponenIndex,NULL AS [DETAIL KOMPONEN INDEX],NULL AS [NILAI INDEX],NULL AS BOBOT,NULL AS SCORE, NULL AS KomponenIndex, KdJenisKomponenIndex FROM V_KomponenIndexKaryawan order by KdJenisKomponenIndex"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intLJmlIndex = rs.RecordCount
    With hgIndex
        Set .DataSource = rs
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 2500
        .ColWidth(3) = 2500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 0
        .ColWidth(7) = 0

        strQuery = "SELECT DISTINCT dbo.KualifikasiJurusan.KdPendidikan, dbo.DataCurrentPegawai.KdJabatan" & _
        " FROM dbo.KualifikasiJurusan INNER JOIN dbo.DataCurrentPegawai ON dbo.KualifikasiJurusan.KdKualifikasiJurusan = dbo.DataCurrentPegawai.KdKualifikasiJurusan" & _
        " where dbo.DataCurrentPegawai.IdPegawai = '" & txtidpegawai.Text & "' " & _
        " GROUP BY dbo.KualifikasiJurusan.KdPendidikan, dbo.DataCurrentPegawai.KdJabatan"

        Call msubRecFO(rs2, strQuery)
        If IsNull(rs2.Fields("KdPendidikan")) Then
        Else
            strKdPendidikan = rs2.Fields("KdPendidikan").Value
        End If

        If IsNull(rs2.Fields("KdJabatan")) Then
        Else
            strKdJabatan = rs2.Fields("KdJabatan").Value
        End If

        Frame2.Visible = True
        hgIndex.Col = 2
        hgIndex.row = 3
        hgIndex.SetFocus

        pbData.Value = 0
        pbData.Max = rs.RecordCount

        For j = 1 To hgIndex.Rows - 1

            If rs(0).Value = "Competency Index" Then

                strsqlx = "SELECT NilaiIndexStandar,RateIndex,JenisKomponenIndex,KomponenIndex,DetailKomponenIndex FROM V_KomponenIndexKaryawan WHERE " _
                & "kdpendidikan='" & strKdPendidikan & "' "
                Set rsx = Nothing
                rsx.Open strsqlx, dbConn, adOpenForwardOnly, adLockReadOnly
                hgIndex.TextMatrix(j, 3) = rsx(0).Value
                hgIndex.TextMatrix(j, 4) = rsx(1).Value
                hgIndex.TextMatrix(j, 5) = (rsx(0).Value * rsx(1).Value)
                hgIndex.TextMatrix(j, 6) = rsx(3).Value
                hgIndex.TextMatrix(j, 2) = rsx("DetailKomponenIndex").Value
                Call subHitungIndex
            End If
            If rs(0).Value = "Position Index" Then

                strsqlx = "SELECT NilaiIndexStandar,RateIndex,JenisKomponenIndex,KomponenIndex,DetailKomponenIndex FROM V_KomponenIndexKaryawan WHERE " _
                & "kdjabatan='" & strKdJabatan & "' "
                Set rsx = Nothing
                rsx.Open strsqlx, dbConn, adOpenForwardOnly, adLockReadOnly
                hgIndex.TextMatrix(j, 3) = rsx(0).Value
                hgIndex.TextMatrix(j, 4) = rsx(1).Value
                hgIndex.TextMatrix(j, 5) = (rsx(0).Value * rsx(1).Value)
                hgIndex.TextMatrix(j, 6) = rsx(3).Value
                hgIndex.TextMatrix(j, 2) = rsx("DetailKomponenIndex").Value
                Call subHitungIndex
            End If
            If rs(0).Value = "Performance Index" Then

                strsqlx = "SELECT NilaiIndexStandar,RateIndex,JenisKomponenIndex,KomponenIndex,DetailKomponenIndex FROM V_KomponenIndexKaryawan WHERE " _
                & "JenisKomponenIndex='Performance Index' "
                Set rsx = Nothing
                rsx.Open strsqlx, dbConn, adOpenForwardOnly, adLockReadOnly
                hgIndex.TextMatrix(j, 3) = CInt(rsxx("Jml").Value / 100000) * 2
                hgIndex.TextMatrix(j, 4) = rsx(1).Value
                hgIndex.TextMatrix(j, 5) = (hgIndex.TextMatrix(j, 3) * rsx(1).Value)
                hgIndex.TextMatrix(j, 6) = rsx(3).Value
                hgIndex.TextMatrix(j, 2) = rsx("DetailKomponenIndex").Value
                Call subHitungIndex
            End If
            If rs(0).Value = "Basic Index" Then

                strsqlx = "SELECT NilaiIndexStandar,RateIndex,JenisKomponenIndex,KomponenIndex,DetailKomponenIndex FROM V_KomponenIndexKaryawan WHERE " _
                & "JenisKomponenIndex='Basic Index' "
                Set rsx = Nothing
                rsx.Open strsqlx, dbConn, adOpenForwardOnly, adLockReadOnly

                strsqlxx = "SELECT MAX(NoUrut) AS NoUrut, MAX(Jumlah) AS Jml FROM DetailRiwayatGaji where idpegawai='" & txtidpegawai.Text & "' and kdkomponengaji='01'"
                Set rsxx = Nothing
                rsxx.Open strsqlxx, dbConn, adOpenForwardOnly, adLockReadOnly

                Dim BasicJml As Integer
                If IsNull(rsxx("Jml")) Then BasicJml = 0
                hgIndex.TextMatrix(j, 3) = (BasicJml / 100000)
                hgIndex.TextMatrix(j, 4) = rsx(1).Value
                hgIndex.TextMatrix(j, 5) = funcRound(hgIndex.TextMatrix(j, 3) * hgIndex.TextMatrix(j, 4), 1)
                hgIndex.TextMatrix(j, 6) = rsx(3).Value
                hgIndex.TextMatrix(j, 2) = rsx("DetailKomponenIndex").Value
                Call subHitungIndex
            End If
            rs.MoveNext
        Next j

        .RowHeight(0) = dcDetKomIndex.Height + 100
        .RowHeightMin = dcDetKomIndex.Height + 10

    End With
End Sub

Private Sub subHitungIndex()
    txtTotalHasil.Text = 0
    For i = 1 To hgIndex.Rows - 1
        If hgIndex.TextMatrix(i, 5) = "" Then GoTo nexti
        txtTotalHasil.Text = CInt(txtTotalHasil.Text) + CInt(hgIndex.TextMatrix(i, 5))
nexti:
    Next i
End Sub

Private Sub subLoadDataPegawai()
    On Error Resume Next
    Set rs = Nothing
    strLSQL = "SELECT * FROM v_S_Pegawai order by NamaLengkap "
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intLJmlPegawai = rs.RecordCount
    With dgPegawai
        Set .DataSource = rs
        .Columns(0).Width = 1500
        .Columns(1).Width = 3000
        .Columns(2).Width = 700
        .Columns(3).Width = 1700
        .Columns(4).Width = 0
        .Columns(5).Width = 2500
        .Columns(6).Width = 0
        .Columns(7).Width = 1500
        .Columns(0).Caption = "ID PEGAWAI"
        .Columns(1).Caption = "NAMA LENGKAP"
        .Columns(2).Caption = "SEX"
        .Columns(3).Caption = "JENIS PEGAWAI"
        .Columns(4).Caption = "KD JABATAN"
        .Columns(5).Caption = "NAMA JABATAN"
        .Columns(6).Caption = "KD PENDIDIKAN"
        .Columns(7).Caption = "PENDIDIKAN"
    End With
End Sub

Sub kosong()
    txtnamapegawai.Text = ""
    txtJK.Text = ""
    txtJnsPeg.Text = ""
    txtJabatan.Text = ""
    txtidpegawai.Enabled = True
    txtnamapegawai.Enabled = True
    txtJK.Enabled = True
    txtJnsPeg.Enabled = True
    txtJabatan.Enabled = True
End Sub
