VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterFungsionalPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Fungsional Pegawai"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "frmMasterFungsionalPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8400
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
      Left            =   5160
      TabIndex        =   1
      Top             =   6720
      Width           =   1455
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
      Left            =   3600
      TabIndex        =   5
      Top             =   6720
      Width           =   1455
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
      Left            =   6720
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
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
      Left            =   2040
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Detail Jenis Jabatan Fungsional"
      TabPicture(0)   =   "frmMasterFungsionalPegawai.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detail Jabatan Fungsional"
      TabPicture(1)   =   "frmMasterFungsionalPegawai.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   12
         Top             =   540
         Width           =   7575
         Begin VB.TextBox txtKdJabatan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
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
            Height          =   315
            Left            =   240
            MaxLength       =   5
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtNamaJabatan 
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
            Height          =   315
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   13
            Top             =   480
            Width           =   3855
         End
         Begin MSDataListLib.DataCombo dcJenisJabatan 
            Height          =   330
            Left            =   5160
            TabIndex        =   15
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataGridLib.DataGrid dgJabatan 
            Height          =   3135
            Left            =   240
            TabIndex        =   16
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
               Size            =   8.25
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Detail Jenis Jabatan"
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
            Left            =   5160
            TabIndex        =   19
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nama Jabatan"
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
            Left            =   1200
            TabIndex        =   18
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
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
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4335
         Left            =   360
         TabIndex        =   6
         Top             =   540
         Width           =   7575
         Begin VB.TextBox txtJenisJabatan 
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
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   8
            Top             =   480
            Width           =   6015
         End
         Begin VB.TextBox txtKdJenisJabatan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
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
            Height          =   315
            Left            =   240
            MaxLength       =   3
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid dgJenisJabatan 
            Height          =   3135
            Left            =   240
            TabIndex        =   9
            Top             =   960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Detail Jenis Jabatan"
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
            Left            =   1320
            TabIndex        =   11
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
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
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   420
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
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
      Left            =   6550
      Picture         =   "frmMasterFungsionalPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterFungsionalPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterFungsionalPegawai.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMasterFungsionalPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDcSource()
    strSQL = "SELECT * FROM DetailJenisJabatanF order by DetailJenisJabatanF"
    Call msubDcSource(dcJenisJabatan, rs, strSQL)
End Sub

Private Function sp_simpan(f_Status As String) As Boolean
    On Error GoTo errLoad
    Select Case sstDataPenunjang.Tab

        Case 0 '
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdDetailJenisJabatanF", adChar, adParamInput, 2, Trim(txtKdJenisJabatan))
                .Parameters.Append .CreateParameter("DetailJenisJabatanF", adVarChar, adParamInput, 50, Trim(txtJenisJabatan))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_DetailJenisJabatanF"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

        Case 1
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdDetailJabatanF", adChar, adParamInput, 5, Trim(txtKdJabatan))
                .Parameters.Append .CreateParameter("DetailJabatanF", adVarChar, adParamInput, 50, Trim(txtNamaJabatan))
                .Parameters.Append .CreateParameter("KdDetailJenisJabatanF", adChar, adParamInput, 2, Trim(dcJenisJabatan.BoundText))
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_DetailJabatanFungsional"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            If (f_Status = "A") Then
                MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
            Else
                MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
            End If
            cmdBatal_Click

    End Select
    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Sub cmdBatal_Click()
    Call subLoadGridSource
    Select Case sstDataPenunjang.Tab

        Case 0 ' jabatan
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 1 ' Jenis jabatan
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
    End Select
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo hell
    If MsgBox("Apakah anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case sstDataPenunjang.Tab

        Case 0
            If Periksa("text", txtJenisJabatan, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub

        Case 1
            If Periksa("datacombo", dcJenisJabatan, "Pilih Data yang akan dihapus") = False Then Exit Sub
            If Periksa("text", txtNamaJabatan, "Isi Nama jabatan") = False Then Exit Sub
            If sp_simpan("D") = False Then Exit Sub
    End Select
    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan
    Select Case sstDataPenunjang.Tab

        Case 0
            If Periksa("text", txtJenisJabatan, "Isi jenis jabatan!") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub

        Case 1
            If Periksa("datacombo", dcJenisJabatan, "Isi Jenis Jabatan!") = False Then Exit Sub
            If Periksa("text", txtNamaJabatan, "Isi Nama jabatan") = False Then Exit Sub
            If sp_simpan("A") = False Then Exit Sub
    End Select
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    Exit Sub
errSimpan:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dgJabatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJabatan.Text = dgJabatan.Columns(0).Value
    dcJenisJabatan.Text = IIf(dgJabatan.Columns(2).Value = Null, "", dgJabatan.Columns(2).Value)
    txtNamaJabatan.Text = dgJabatan.Columns(1).Value
End Sub

Private Sub dgJenisJabatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJenisJabatan.Text = dgJenisJabatan.Columns(0).Value
    txtJenisJabatan.Text = dgJenisJabatan.Columns(1).Value
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subDcSource
    sstDataPenunjang.Tab = 0
    Call subLoadGridSource
End Sub

Sub subKosong()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab

        Case 0 'Jenis jabatan
            txtKdJenisJabatan.Text = ""
            txtJenisJabatan.Text = ""
            txtJenisJabatan.SetFocus

        Case 1 'jabatan
            txtKdJabatan.Text = ""
            dcJenisJabatan.Text = ""
            txtNamaJabatan.Text = ""
            txtNamaJabatan.SetFocus
    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call subDcSource
    Call subLoadGridSource
    Call cmdBatal_Click
End Sub

Private Sub txtJenisJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKdJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaJabatan.SetFocus
End Sub

Sub subLoadGridSource()
    On Error GoTo errLoad
    Select Case sstDataPenunjang.Tab

        Case 0
            Set rs = Nothing
            strSQL = "select * from DetailJenisJabatanF"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisJabatan.DataSource = rs
            dgJenisJabatan.Columns(0).DataField = rs(0).Name
            dgJenisJabatan.Columns(1).DataField = rs(1).Name
            dgJenisJabatan.Columns(0).Width = 1250
            dgJenisJabatan.Columns(0).Caption = "Kode"
            dgJenisJabatan.Columns(1).Width = 5500
            dgJenisJabatan.Columns(1).Caption = "Jenis Jabatan"
            Set rs = Nothing

        Case 1
            Set rs = Nothing
            strSQL = "SELECT dbo.DetailJabatanFungsional.KdDetailJabatanF, dbo.DetailJabatanFungsional.DetailJabatanF, dbo.DetailJenisJabatanF.DetailJenisJabatanF " & _
            " FROM dbo.DetailJabatanFungsional LEFT OUTER JOIN" & _
            " dbo.DetailJenisJabatanF ON dbo.DetailJabatanFungsional.KdDetailJenisJabatanF = dbo.DetailJenisJabatanF.KdDetailJenisJabatanF"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJabatan.DataSource = rs
            dgJabatan.Columns(0).DataField = rs(0).Name
            dgJabatan.Columns(1).DataField = rs(1).Name
            dgJabatan.Columns(2).DataField = rs(2).Name
            dgJabatan.Columns(0).Width = 1250
            dgJabatan.Columns(0).Caption = "Kode"
            dgJabatan.Columns(1).Width = 3900
            dgJabatan.Columns(1).Caption = "Nama Jabatan Fungsional"
            dgJabatan.Columns(2).Width = 1800
            dgJabatan.Columns(2).Caption = "Jenis Jabatan Fungsional"

            Set rs = Nothing
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub txtNamaJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisJabatan.SetFocus
End Sub
