VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmStatusAbsensi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Status Absensi Pegawai"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatusAbsensi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5055
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
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
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtKodeExternal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtNamaExternal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CheckBox CheckStatusEnbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Status Enabled"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtKdStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtStatusAbsensi 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
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
      Left            =   3960
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdbatal 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdsimpan 
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
      Left            =   3000
      TabIndex        =   9
      Top             =   5880
      Width           =   975
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
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
   Begin MSDataGridLib.DataGrid dgStatusAbsensi 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Kode External"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nama External"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status Absensi:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kode:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStatusAbsensi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3600
      Picture         =   "frmStatusAbsensi.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStatusAbsensi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmStatusAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCommand As New ADODB.Command

Private Sub CheckStatusEnbl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub cmdBatal_Click()
    txtStatusAbsensi.Enabled = True
    txtStatusAbsensi.SetFocus
    clear
    tampilData
    cmdHapus.Enabled = False
    cmdSimpan.Enabled = True
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell:
    If dgStatusAbsensi.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakStatusAbsensi.Show
hell:
End Sub

Private Sub cmdHapus_Click()
    If MsgBox("Yakin data ini akan dihapus", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    With adoCommand
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AUD_StatusAbsensiPegawai"

        .Parameters.Append .CreateParameter("KdStatusAbsensi", adChar, adParamInput, 2, Trim(txtKdStatus.Text))
        .Parameters.Append .CreateParameter("StatusAbsensi", adVarChar, adParamInput, 30, Trim(txtStatusAbsensi.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.Value)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, "D")
        .Execute

        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    tampilData
    clear
    MsgBox "Data berhasil dihapus", vbInformation, "Validasi"
    cmdHapus.Enabled = False
    cmdSimpan.Enabled = True
    txtStatusAbsensi.SetFocus
End Sub

Private Sub cmdSimpan_Click()
    If Periksa("text", txtStatusAbsensi, "Silahkan isi status absensi") = False Then Exit Sub
    'If Periksa("datacombo", dcJK, "Silahkan isi jenis kelamin") = False Then Exit Sub
    With adoCommand
        If txtKdStatus.Text = "" Then
            .Parameters.Append .CreateParameter("KdStatusAbsensi", adChar, adParamInput, 2, Null)
        Else
            .Parameters.Append .CreateParameter("KdStatusAbsensi", adChar, adParamInput, 2, Trim(txtKdStatus.Text))
        End If
        .Parameters.Append .CreateParameter("StatusAbsensi", adVarChar, adParamInput, 30, Trim(txtStatusAbsensi.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.Value)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, "A")

        .ActiveConnection = dbConn
        .CommandText = "AUD_StatusAbsensiPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    tampilData
    clear
    txtKdStatus.Enabled = False
    txtStatusAbsensi.Enabled = False
    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
    cmdBatal_Click
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgStatusAbsensi_Click()
    On Error Resume Next
    'Cara isi kolim kode extrnal-statusenabled
    If dgStatusAbsensi.ApproxCount = 0 Then Exit Sub
    cmdHapus.Enabled = True
    txtKdStatus.Text = dgStatusAbsensi.Columns(0).Value
    txtStatusAbsensi.Text = dgStatusAbsensi.Columns(1).Value
    If dgStatusAbsensi.Columns("KodeExternal").Value = Null Then
        txtKodeExternal.Text = ""
    Else
        txtKodeExternal.Text = dgStatusAbsensi.Columns("KodeExternal").Value
    End If
    If dgStatusAbsensi.Columns("KodeExternal").Value = Null Then
        txtNamaExternal.Text = ""
    Else
        txtNamaExternal.Text = dgStatusAbsensi.Columns("NamaExternal").Value
    End If
    CheckStatusEnbl.Value = dgStatusAbsensi.Columns("StatusEnabled").Value
    txtKdStatus.Enabled = False
    txtStatusAbsensi.Enabled = True

    WheelHook.WheelUnHook
    Set MyProperty = dgStatusAbsensi
    WheelHook.WheelHook dgStatusAbsensi
End Sub

Private Sub dgStatusAbsensi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad
    If dgStatusAbsensi.ApproxCount = 0 Then Exit Sub
    txtKdStatus.Text = dgStatusAbsensi.Columns(0).Value
    txtStatusAbsensi.Text = dgStatusAbsensi.Columns(1).Value
    If dgStatusAbsensi.Columns("KodeExternal").Value = Null Then
        txtKodeExternal.Text = ""
    Else
        txtKodeExternal.Text = dgStatusAbsensi.Columns("KodeExternal").Value
    End If
    If dgStatusAbsensi.Columns("KodeExternal").Value = Null Then
        txtNamaExternal.Text = ""
    Else
        txtNamaExternal.Text = dgStatusAbsensi.Columns("NamaExternal").Value
    End If
    txtKdStatus.Enabled = False
    txtStatusAbsensi.Enabled = True
    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    cmdHapus.Enabled = False
    tampilData
End Sub

Sub clear()
    txtKdStatus.Text = ""
    txtStatusAbsensi.Text = ""
    txtKodeExternal.Text = ""
    txtNamaExternal.Text = ""
    CheckStatusEnbl.Value = 1
End Sub

Sub tampilData()
    Set rs = Nothing
    strSQL = "select * from StatusAbsensi"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgStatusAbsensi.DataSource = rs
    dgStatusAbsensi.Columns(0).Width = 750
    dgStatusAbsensi.Columns(1).Width = 3000
    dgStatusAbsensi.Columns(0).Alignment = dbgCenter
    dgStatusAbsensi.Columns(0).Caption = "Kode"
    dgStatusAbsensi.Columns(1).Caption = "Status Absensi"
    dgStatusAbsensi.Columns(5).Width = 1150
    Set rs = Nothing
End Sub

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl.SetFocus
End Sub

Private Sub txtstatusabsensi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub
