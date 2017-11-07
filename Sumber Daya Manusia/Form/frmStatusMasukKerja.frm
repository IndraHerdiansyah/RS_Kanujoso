VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Begin VB.Form frmStatusMasukKerja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Status Masuk Kerja Pegawai"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatusMasukKerja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7410
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   5400
      Width           =   7335
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdubah 
         Caption         =   "&Ubah"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   7335
      Begin VB.TextBox txtStatusMasuk 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtKdStatus 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgStatusMasuk 
         Height          =   2775
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4895
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kode Status Masuk"
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status Masuk"
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1080
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStatusMasukKerja.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5520
      Picture         =   "frmStatusMasukKerja.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStatusMasukKerja.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmStatusMasukKerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCommand As New ADODB.Command

Private Sub cmdBatal_Click()
    txtStatusMasuk.Enabled = True
    txtStatusMasuk.SetFocus
    clear
    cmdHapus.Enabled = False
    cmdUbah.Enabled = False
    cmdSimpan.Enabled = True
End Sub

Private Sub cmdHapus_Click()
        With adoCommand
            .ActiveConnection = dbConn
            .CommandType = adCmdStoredProc
            .CommandText = "AUD_StatusMasukKerja"
            
            .Parameters.Append .CreateParameter("KdStatusMasuk", adChar, adParamInput, 2, txtKdStatus.Text)
            .Parameters.Append .CreateParameter("StatusMasuk", adVarChar, adParamInput, 30, Trim(txtStatusMasuk.Text))
            .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
            MsgBox "Data berhasil dihapus", vbInformation, "Validasi"
            .Execute
            
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            
        End With
    tampilData
    clear
    cmdHapus.Enabled = False
    cmdSimpan.Enabled = True
    txtStatusMasuk.SetFocus
End Sub

Private Sub cmdSimpan_Click()
        With adoCommand
            .ActiveConnection = dbConn
            .CommandType = adCmdStoredProc
            .CommandText = "AUD_StatusMasukKerja"
            
            If txtKdStatus.Text = "" Then
                .Parameters.Append .CreateParameter("KdStatusMasuk", adChar, adParamInput, 2, Null)
            Else
                .Parameters.Append .CreateParameter("KdStatusMasuk", adChar, adParamInput, 2, txtKdStatus.Text)
            End If
            
            .Parameters.Append .CreateParameter("StatusMasuk", adVarChar, adParamInput, 30, Trim(txtStatusMasuk.Text))
            .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
            
            If Not IsNull(.Parameters("KdStatusMasuk").Value) Then
                '' Update
                MsgBox "Update data berhasil", vbInformation, "Informasi"
                .Execute
            Else
                '' Input data baru
                MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
                .Execute
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            
        End With
    tampilData
    clear
    txtKdStatus.Enabled = False
    txtStatusMasuk.Enabled = False
    cmdBatal_Click
End Sub

Private Sub cmdsimpan_GotFocus()
    If Periksa("text", txtStatusMasuk, "Nama Jenis Registrasi Harus Diisi!") = False Then Exit Sub
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdubah_Click()
    Set rs = Nothing
    strSQL = "update JenisRegistrasi set JenisRegistrasi = '" & txtStatusMasuk.Text & "' where KdJenisRegistrasi = '" & txtKdStatus.Text & "'"
    rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set rs = Nothing
    tampilData
End Sub

Private Sub dgStatusMasuk_Click()
    If dgStatusMasuk.ApproxCount = 0 Then Exit Sub
    cmdHapus.Enabled = True
    txtKdStatus.Text = dgStatusMasuk.Columns(0).Value
    txtStatusMasuk.Text = dgStatusMasuk.Columns(1).Value
    txtKdStatus.Enabled = False
    txtStatusMasuk.Enabled = True
End Sub

Private Sub dgStatusMasuk_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errload
    If dgStatusMasuk.ApproxCount = 0 Then Exit Sub
    txtKdStatus.Text = dgStatusMasuk.Columns(0).Value
    txtStatusMasuk.Text = dgStatusMasuk.Columns(1).Value
    txtKdStatus.Enabled = False
    txtStatusMasuk.Enabled = True
Exit Sub
errload:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    cmdHapus.Enabled = False
    cmdUbah.Enabled = False

    Set rs = Nothing
    strSQL = "select * from StatusMasukKerja"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgStatusMasuk.DataSource = rs
        dgStatusMasuk.Columns(0).Width = 2000
        dgStatusMasuk.Columns(1).Width = 4000
    Set rs = Nothing
End Sub

Sub clear()
    txtKdStatus.Text = ""
    txtStatusMasuk.Text = ""
End Sub

Sub tampilData()
    Set rs = Nothing
    strSQL = "select * from StatusMasukKerja"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgStatusMasuk.DataSource = rs
        dgStatusMasuk.Columns(0).Width = 2000
        dgStatusMasuk.Columns(1).Width = 4000
    Set rs = Nothing
End Sub

Private Sub txtStatusMasuk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtStatusMasuk_LostFocus()
    txtStatusMasuk.Text = StrConv(txtStatusMasuk.Text, vbProperCase)
End Sub
