VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTitle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Title"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTitle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   6135
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   5880
         Width           =   975
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   5880
         Width           =   975
      End
      Begin VB.TextBox txtTit 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   6
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtKdTit 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtKdExt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtNmExt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CheckBox chkSts 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dgTit 
         Height          =   3735
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kode Title"
         Height          =   210
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   210
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   1275
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTitle.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4440
      Picture         =   "frmTitle.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTitle.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCommand As New ADODB.Command

Private Sub cmdBatal_Click()
    clear
    tampilData
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    If dgTit.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakTitle.Show
hell:
End Sub

Private Sub cmdHapus_Click()
'    On Error GoTo xxx
'    If MsgBox("Hapus Data ini? ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
'    dbConn.Execute "Delete Title WHERE KdTitle = '" & txtKdTit.Text & "'"
'    tampilData
'    clear
'    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
'    Exit Sub
'xxx:
'    MsgBox "Data digunakan, tidak dapat dihapus..", vbOKOnly, "Validasi"

On Error GoTo errLoad
    If Periksa("text", txtTit, "Pilih Data yang akan dihapus") = False Then Exit Sub
    If MsgBox("Hapus Data ini? ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "Select * from DataPegawai where KdTitle='" & txtKdTit & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
        Exit Sub
    Else
        dbConn.Execute "Delete Title WHERE KdTitle = '" & txtKdTit.Text & "'"
        tampilData
        clear
        MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
        Exit Sub
    End If
errLoad:
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell

    If Periksa("text", txtTit, "Silahkan isi nama title ") = False Then Exit Sub

    With adoCommand
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AUD_Title"

        If txtKdTit.Text = "" Then
            .Parameters.Append .CreateParameter("KdTitle", adChar, adParamInput, 2, Null)
        Else
            .Parameters.Append .CreateParameter("KdTitle", adChar, adParamInput, 2, txtKdTit.Text)
        End If

        .Parameters.Append .CreateParameter("NamaTitle", adVarChar, adParamInput, 20, Trim(txtTit.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExt.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExt.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, "A")

        If Not IsNull(.Parameters("KdTitle").Value) Then
            '' Update
            MsgBox "Data berhasil di update ..", vbInformation, "Informasi"
            .Execute
        Else
            '' Input data baru
            MsgBox "Data berhasil disimpan.. ", vbInformation, "Informasi"
            .Execute
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    cmdBatal_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgTit_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTit
    WheelHook.WheelHook dgTit
End Sub

Private Sub dgTit_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgTit.ApproxCount = 0 Then Exit Sub
    txtKdTit.Text = dgTit.Columns(0).Value
    txtTit.Text = dgTit.Columns(1).Value
    txtKdExt.Text = dgTit.Columns(2).Value
    txtNmExt.Text = dgTit.Columns(3).Value
    If dgTit.Columns(4).Value = "<Type mismacth>" Then
        chkSts.Value = 0
    Else
        If dgTit.Columns(4).Value = 1 Then
            chkSts.Value = 1
        Else
            chkSts.Value = 0
        End If
    End If

    txtKdTit.Enabled = False
    txtTit.Enabled = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call cmdBatal_Click

    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub clear()
    On Error Resume Next
    txtKdTit.Text = ""
    txtTit.Text = ""
    txtKdExt = ""
    txtNmExt = ""
    chkSts.Value = 1
    txtTit.SetFocus
End Sub

Sub tampilData()
    On Error GoTo hell
    Set rs = Nothing
    strSQL = "select * from Title order by KdTitle "
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTit.DataSource = rs
    dgTit.Columns(0).Width = 750
    dgTit.Columns(0).Caption = "Kode"
    dgTit.Columns(1).Width = 1500
    dgTit.Columns(1).Caption = "Title"
    Set rs = Nothing
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt.SetFocus
End Sub

Private Sub txtNmExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtTit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub
