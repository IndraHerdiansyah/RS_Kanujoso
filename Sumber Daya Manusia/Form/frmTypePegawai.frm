VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmTypePegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Type Pegawai"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTypePegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   8115
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
   End
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
      Height          =   5775
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   7920
      Begin VB.CheckBox chkStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         Height          =   255
         Left            =   5880
         TabIndex        =   4
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtNamaExt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1800
         Width           =   5055
      End
      Begin VB.TextBox txtKdExt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtRepDisplay 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtKdType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   5055
      End
      Begin MSDataGridLib.DataGrid dgType 
         Height          =   3330
         Left            =   255
         TabIndex        =   6
         Top             =   2280
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   5874
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama external"
         Height          =   210
         Left            =   480
         TabIndex        =   18
         Top             =   1860
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   210
         Left            =   480
         TabIndex        =   17
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Report Display"
         Height          =   210
         Left            =   480
         TabIndex        =   16
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Type Pegawai"
         Height          =   210
         Left            =   480
         TabIndex        =   15
         Top             =   735
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode Type"
         Height          =   210
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   900
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
   End
   Begin VB.Image Image4 
      Height          =   945
      Left            =   6240
      Picture         =   "frmTypePegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTypePegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTypePegawai.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmTypePegawai.frx":5A71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmTypePegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub blankfield()
    On Error GoTo hell
    txtKdType.Text = ""
    txtType.Text = ""
    txtRepDisplay.Text = ""
    txtKdExt.Text = ""
    txtNamaExt.Text = ""
    chkStatus.Value = 1
hell:
End Sub

Sub Dag()
    On Error GoTo hell
    strSQL = "SELECT KdTypePegawai , TypePegawai , ReportDisplay , KodeExternal , NamaExternal , StatusEnabled  FROM TypePegawai" 'WHERE (StatusEnabled <> 0) OR (StatusEnabled IS NULL)"
    Call msubRecFO(rs, strSQL)
    Set dgType.DataSource = rs
    dgType.Columns(0).Width = 1500
    dgType.Columns(0).Caption = "Kode"
    dgType.Columns(1).Width = 2500
    dgType.Columns(1).Caption = "Type"
    dgType.Columns(2).Width = 2000
    dgType.Columns(2).Caption = "Report"
    dgType.Columns(3).Width = 1500
    dgType.Columns(3).Caption = "Kode External"
    dgType.Columns(4).Width = 2000
    dgType.Columns(4).Caption = "Nama External"
    dgType.Columns(5).Width = 1500
    dgType.Columns(5).Caption = "Status"
hell:
End Sub

Private Function sp_TypePegawai(f_Status As String) As Boolean
    sp_TypePegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdTypePegawai", adChar, adParamInput, 2, IIf(txtKdType.Text = "", Null, txtKdType.Text))
        .Parameters.Append .CreateParameter("TypePegawai", adVarChar, adParamInput, 30, Trim(txtType.Text))
        .Parameters.Append .CreateParameter("ReportDisplay", adVarChar, adParamInput, 30, IIf(txtRepDisplay.Text = "", Null, Trim(txtRepDisplay.Text)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatus.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_TypePegawai"
        .CommandType = adCmdStoredProc
        .Execute
        Call Dag

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_TypePegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        Call blankfield
    End With
End Function

Private Sub cmdCancel_Click()
    On Error GoTo hell
    Call blankfield
    Exit Sub
hell:
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    If dgType.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakTypePegawai.Show
hell:
End Sub

Private Sub cmdDel_Click()
    On Error GoTo hell

    If MsgBox("Yakin akan menghapus data ini?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    If Periksa("text", txtType, "Nama type pegawai kosong") = False Then Exit Sub
    If Periksa("text", txtRepDisplay, "Report Display kosong") = False Then Exit Sub
    If sp_TypePegawai("D") = False Then Exit Sub

    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"

hell:
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errload

    If Periksa("text", txtType, "Silahkan isi nama type pegawai ") = False Then Exit Sub
    If Periksa("text", txtRepDisplay, "Silahkan isi Report Display ") = False Then Exit Sub
    If sp_TypePegawai("A") = False Then Exit Sub

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"

    Exit Sub
Errload:
End Sub

Private Sub dgType_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgType
    WheelHook.WheelHook dgType
End Sub

Private Sub dgType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtType.SetFocus
End Sub

Private Sub dgType_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo hell
    If dgType.ApproxCount = 0 Then Exit Sub
    txtKdType.Text = dgType.Columns(0).Value
    txtType.Text = dgType.Columns(1)
    If IsNull(dgType.Columns(2)) Then txtRepDisplay.Text = "" Else txtRepDisplay.Text = dgType.Columns(2)
    If IsNull(dgType.Columns(3)) Then txtKdExt.Text = "" Else txtKdExt.Text = dgType.Columns(3)
    If IsNull(dgType.Columns(4)) Then txtNamaExt.Text = "" Else txtNamaExt.Text = dgType.Columns(4)
    chkStatus.Value = dgType.Columns(5).Value
Exit Sub
hell:
End Sub

Private Sub Form_Activate()
    txtType.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call blankfield
    Call Dag

End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtNamaExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtRepDisplay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtType_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgType.SetFocus
    End Select
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRepDisplay.SetFocus
End Sub
