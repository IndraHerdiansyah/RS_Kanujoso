VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmConvertFingerPrintToIDPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Konversi FingerPrint To ID Pegawai"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConvertFingerPrintToIDPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   6795
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
      Top             =   7440
      Width           =   6735
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "F1 - Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1455
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
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6735
      Begin VB.TextBox txtCariPegawai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   6120
         Width           =   2895
      End
      Begin VB.TextBox txtIdPegawai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3600
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame fraDataPegawai 
         Caption         =   "Daftar Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         Begin MSDataGridLib.DataGrid dgDaftarPegawai 
            Height          =   4455
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   7858
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
      End
      Begin VB.TextBox txtNoAbsensi 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4920
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPegawai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin MSDataGridLib.DataGrid dgConvertFingerToIDpegawai 
         Height          =   4695
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8281
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pegawai"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai"
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No Absensi"
         Height          =   210
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   900
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1800
      _cx             =   4197479
      _cy             =   4196024
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
      Picture         =   "frmConvertFingerPrintToIDPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4920
      Picture         =   "frmConvertFingerPrintToIDPegawai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmConvertFingerPrintToIDPegawai.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmConvertFingerPrintToIDPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCommand As New ADODB.Command
Private Function sp_ConvertFingerToIDPegawai(f_Status) As Boolean
Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIdPegawai.Text)
        .Parameters.Append .CreateParameter("PR_FingerID", adInteger, adParamInput, 8, txtNoAbsensi.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_HRD_ConvertFingerToIDPegawai"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan No Absensi Pegawai", vbCritical
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            Exit Function
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Call subLoadGridSource
    Call clear
    txtNamaPegawai.SetFocus
    fraDataPegawai.Visible = False
End Function
'untuk meload data pegawai di grid
Private Sub subLoadDataPegawai()
    On Error Resume Next
    strSQL = "SELECT IdPegawai, NamaLengkap FROM DataPegawai WHERE NamaLengkap like '" & txtNamaPegawai.Text & "%'ORDER BY NamaLengkap "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
  
    With dgDaftarPegawai
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
    End With
    fraDataPegawai.Left = 140
    fraDataPegawai.Top = 900
End Sub
Private Sub cmdBatal_Click()
    Call clear
End Sub

Private Sub cmdHapus_Click()
If MsgBox("Yakin No Absensi " & vbCrLf & dgDaftarPegawai.Columns(1).Value & "  Akan Dihapus??", vbQuestion + vbYesNo, "Validasi") = vbNo Then Exit Sub
If Periksa("text", txtNamaPegawai, "Pilih Nama Pegawai yang akan dihapus!!") = False Then Exit Sub
        If sp_ConvertFingerToIDPegawai("D") = False Then Exit Sub
End Sub

Private Sub cmdSimpan_Click()
    If Periksa("text", txtNamaPegawai, "Nama pegawai harus diisi!") = False Then Exit Sub
    If Periksa("text", txtNoAbsensi, "No Absensi harus diisi!") = False Then Exit Sub
    strSQL = "SELECT PR_FingerID From ConvertFingerToDataPegawai WHERE PR_FingerID ='" & txtNoAbsensi.Text & "'  "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        MsgBox "No Absensi Sudah Ada!!", vbCritical, "Validasi"
        txtNoAbsensi.Text = ""
        txtNoAbsensi.SetFocus
        Exit Sub
    End If
    
    
    
    If sp_ConvertFingerToIDPegawai("A") = False Then Exit Sub
    
End Sub
Private Sub cmdTutup_Click()
    Unload Me
End Sub


Private Sub dgConvertFingerToIDpegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    txtIdPegawai.Text = dgConvertFingerToIDpegawai.Columns(0).Value
    txtNoAbsensi.Text = dgConvertFingerToIDpegawai.Columns(1).Value
    txtNamaPegawai.Text = dgConvertFingerToIDpegawai.Columns(2).Value
    fraDataPegawai.Visible = False
End Sub

Private Sub dgDaftarPegawai_DblClick()
    Call dgDaftarPegawai_KeyPress(13)
End Sub

Private Sub dgDaftarPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtIdPegawai.Text = dgDaftarPegawai.Columns(0).Value
        txtNamaPegawai.Text = dgDaftarPegawai.Columns(1).Value
        
        If txtNamaPegawai.Text = "" Then
            MsgBox "Pilih dulu nama pegawai", vbCritical, "Validasi"
            txtNamaPegawai.Text = ""
            dgDaftarPegawai.SetFocus
            Exit Sub
        End If
        fraDataPegawai.Visible = False
        txtNoAbsensi.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDataPegawai.Visible = False
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    txtCariPegawai_Change
    frmCetakDaftarNoAbsensi.Show
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call txtCariPegawai_Change
    Call subLoadGridSource
End Sub

Sub clear()
    txtIdPegawai.Text = ""
    txtNamaPegawai.Text = ""
    txtNoAbsensi.Text = ""
    txtNamaPegawai.SetFocus
    fraDataPegawai.Visible = False
End Sub

Sub subLoadGridSource()
   Set rs = Nothing
    strSQL = "SELECT DISTINCT " & _
            " dbo.ConvertFingerToDataPegawai.IdPegawai, dbo.ConvertFingerToDataPegawai.PR_FingerID AS NoAbsensi, dbo.DataPegawai.NamaLengkap AS NamaPegawai " & _
            " FROM dbo.DataPegawai INNER JOIN " & _
            " dbo.ConvertFingerToDataPegawai ON dbo.DataPegawai.IdPegawai = dbo.ConvertFingerToDataPegawai.IdPegawai " & mstrFilterData
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgConvertFingerToIDpegawai.DataSource = rs
        dgConvertFingerToIDpegawai.Columns(0).Width = 1200
        dgConvertFingerToIDpegawai.Columns(1).Width = 1000
        dgConvertFingerToIDpegawai.Columns(2).Width = 3900
        dgConvertFingerToIDpegawai.Columns(0).Alignment = vbCenter
        dgConvertFingerToIDpegawai.Columns(1).Alignment = vbCenter
    Set rs = Nothing
End Sub
Private Sub txtCariPegawai_Change()
    mstrFilterData = " WHERE (dbo.DataPegawai.NamaLengkap LIKE '%" & txtCariPegawai.Text & "%') ORDER BY NamaPegawai  "
    Call subLoadGridSource
End Sub

Private Sub txtNamaPegawai_Change()
    fraDataPegawai.Visible = True
    Call subLoadDataPegawai
End Sub

Private Sub txtNamaPegawai_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        If fraDataPegawai.Visible = True Then
            dgDaftarPegawai.SetFocus
        Else
            txtNoAbsensi.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraDataPegawai.Visible = False
    End If
End Sub

Private Sub txtNoAbsensi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsimpan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub
