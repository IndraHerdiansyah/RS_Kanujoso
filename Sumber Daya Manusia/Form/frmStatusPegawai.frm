VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmStatusPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Status Pegawai"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatusPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   5775
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
      Begin VB.TextBox txtStatusPegawai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtKdStatus 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   3975
      End
      Begin MSDataGridLib.DataGrid dgStatusPegawai 
         Height          =   4815
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Status"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   7560
      Width           =   1095
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
      Picture         =   "frmStatusPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4200
      Picture         =   "frmStatusPegawai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStatusPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmStatusPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCommand As New ADODB.Command

Private Sub cmdBatal_Click()
    txtStatusPegawai.Enabled = True
    txtStatusPegawai.SetFocus
    clear
    cmdhapus.Enabled = False
    cmdsimpan.Enabled = True
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo xxx
    ' agar gak bisa ngapus
    If InputBox("Administrator Only!!" & vbCr & "Masukan Password Anda!!") <> "admin" Then Exit Sub
    If MsgBox("Hapus Data Master Status?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    dbConn.Execute "DELETE StatusPegawai WHERE KdStatus = '" & txtKdStatus.Text & "'"
    tampilData
    clear
    MsgBox "Data telah dihapus..", vbInformation, "Informasi"
    Exit Sub
xxx:
    MsgBox "Data Tidak Dapat Dihapus..", vbOKOnly, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    With adoCommand
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AUD_StatusPegawai"

        If txtKdStatus.Text = "" Then
            .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, Null)
        Else
            .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, txtKdStatus.Text)
        End If

        .Parameters.Append .CreateParameter("Status", adVarChar, adParamInput, 30, Trim(txtStatusPegawai.Text))
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, "A")

        If Not IsNull(.Parameters("KdStatus").Value) Then
            '' Update
            .Execute
        Else
            '' Input data baru
            .Execute
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    tampilData
    clear
    txtKdStatus.Enabled = False
    txtStatusPegawai.Enabled = False
    cmdBatal_Click
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgStatusPegawai_Click()
    If dgStatusPegawai.ApproxCount = 0 Then Exit Sub
    cmdhapus.Enabled = True
    txtKdStatus.Text = dgStatusPegawai.Columns(0).Value
    txtStatusPegawai.Text = dgStatusPegawai.Columns(1).Value
    txtKdStatus.Enabled = False
    txtStatusPegawai.Enabled = True
End Sub

Private Sub dgStatusPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad
    If dgStatusPegawai.ApproxCount = 0 Then Exit Sub
    txtKdStatus.Text = dgStatusPegawai.Columns(0).Value
    txtStatusPegawai.Text = dgStatusPegawai.Columns(1).Value
    txtKdStatus.Enabled = False
    txtStatusPegawai.Enabled = True
    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    cmdhapus.Enabled = False

    Set rs = Nothing
    strSQL = "select * from StatusPegawai"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgStatusPegawai.DataSource = rs
    dgStatusPegawai.Columns(0).Width = 1000
    dgStatusPegawai.Columns(0).Alignment = vbCenter
    dgStatusPegawai.Columns(0).Caption = "Kode"
    dgStatusPegawai.Columns(1).Width = 3500
    dgStatusPegawai.Columns(1).Caption = "Nama Status"
    Set rs = Nothing
End Sub

Sub clear()
    txtKdStatus.Text = ""
    txtStatusPegawai.Text = ""
End Sub

Sub tampilData()
    Set rs = Nothing
    strSQL = "select * from StatusPegawai"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgStatusPegawai.DataSource = rs
    dgStatusPegawai.Columns(0).Width = 750
    dgStatusPegawai.Columns(0).Alignment = vbCenter
    dgStatusPegawai.Columns(0).Caption = "Kode"
    dgStatusPegawai.Columns(1).Width = 3500
    dgStatusPegawai.Columns(1).Caption = "Nama Status"
    Set rs = Nothing
End Sub

Private Sub txtStatusPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdsimpan.SetFocus
End Sub

Private Sub txtStatusPegawai_LostFocus()
    txtStatusPegawai.Text = StrConv(txtStatusPegawai.Text, vbProperCase)
End Sub
