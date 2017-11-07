VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRegistrasiUser_Offline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Finger Print Pegawai"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "frmRegistrasiUser_Offline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9120
   Begin VB.TextBox txtFilterPeg 
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
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   10
      Top             =   8040
      Width           =   2655
   End
   Begin VB.CheckBox chkUnregFP 
      Caption         =   "Belum Registrasi Finger Print"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2415
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
      Left            =   7560
      TabIndex        =   12
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdBatal 
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
      Left            =   6000
      TabIndex        =   11
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Finger Print"
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
      TabIndex        =   20
      Top             =   1080
      Width           =   8895
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "Ubah ID"
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
         Left            =   4560
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNoPIN 
         Alignment       =   2  'Center
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
         Left            =   120
         MaxLength       =   8
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Format          =   111476737
         UpDown          =   -1  'True
         CurrentDate     =   40106
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Finger Print"
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Daftar"
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
         Index           =   5
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   8895
      Begin VB.TextBox txtNamaPegawai 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtTempatTugas 
         Appearance      =   0  'Flat
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
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtjabatan 
         Appearance      =   0  'Flat
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
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtjk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
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
         Index           =   8
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Bertugas"
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
         Index           =   4
         Left            =   4560
         TabIndex        =   18
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JK"
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
         Index           =   3
         Left            =   4080
         TabIndex        =   17
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
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
         Index           =   2
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID Pegawai"
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
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1110
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
   Begin MSDataGridLib.DataGrid dgFP 
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cari nama pegawai atau ID"
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
      Left            =   120
      TabIndex        =   23
      Top             =   8040
      Width           =   2280
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7200
      Picture         =   "frmRegistrasiUser_Offline.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1995
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRegistrasiUser_Offline.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmRegistrasiUser_Offline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strIDfpLama As String

Private Sub subLoadDataFP()
    If Me.chkUnregFP.Value = 0 Then
        strSQL = "select * from v_PIN"
    Else
        strSQL = "select * from v_PIN where [ID FP] is null"
    End If
    Call msubRecFO(rs, strSQL)
    Set dgFP.DataSource = rs

    With Me.dgFP
        .Columns("FP").Visible = False
        .Columns("Tgl. Mulai").Visible = False
    End With
End Sub

Private Sub subBlankForm()
    txtIDPegawai.Text = ""
    txtNamaPegawai.Text = ""
    txtjk.Text = ""
    txttempattugas.Text = ""
    txtjabatan.Text = ""
    txtNoPIN.Text = ""
    DTPicker1.Value = Now
    txtFilterPeg.Text = ""
End Sub

Private Sub chkUnregFP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call subLoadDataFP
End Sub

Private Sub cmdBatal_Click()
    Call subBlankForm
    Call subLoadDataFP
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    Dim pesan As VbMsgBoxResult
    Dim sqlGantiPIN As String
    Dim cmdGantiPIN As ADODB.Command

    If txtNoPIN.Text = "" Then
        MsgBox "Silahkan Isi ID Finger Pin Terlebih Dahulu.", vbExclamation, "Perhatian"
        Exit Sub
    End If

    If Me.txtIDPegawai.Text = "" Then
        MsgBox "Tidak ada ID Pegawai yang dipilih!" & vbNewLine & _
        "Silahkan pilih ID Pegawai dari daftar yang ada.", vbExclamation, "Perhatian"
        Exit Sub
    End If
    Set rs = Nothing
    strSQL = "select * from PINAbsensiPegawai where PINAbsensi= " & funcPrepareString(Me.txtNoPIN.Text) & ""
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        MsgBox "ID Finger Print sudah terpakai.", vbInformation
        Me.txtNoPIN.Text = ""
        Exit Sub
    End If
    If cmdSimpan.Caption = "Ubah ID" Then
        If Len(Trim(txtNoPIN.Text)) = 0 Then
            pesan = MsgBox("Hapus ID Finger Print untuk " & Me.txtNamaPegawai.Text & "?" _
            , vbInformation Or vbYesNo, "Hapus ID Finger Print")
            If pesan = vbYes Then
                sqlGantiPIN = "DELETE PINAbsensiPegawai" & _
                " WHERE IdPegawai= " & funcPrepareString(Me.txtIDPegawai.Text)
                Set cmdGantiPIN = New ADODB.Command
                With cmdGantiPIN
                    .ActiveConnection = dbConn
                    .CommandText = sqlGantiPIN
                    .CommandType = adCmdText
                    .Execute
                End With
            End If
        Else
            pesan = MsgBox("Ganti ID Finger Print untuk " & Me.txtNamaPegawai.Text & " dari " & strIDfpLama & _
            " ke nomor PIN " & Me.txtNoPIN.Text & "?", vbInformation Or vbYesNo, "Ganti ID Finger Print")
            If pesan = vbYes Then
                sqlGantiPIN = "UPDATE PINAbsensiPegawai SET" & _
                " PINAbsensi=" & funcPrepareString(Me.txtNoPIN.Text) & ", TglDaftar='" & Format(DTPicker1.Value, "yyyy/MM/dd 00:00:00") & "' " & _
                " WHERE IdPegawai= " & funcPrepareString(Me.txtIDPegawai.Text)
                Set cmdGantiPIN = New ADODB.Command
                With cmdGantiPIN
                    .ActiveConnection = dbConn
                    .CommandText = sqlGantiPIN
                    .CommandType = adCmdText
                    .Execute
                End With
            End If
        End If
    Else
        pesan = MsgBox("Tambah PIN untuk " & Me.txtNamaPegawai.Text & " dengan nomor PIN: " & _
        Me.txtNoPIN.Text & "?", vbInformation Or vbYesNo, "Tambah PIN")
        If pesan = vbYes Then
            sqlGantiPIN = "INSERT INTO PINAbsensiPegawai (IdPegawai,PINAbsensi,TglDaftar) VALUES (" & funcPrepareString(Me.txtIDPegawai.Text) & "," & funcPrepareString(Me.txtNoPIN.Text) & ",'" & Format(Me.DTPicker1.Value, "yyyy/mm/dd") & "') "
            Set cmdGantiPIN = New ADODB.Command
            With cmdGantiPIN
                .ActiveConnection = dbConn
                .CommandText = sqlGantiPIN
                .CommandType = adCmdText
                .Execute
            End With
        End If
        MsgBox "ID sudah ditambahkan", vbInformation, "Informasi"
        txtNoPIN.Text = ""
    End If
    Call subLoadDataFP
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgFP_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgFP
    WheelHook.WheelHook dgFP
End Sub

Private Sub dgFP_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo jump
    With dgFP
        txtIDPegawai.Text = .Columns("ID").Text
        txtNamaPegawai.Text = .Columns("Nama").Text
        txtjk.Text = .Columns("JK").Text
        txttempattugas.Text = .Columns("Ruangan").Text
        txtjabatan.Text = .Columns("Jabatan").Text
        txtNoPIN.Text = .Columns("ID FP").Text
        strIDfpLama = .Columns("ID FP").Text

        If Len(Trim(.Columns("ID FP").Text)) = 0 Then
            cmdSimpan.Caption = "Tambah ID"
        Else
            cmdSimpan.Caption = "Ubah ID"
        End If
    End With
jump:
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subLoadDataFP

End Sub

Private Sub txtFilterPeg_Change()
    If Me.chkUnregFP.Value = 0 Then
        strSQL = "select * from v_PIN where Nama like '%" & txtFilterPeg.Text & "%' or ID like '%" & txtFilterPeg.Text & "%'  or [ID FP] like '%" & txtFilterPeg.Text & "%'"
    Else
        strSQL = "select * from v_PIN where [ID FP] is null and Nama like '%" & txtFilterPeg.Text & "%'"
    End If
    Call msubRecFO(rs, strSQL)
    Set dgFP.DataSource = rs

End Sub

Private Sub txtFilterPeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtNoPIN_KeyPress(KeyAscii As Integer)
Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        DTPicker1.SetFocus
    End If
End Sub

